export function parseScheduleFromOcr(data, imageWidth, imageHeight, createId = defaultCreateId) {
  const words = collectOcrWords(data)
    .map((word) => normaliseWord(word))
    .filter((word) => word.text && Number.isFinite(word.x0) && Number.isFinite(word.y0));

  const metadata = extractMetadata(words, data.text ?? "");
  const rowsFromWords = words.length > 0 ? extractRows(words, imageWidth, imageHeight, createId) : [];
  const rowsFromText = extractRowsFromText(data.text ?? "", createId);
  const rows = rowsFromWords.length >= rowsFromText.length ? rowsFromWords : rowsFromText;

  return {
    metadata: {
      calendarTitle: metadata.patientName
        ? `Royal Rehab - ${metadata.patientName}`
        : "Royal Rehab Schedule",
      patientName: metadata.patientName,
      mrn: metadata.mrn
    },
    rows
  };
}

export function buildSummary(row) {
  const parts = [row.therapist || row.procedure || "Royal Rehab appointment"];
  if (row.location) {
    parts.push(row.location);
  }
  return parts.filter(Boolean).join(" - ");
}

export function dedupeRows(rows) {
  const seen = new Set();
  return rows.filter((row) => {
    const key = [
      row.date,
      row.startTime,
      row.endTime,
      row.therapist,
      row.location,
      row.procedure
    ]
      .map((value) => String(value ?? "").trim().toLowerCase())
      .join("|");

    if (seen.has(key)) {
      return false;
    }

    seen.add(key);
    return true;
  });
}

function collectOcrWords(data) {
  if (Array.isArray(data.words) && data.words.length > 0) {
    return data.words;
  }

  const words = [];
  for (const block of data.blocks ?? []) {
    for (const paragraph of block.paragraphs ?? []) {
      for (const line of paragraph.lines ?? []) {
        for (const word of line.words ?? []) {
          words.push(word);
        }
      }
    }
  }

  return words;
}

function normaliseWord(word) {
  const text = String(word.text ?? "").replace(/\s+/g, " ").trim();
  return {
    text,
    confidence: Number.parseFloat(word.confidence ?? word.conf ?? "0"),
    x0: Number(word.bbox?.x0 ?? word.x0 ?? word.left ?? 0),
    y0: Number(word.bbox?.y0 ?? word.y0 ?? word.top ?? 0),
    x1: Number(word.bbox?.x1 ?? word.x1 ?? word.right ?? 0),
    y1: Number(word.bbox?.y1 ?? word.y1 ?? word.bottom ?? 0)
  };
}

function extractMetadata(words, text) {
  const normalisedText = text.replace(/\s+/g, " ");
  const patientMatch = normalisedText.match(/Patient\s+Name[:\s]+(.+?)\s+MRN[:\s]/i);
  const mrnMatch = normalisedText.match(/MRN[:\s]+([A-Z0-9]+)/i);

  if (patientMatch || mrnMatch) {
    return {
      patientName: patientMatch?.[1]?.trim() ?? "",
      mrn: mrnMatch?.[1]?.trim() ?? ""
    };
  }

  const patientWords = words
    .filter((word) => word.y0 >= 0 && word.y0 <= 220)
    .sort(sortByYThenX);
  const patientLine = patientWords.map((word) => word.text).join(" ");
  const fallbackPatientMatch = patientLine.match(/Patient\s*Name[:\s]+(.+?)\s+MRN[:\s]+/i);

  return {
    patientName: fallbackPatientMatch?.[1]?.trim() ?? "",
    mrn: mrnMatch?.[1]?.trim() ?? ""
  };
}

function extractRows(words, imageWidth, imageHeight, createId) {
  const headerY = findHeaderY(words);
  const footerY = findFooterY(words, imageHeight);
  const columnEdges = detectColumnEdges(words, imageWidth, headerY);

  const candidateWords = words.filter((word) => {
    if (word.y0 <= headerY + 6 || word.y0 >= footerY) {
      return false;
    }

    if (/^(Patient|Name|MRN)$/i.test(word.text)) {
      return false;
    }

    return true;
  });

  const groupedRows = groupWordsIntoRows(candidateWords);
  const parsedRows = [];

  for (const rowWords of groupedRows) {
    const row = parseRow(rowWords, columnEdges, createId);
    if (row) {
      parsedRows.push(row);
    }
  }

  return parsedRows;
}

function extractRowsFromText(text, createId) {
  const lines = text
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean);
  const startIndex = lines.findIndex((line) => /booking\s+date/i.test(line));
  const relevantLines = startIndex >= 0 ? lines.slice(startIndex + 1) : lines;
  const rows = [];

  for (const line of relevantLines) {
    if (/printed|ghaz|report|patient diary/i.test(line)) {
      continue;
    }

    const row = parseTextLine(line, createId);
    if (row) {
      rows.push(row);
    }
  }

  return rows;
}

function findHeaderY(words) {
  const headerWords = words.filter((word) => /book|date|start|stan|therap|facility|location|procedure/i.test(word.text));
  if (headerWords.length === 0) {
    return 180;
  }

  return Math.max(...headerWords.map((word) => word.y1));
}

function findFooterY(words, imageHeight) {
  const footerWords = words.filter((word) => /printed|ghaz|report/i.test(word.text));
  if (footerWords.length === 0) {
    return imageHeight - 80;
  }

  return Math.min(...footerWords.map((word) => word.y0)) - 6;
}

function detectColumnEdges(words, imageWidth, headerY) {
  const headerBand = words.filter((word) => word.y0 >= headerY - 28 && word.y0 <= headerY + 16);
  const bookingX = findWordX(headerBand, /book|date/i) ?? imageWidth * 0.09;
  const startX = findWordX(headerBand, /start|stan/i) ?? imageWidth * 0.19;
  const endX = findWordX(headerBand, /^endtime$|^end$|ime/i) ?? imageWidth * 0.28;
  const therapistX = findWordX(headerBand, /therap/i) ?? imageWidth * 0.38;
  const facilityX = findWordX(headerBand, /facility|location/i) ?? imageWidth * 0.66;
  const procedureX = findWordX(headerBand, /procedure/i) ?? imageWidth * 0.86;

  return {
    dateEnd: midpoint(bookingX, startX),
    startEnd: midpoint(startX, endX),
    endEnd: midpoint(endX, therapistX),
    therapistEnd: midpoint(therapistX, facilityX),
    locationEnd: midpoint(facilityX, procedureX)
  };
}

function findWordX(words, matcher) {
  const match = words.find((word) => matcher.test(word.text));
  return match ? match.x0 : null;
}

function midpoint(left, right) {
  return left + (right - left) / 2;
}

function groupWordsIntoRows(words) {
  const sortedWords = [...words].sort(sortByYThenX);
  const groups = [];
  const averageHeight =
    sortedWords.reduce((sum, word) => sum + Math.max(8, word.y1 - word.y0), 0) / Math.max(sortedWords.length, 1);
  const tolerance = Math.max(12, averageHeight * 0.9);

  for (const word of sortedWords) {
    const centerY = (word.y0 + word.y1) / 2;
    const currentGroup = groups.at(-1);

    if (!currentGroup || Math.abs(currentGroup.centerY - centerY) > tolerance) {
      groups.push({ centerY, words: [word] });
      continue;
    }

    currentGroup.words.push(word);
    currentGroup.centerY = (currentGroup.centerY * (currentGroup.words.length - 1) + centerY) / currentGroup.words.length;
  }

  return groups.map((group) => group.words.sort((left, right) => left.x0 - right.x0));
}

function parseRow(words, edges, createId) {
  const columns = {
    date: [],
    start: [],
    end: [],
    therapist: [],
    location: [],
    procedure: []
  };

  for (const word of words) {
    if (word.x0 < edges.dateEnd) {
      columns.date.push(word);
    } else if (word.x0 < edges.startEnd) {
      columns.start.push(word);
    } else if (word.x0 < edges.endEnd) {
      columns.end.push(word);
    } else if (word.x0 < edges.therapistEnd) {
      columns.therapist.push(word);
    } else if (word.x0 < edges.locationEnd) {
      columns.location.push(word);
    } else {
      columns.procedure.push(word);
    }
  }

  const date = parseDate(joinWords(columns.date));
  if (!date) {
    return null;
  }

  const startTime = parseTime(joinWords(columns.start));
  const endTime = parseTime(joinWords(columns.end));
  if (!startTime || !endTime) {
    return null;
  }

  const therapist = cleanPhrase(joinWords(columns.therapist), "therapist");
  const location = cleanPhrase(joinWords(columns.location), "location");
  const procedure = cleanPhrase(joinWords(columns.procedure), "procedure");

  const row = {
    id: createId(),
    date,
    startTime,
    endTime,
    therapist,
    location,
    procedure,
    selected: true,
    summary: "",
    generatedSummary: ""
  };

  row.summary = buildSummary(row);
  row.generatedSummary = row.summary;
  return row;
}

function parseTextLine(line, createId) {
  const normalizedLine = line.replace(/\s+/g, " ").trim();
  const lineMatch = normalizedLine.match(
    /^(Mon|Tue|Wed|Thu|Fri|Sat|Sun)\s+([0-9\/.-]+)\s+([0-9:.]{4,5})\s+([0-9:.]{4,5})\s+(.+)$/i
  );
  if (!lineMatch) {
    return null;
  }

  const [, , rawDate, rawStart, rawEnd, remainder] = lineMatch;
  const date = parseDate(rawDate);
  if (!date) {
    return null;
  }

  const tokens = remainder.split(/\s+/);
  const locationIndex = tokens.findIndex((token) => /\d/.test(token));
  const therapistTokens = locationIndex >= 0 ? tokens.slice(0, locationIndex) : tokens;
  const locationTokens = locationIndex >= 0 ? tokens.slice(locationIndex) : [];

  const row = {
    id: createId(),
    date,
    startTime: parseTime(rawStart),
    endTime: parseTime(rawEnd),
    therapist: cleanPhrase(therapistTokens.join(" "), "therapist"),
    location: cleanPhrase(locationTokens.join(" "), "location"),
    procedure: "",
    selected: true,
    summary: "",
    generatedSummary: ""
  };

  if (!row.startTime || !row.endTime) {
    return null;
  }

  row.summary = buildSummary(row);
  row.generatedSummary = row.summary;
  return row;
}

function joinWords(words) {
  return words
    .map((word) => word.text)
    .filter(Boolean)
    .join(" ")
    .replace(/\s+/g, " ")
    .trim();
}

function parseDate(text) {
  const match = text.match(/(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{2,4})/);
  if (match) {
    const [, day, month, year] = match;
    const fullYear = year.length === 2 ? `20${year}` : year;
    return `${fullYear.padStart(4, "0")}-${month.padStart(2, "0")}-${day.padStart(2, "0")}`;
  }

  return parseDigitOnlyDate(text) ?? "";
}

function parseDigitOnlyDate(text) {
  const digits = (text.match(/\d/g) ?? []).join("");
  if (digits.length < 8) {
    return null;
  }

  const yearMatch = digits.match(/(20\d{2})$/);
  const year = yearMatch?.[1] ?? digits.slice(-4);
  const day = digits.slice(0, 2);
  const middle = digits.slice(2, digits.length - 4);
  const month = guessMonthFromMiddleDigits(middle);

  if (!isValidDateParts(day, month, year)) {
    return null;
  }

  return `${year.padStart(4, "0")}-${month.padStart(2, "0")}-${day.padStart(2, "0")}`;
}

function guessMonthFromMiddleDigits(middle) {
  if (!middle) {
    return "";
  }

  if (middle.length === 1) {
    return `0${middle}`;
  }

  const candidates = [];
  for (let index = 0; index <= middle.length - 2; index += 1) {
    const candidate = middle.slice(index, index + 2);
    const value = Number.parseInt(candidate, 10);
    if (value >= 1 && value <= 12) {
      candidates.push(candidate);
    }
  }

  return candidates.at(-1) ?? middle.slice(0, 2);
}

function isValidDateParts(day, month, year) {
  const dayValue = Number.parseInt(day, 10);
  const monthValue = Number.parseInt(month, 10);
  const yearValue = Number.parseInt(year, 10);

  return Number.isInteger(dayValue) &&
    Number.isInteger(monthValue) &&
    Number.isInteger(yearValue) &&
    dayValue >= 1 &&
    dayValue <= 31 &&
    monthValue >= 1 &&
    monthValue <= 12 &&
    yearValue >= 2000 &&
    yearValue <= 2100;
}

function parseTime(text) {
  const compact = text
    .toLowerCase()
    .replace(/[oO]/g, "0")
    .replace(/\s+/g, " ")
    .trim();

  const directMatch = compact.match(/(\d{1,2})[:.](\d{2})/);
  if (directMatch) {
    return `${directMatch[1].padStart(2, "0")}:${directMatch[2]}`;
  }

  const digits = compact.match(/\d+/g) ?? [];
  if (digits.length >= 2) {
    return `${digits[0].padStart(2, "0")}:${digits[1].padStart(2, "0")}`;
  }

  return "";
}

function cleanPhrase(text, type) {
  let value = text
    .replace(/\b0T\b/g, "OT")
    .replace(/\b040[7T),.]?\b/gi, "04 OT")
    .replace(/\b050[7T),.]?\b/gi, "05 OT")
    .replace(/\bPHYSI0\b/g, "PHYSIO")
    .replace(/\bPHYSI01\b/g, "PHYSIO1")
    .replace(/\bPHYSIOI\b/g, "PHYSIO1")
    .replace(/\bPHYSION\b/g, "PHYSIO1")
    .replace(/\bPHYSIOA\b/g, "PHYSIO1")
    .replace(/\b02PHYSIO\b/g, "02 PHYSIO")
    .replace(/\bEL\)/g, "(BL)")
    .replace(/\bog\b/g, "09")
    .replace(/\bTHERAPY occupational\b/i, "THERAPY, Occupational")
    .replace(/\s+,/g, ",")
    .replace(/\(\s+/g, "(")
    .replace(/\s+\)/g, ")")
    .replace(/\s+/g, " ")
    .trim();

  if (type === "location") {
    value = value
      .replace(/\bEX PHYS A \((BL|MG)\)\b/gi, (_, suffix) => `EX PHYS A (${suffix.toUpperCase()})`)
      .replace(/\b(\d{2})([A-Z])/g, "$1 $2")
      .replace(/[.,;:]+$/g, "");
    if (value.endsWith(")") && !value.includes("(")) {
      value = value.slice(0, -1);
    }
  }

  if (type === "therapist") {
    value = value
      .replace(/^THERAPY,\s*THERAPY,\s*/i, "THERAPY, ")
      .replace(/^THERAPY\s+Occupational$/i, "THERAPY, Occupational")
      .replace(/^THERAPY\s+Exercise$/i, "THERAPY, Exercise")
      .replace(/^Occupational$/i, "THERAPY, Occupational")
      .replace(/^Exercise$/i, "THERAPY, Exercise");
  }

  return value;
}

function defaultCreateId() {
  if (globalThis.crypto?.randomUUID) {
    return globalThis.crypto.randomUUID();
  }

  return `row-${Math.random().toString(16).slice(2)}-${Date.now().toString(16)}`;
}

function sortByYThenX(left, right) {
  return left.y0 === right.y0 ? left.x0 - right.x0 : left.y0 - right.y0;
}
