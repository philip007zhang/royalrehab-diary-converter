const state = {
  sourceFile: null,
  sourceName: "",
  previewUrl: "",
  sourceKind: "",
  sourcePageCount: 0,
  rows: [],
  ocrText: "",
  googleCalendar: {
    clientId: "",
    enabled: false,
    ready: false,
    connected: false,
    tokenClient: null,
    accessToken: "",
    tokenExpiresAt: 0
  },
  metadata: {
    calendarTitle: "Royal Rehab Schedule",
    timezone: "Australia/Sydney",
    patientName: "",
    mrn: ""
  }
};

const elements = {
  addRowButton: document.querySelector("#add-row-btn"),
  addCalendarButton: document.querySelector("#add-calendar-btn"),
  calendarProvider: document.querySelector("#calendar-provider"),
  calendarTitle: document.querySelector("#calendar-title"),
  timezone: document.querySelector("#calendar-timezone"),
  patientName: document.querySelector("#patient-name"),
  patientMrn: document.querySelector("#patient-mrn"),
  downloadCsvButton: document.querySelector("#download-csv-btn"),
  downloadIcsButton: document.querySelector("#download-ics-btn"),
  downloadJsonButton: document.querySelector("#download-json-btn"),
  dropzone: document.querySelector("#dropzone"),
  dropzonePreview: document.querySelector("#dropzone-preview"),
  dropzoneSubtitle: document.querySelector("#dropzone-subtitle"),
  dropzoneTitle: document.querySelector("#dropzone-title"),
  extractButton: document.querySelector("#extract-btn"),
  fileInput: document.querySelector("#file-input"),
  filePill: document.querySelector("#file-pill"),
  firstAppointment: document.querySelector("#first-appointment"),
  googleConnectButton: document.querySelector("#google-connect-btn"),
  googleStatusPill: document.querySelector("#google-status-pill"),
  lastAppointment: document.querySelector("#last-appointment"),
  ocrText: document.querySelector("#ocr-text"),
  previewImage: document.querySelector("#preview-image"),
  rowsBody: document.querySelector("#rows-body"),
  selectedCountLabel: document.querySelector("#selected-count-label"),
  selectedList: document.querySelector("#selected-list"),
  selectedProviderHint: document.querySelector("#selected-provider-hint"),
  selectAllCheckbox: document.querySelector("#select-all-checkbox"),
  sessionsCount: document.querySelector("#sessions-count"),
  statusPanel: document.querySelector("#status-panel"),
  statusText: document.querySelector("#status-text"),
  statusTitle: document.querySelector("#status-title"),
  useSampleButton: document.querySelector("#use-sample-btn")
};

const dateFormatter = new Intl.DateTimeFormat("en-AU", {
  day: "2-digit",
  month: "short",
  year: "numeric"
});

bindEvents();
render();
void initGoogleCalendarSupport();

async function logAuditEvent(activity, status, details = {}) {
  try {
    await fetch("/api/audit-events", {
      method: "POST",
      headers: {
        "Content-Type": "application/json"
      },
      body: JSON.stringify({ activity, status, details })
    });
  } catch (error) {
    console.error("Audit logging failed", error);
  }
}

function generateClientId() {
  if (globalThis.crypto?.randomUUID) {
    return globalThis.crypto.randomUUID();
  }

  const bytes = new Uint8Array(16);
  if (globalThis.crypto?.getRandomValues) {
    globalThis.crypto.getRandomValues(bytes);
  } else {
    for (let index = 0; index < bytes.length; index += 1) {
      bytes[index] = Math.floor(Math.random() * 256);
    }
  }

  bytes[6] = (bytes[6] & 0x0f) | 0x40;
  bytes[8] = (bytes[8] & 0x3f) | 0x80;

  const hex = Array.from(bytes, (value) => value.toString(16).padStart(2, "0")).join("");
  return `${hex.slice(0, 8)}-${hex.slice(8, 12)}-${hex.slice(12, 16)}-${hex.slice(16, 20)}-${hex.slice(20)}`;
}

function bindEvents() {
  elements.fileInput.addEventListener("change", async (event) => {
    const [file] = event.target.files ?? [];
    if (file) {
      await loadFile(file);
    }
  });

  elements.useSampleButton.addEventListener("click", useBundledSample);
  elements.extractButton.addEventListener("click", extractSchedule);
  elements.downloadIcsButton.addEventListener("click", downloadIcs);
  elements.downloadCsvButton.addEventListener("click", downloadCsv);
  elements.downloadJsonButton.addEventListener("click", downloadJson);
  elements.addRowButton.addEventListener("click", addEmptyRow);
  elements.addCalendarButton.addEventListener("click", addSelectedToCalendar);
  elements.googleConnectButton.addEventListener("click", connectGoogleCalendar);

  elements.calendarTitle.addEventListener("input", (event) => {
    state.metadata.calendarTitle = event.target.value.trim() || "Royal Rehab Schedule";
    render();
  });

  elements.timezone.addEventListener("input", (event) => {
    state.metadata.timezone = event.target.value.trim() || "Australia/Sydney";
  });

  elements.patientName.addEventListener("input", (event) => {
    state.metadata.patientName = event.target.value.trim();
  });

  elements.patientMrn.addEventListener("input", (event) => {
    state.metadata.mrn = event.target.value.trim();
  });

  elements.calendarProvider.addEventListener("change", () => {
    syncCalendarActionLabel();
  });

  elements.selectAllCheckbox.addEventListener("change", (event) => {
    const checked = Boolean(event.target.checked);
    state.rows = state.rows.map((row) => ({ ...row, selected: checked }));
    render(false);
  });

  elements.rowsBody.addEventListener("input", (event) => {
    const target = event.target;
    if (!(target instanceof HTMLInputElement)) {
      return;
    }

    const rowId = target.dataset.rowId;
    const field = target.dataset.field;
    if (!rowId || !field) {
      return;
    }

    const row = state.rows.find((item) => item.id === rowId);
    if (!row) {
      return;
    }

    row[field] = target.value.trim();
    if (field !== "summary" && shouldAutoRefreshSummary(row)) {
      row.summary = buildSummary(row);
      row.generatedSummary = row.summary;
    }
    render(false);
  });

  elements.rowsBody.addEventListener("change", (event) => {
    const target = event.target;
    if (!(target instanceof HTMLInputElement) || target.type !== "checkbox") {
      return;
    }

    const rowId = target.dataset.selectRow;
    if (!rowId) {
      return;
    }

    const row = state.rows.find((item) => item.id === rowId);
    if (!row) {
      return;
    }

    row.selected = target.checked;
    render(false);
  });

  elements.rowsBody.addEventListener("click", (event) => {
    const button = event.target;
    if (!(button instanceof HTMLButtonElement)) {
      return;
    }

    const rowId = button.dataset.removeId;
    if (!rowId) {
      return;
    }

    state.rows = state.rows.filter((row) => row.id !== rowId);
    render();
  });

  for (const eventName of ["dragenter", "dragover"]) {
    elements.dropzone.addEventListener(eventName, (event) => {
      event.preventDefault();
      elements.dropzone.classList.add("dragging");
    });
  }

  for (const eventName of ["dragleave", "drop"]) {
    elements.dropzone.addEventListener(eventName, (event) => {
      event.preventDefault();
      elements.dropzone.classList.remove("dragging");
    });
  }

  elements.dropzone.addEventListener("drop", async (event) => {
    const [file] = event.dataTransfer?.files ?? [];
    if (file) {
      await loadFile(file);
    }
  });
}

async function initGoogleCalendarSupport() {
  try {
    const response = await fetch("/api/google-config", { cache: "no-store" });
    if (!response.ok) {
      throw new Error(`Google config request failed with status ${response.status}`);
    }

    const config = await response.json();
    state.googleCalendar.enabled = Boolean(config.enabled);
    state.googleCalendar.clientId = String(config.clientId ?? "");

    if (!state.googleCalendar.enabled) {
      syncGoogleCalendarUi();
      return;
    }

    await waitForGlobal("google");
    initGoogleCalendarClient();
  } catch (error) {
    console.error(error);
  } finally {
    syncGoogleCalendarUi();
    syncCalendarActionLabel();
  }
}

function waitForGlobal(name, timeoutMs = 15000) {
  return new Promise((resolve, reject) => {
    const startedAt = Date.now();

    function check() {
      if (globalThis[name]) {
        resolve(globalThis[name]);
        return;
      }

      if (Date.now() - startedAt >= timeoutMs) {
        reject(new Error(`${name} did not load in time`));
        return;
      }

      window.setTimeout(check, 100);
    }

    check();
  });
}

function initGoogleCalendarClient() {
  state.googleCalendar.tokenClient = window.google.accounts.oauth2.initTokenClient({
    client_id: state.googleCalendar.clientId,
    scope: "https://www.googleapis.com/auth/calendar.events",
    callback: () => {}
  });
  state.googleCalendar.ready = true;
}

async function useBundledSample() {
  setStatus("processing", "Loading sample", "Pulling the bundled Royal Rehab example into the app.");
  const response = await fetch("/sample/Sample-schedule.png", { cache: "no-store" });
  const blob = await response.blob();
  const file = new File([blob], "Sample-schedule.png", { type: blob.type || "image/png" });
  await loadFile(file);
  setStatus("idle", "Sample loaded", "The bundled sample is ready. Click Extract schedule to run OCR.");
}

async function loadFile(file) {
  if (!isSupportedSourceFile(file)) {
    setStatus("error", "Unsupported file", "Use a PNG, JPG, WEBP, BMP, or PDF diary export.");
    return;
  }

  const isPdf = isPdfFile(file);
  state.sourceFile = file;
  state.sourceName = file.name.replace(/\.[^.]+$/, "");
  state.sourceKind = isPdf ? "pdf" : "image";
  state.sourcePageCount = 1;
  state.rows = [];
  state.ocrText = "";

  if (state.previewUrl) {
    URL.revokeObjectURL(state.previewUrl);
  }

  let previewUrl = "";
  let pageCount = 1;

  try {
    if (isPdf) {
      const pdfPreview = await buildPdfPreview(file);
      previewUrl = pdfPreview.previewUrl;
      pageCount = pdfPreview.pageCount;
    } else {
      previewUrl = URL.createObjectURL(file);
    }
  } catch (error) {
    console.error(error);
    setStatus(
      "error",
      "Preview failed",
      error instanceof Error ? error.message : "The file preview could not be prepared."
    );
    return;
  }

  state.previewUrl = previewUrl;
  state.sourcePageCount = pageCount;
  elements.previewImage.src = state.previewUrl;
  elements.dropzonePreview.classList.add("has-image");
  elements.filePill.textContent = isPdf && pageCount > 1 ? `${file.name} (${pageCount} pages)` : file.name;
  elements.ocrText.textContent = "No OCR data yet.";
  elements.dropzone.classList.add("loaded");
  elements.dropzoneTitle.textContent = `${isPdf ? "PDF" : "Image"} uploaded successfully`;
  elements.dropzoneSubtitle.textContent = isPdf && pageCount > 1 ? `${file.name} · ${pageCount} pages` : file.name;
  setStatus(
    "idle",
    "File loaded",
    isPdf
      ? `The PDF is ready. OCR will scan ${pageCount} page${pageCount === 1 ? "" : "s"} when you start extraction.`
      : "The image is ready. Start extraction when you're ready."
  );
  void logAuditEvent("image_loaded", "success", {
    fileName: file.name,
    fileType: file.type,
    sizeBytes: file.size,
    sourceKind: state.sourceKind,
    pageCount: state.sourcePageCount
  });
  render();
}

async function extractSchedule() {
  if (!state.sourceFile) {
    setStatus("error", "No file loaded", "Choose a schedule image or PDF first.");
    return;
  }

  if (!window.Tesseract) {
    setStatus("error", "OCR library unavailable", "Tesseract.js did not load. Refresh the page and try again.");
    return;
  }

  elements.extractButton.disabled = true;
  setStatus("processing", "Preparing image", "Enhancing the schedule image for OCR.");

  let worker;

  try {
    worker = await window.Tesseract.createWorker("eng", 1, {
      logger: (message) => {
        if (message.status && typeof message.progress === "number") {
          const percentage = Math.round(message.progress * 100);
          setStatus(
            "processing",
            "Running OCR",
            `${capitalize(message.status)}... ${Number.isFinite(percentage) ? `${percentage}%` : ""}`.trim()
          );
        }
      }
    });

    const extraction = state.sourceKind === "pdf"
      ? await extractFromPdf(worker, state.sourceFile)
      : await extractFromImage(worker, state.sourceFile);

    state.ocrText = extraction.ocrText || "No OCR text returned.";
    elements.ocrText.textContent = state.ocrText;
    state.rows = extraction.rows.map((row) => ({
      ...row,
      generatedSummary: row.summary,
      selected: row.selected ?? true
    }));

    state.metadata = {
      ...state.metadata,
      ...extraction.metadata,
      calendarTitle: extraction.metadata.calendarTitle || state.metadata.calendarTitle || "Royal Rehab Schedule",
      timezone: state.metadata.timezone || "Australia/Sydney"
    };

    syncMetadataInputs();

    if (state.rows.length === 0) {
      void logAuditEvent("ocr_extraction", "warning", {
        fileName: state.sourceFile?.name ?? "",
        rowsFound: 0,
        sourceKind: state.sourceKind,
        pageCount: state.sourcePageCount
      });
      setStatus(
        "error",
        "No sessions found",
        "OCR ran, but no appointment rows were confidently parsed. You can still add rows manually."
      );
    } else {
      void logAuditEvent("ocr_extraction", "success", {
        fileName: state.sourceFile?.name ?? "",
        rowsFound: state.rows.length,
        patientName: state.metadata.patientName,
        sourceKind: state.sourceKind,
        pageCount: state.sourcePageCount
      });
      setStatus(
        "success",
        "Extraction complete",
        `${state.rows.length} appointment${state.rows.length === 1 ? "" : "s"} extracted. Review the rows before download.`
      );
    }
  } catch (error) {
    console.error(error);
    void logAuditEvent("ocr_extraction", "error", {
      fileName: state.sourceFile?.name ?? "",
      message: error instanceof Error ? error.message : "Unexpected OCR error",
      sourceKind: state.sourceKind,
      pageCount: state.sourcePageCount
    });
    setStatus("error", "Extraction failed", error instanceof Error ? error.message : "Unexpected OCR error.");
  } finally {
    if (worker) {
      await worker.terminate().catch(() => {});
    }
    elements.extractButton.disabled = false;
    render();
  }
}

function isSupportedSourceFile(file) {
  return file.type.startsWith("image/") || isPdfFile(file);
}

function isPdfFile(file) {
  return file.type === "application/pdf" || /\.pdf$/i.test(file.name);
}

async function buildPdfPreview(file) {
  const pdfDocument = await loadPdfDocument(file);
  try {
    const firstPage = await pdfDocument.getPage(1);
    const previewCanvas = await renderPdfPageToCanvas(firstPage, 1100);
    return {
      previewUrl: await canvasToObjectUrl(previewCanvas),
      pageCount: pdfDocument.numPages
    };
  } finally {
    await pdfDocument.destroy();
  }
}

async function extractFromImage(worker, file) {
  const processedCanvas = await preprocessImageFile(file);
  const result = await worker.recognize(processedCanvas, {}, { text: true, blocks: true });
  const parsed = parseScheduleFromOcr(result.data, processedCanvas.width, processedCanvas.height);

  return {
    metadata: parsed.metadata,
    ocrText: result.data.text.trim(),
    rows: parsed.rows
  };
}

async function extractFromPdf(worker, file) {
  const pdfDocument = await loadPdfDocument(file);
  const rows = [];
  const ocrParts = [];
  let metadata = {};

  try {
    for (let pageNumber = 1; pageNumber <= pdfDocument.numPages; pageNumber += 1) {
      setStatus(
        "processing",
        "Preparing PDF",
        `Rendering page ${pageNumber} of ${pdfDocument.numPages} for OCR.`
      );

      const page = await pdfDocument.getPage(pageNumber);
      const renderedCanvas = await renderPdfPageToCanvas(page, 1800);
      const processedCanvas = await preprocessCanvas(renderedCanvas);
      const result = await worker.recognize(processedCanvas, {}, { text: true, blocks: true });
      const parsed = parseScheduleFromOcr(result.data, processedCanvas.width, processedCanvas.height);

      if (result.data.text?.trim()) {
        ocrParts.push(
          pdfDocument.numPages > 1
            ? `--- Page ${pageNumber} ---\n${result.data.text.trim()}`
            : result.data.text.trim()
        );
      }

      if (!metadata.patientName && parsed.metadata.patientName) {
        metadata = { ...metadata, patientName: parsed.metadata.patientName };
      }
      if (!metadata.mrn && parsed.metadata.mrn) {
        metadata = { ...metadata, mrn: parsed.metadata.mrn };
      }
      if (!metadata.calendarTitle && parsed.metadata.calendarTitle) {
        metadata = { ...metadata, calendarTitle: parsed.metadata.calendarTitle };
      }

      rows.push(...parsed.rows);
    }
  } finally {
    await pdfDocument.destroy();
  }

  return {
    metadata,
    ocrText: ocrParts.join("\n\n").trim(),
    rows: dedupeRows(rows)
  };
}

async function loadPdfDocument(file) {
  const pdfjs = await ensurePdfJsReady();
  const bytes = new Uint8Array(await file.arrayBuffer());
  return pdfjs.getDocument({ data: bytes }).promise;
}

async function ensurePdfJsReady() {
  const pdfjs = await waitForGlobal("pdfjsLib");
  if (!pdfjs.GlobalWorkerOptions.workerSrc) {
    pdfjs.GlobalWorkerOptions.workerSrc = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";
  }
  return pdfjs;
}

async function renderPdfPageToCanvas(page, targetWidth) {
  const initialViewport = page.getViewport({ scale: 1 });
  const scale = Math.max(1.25, targetWidth / Math.max(initialViewport.width, 1));
  const viewport = page.getViewport({ scale });
  const canvas = document.createElement("canvas");
  canvas.width = Math.ceil(viewport.width);
  canvas.height = Math.ceil(viewport.height);
  const context = canvas.getContext("2d", { willReadFrequently: true });

  context.fillStyle = "#ffffff";
  context.fillRect(0, 0, canvas.width, canvas.height);
  await page.render({ canvasContext: context, viewport }).promise;
  return canvas;
}

async function canvasToObjectUrl(canvas) {
  const blob = await new Promise((resolve, reject) => {
    canvas.toBlob((value) => {
      if (value) {
        resolve(value);
      } else {
        reject(new Error("Canvas preview could not be created."));
      }
    }, "image/png");
  });

  return URL.createObjectURL(blob);
}

async function preprocessImageFile(file) {
  const bitmap = await createImageBitmap(file);
  try {
    return await preprocessDrawable(bitmap, bitmap.width, bitmap.height);
  } finally {
    bitmap.close();
  }
}

async function preprocessCanvas(sourceCanvas) {
  return preprocessDrawable(sourceCanvas, sourceCanvas.width, sourceCanvas.height);
}

async function preprocessDrawable(source, sourceWidth, sourceHeight) {
  const scale = Math.max(1.6, Math.min(2.1, 1900 / sourceWidth));
  const canvas = document.createElement("canvas");
  canvas.width = Math.round(sourceWidth * scale);
  canvas.height = Math.round(sourceHeight * scale);
  const context = canvas.getContext("2d", { willReadFrequently: true });

  context.fillStyle = "#ffffff";
  context.fillRect(0, 0, canvas.width, canvas.height);
  context.filter = "grayscale(1) contrast(1.8) brightness(1.12)";
  context.drawImage(source, 0, 0, canvas.width, canvas.height);
  context.filter = "none";

  const imageData = context.getImageData(0, 0, canvas.width, canvas.height);
  const { data } = imageData;

  for (let index = 0; index < data.length; index += 4) {
    const luminance = data[index] * 0.299 + data[index + 1] * 0.587 + data[index + 2] * 0.114;
    const boosted = luminance > 238 ? 255 : luminance < 150 ? 0 : luminance;
    data[index] = boosted;
    data[index + 1] = boosted;
    data[index + 2] = boosted;
    data[index + 3] = 255;
  }

  context.putImageData(imageData, 0, 0);
  return canvas;
}

function parseScheduleFromOcr(data, imageWidth, imageHeight) {
  const words = collectOcrWords(data)
    .map((word) => normaliseWord(word))
    .filter((word) => word.text && Number.isFinite(word.x0) && Number.isFinite(word.y0));

  const metadata = extractMetadata(words, data.text ?? "");
  const rowsFromWords = words.length > 0 ? extractRows(words, imageWidth, imageHeight) : [];
  const rowsFromText = extractRowsFromText(data.text ?? "");
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

function extractRows(words, imageWidth, imageHeight) {
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
    const row = parseRow(rowWords, columnEdges);
    if (row) {
      parsedRows.push(row);
    }
  }

  return parsedRows;
}

function extractRowsFromText(text) {
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

    const row = parseTextLine(line);
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

function parseRow(words, edges) {
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
    id: generateClientId(),
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

function parseTextLine(line) {
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
    id: generateClientId(),
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

function buildSummary(row) {
  const parts = [row.therapist || row.procedure || "Royal Rehab appointment"];
  if (row.location) {
    parts.push(row.location);
  }
  return parts.filter(Boolean).join(" - ");
}

function shouldAutoRefreshSummary(row) {
  return !row.summary || row.summary === row.generatedSummary;
}

function dedupeRows(rows) {
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

function addEmptyRow() {
  const row = {
    id: generateClientId(),
    date: "",
    startTime: "",
    endTime: "",
    therapist: "",
    location: "",
    procedure: "",
    selected: true,
    summary: "",
    generatedSummary: ""
  };

  state.rows = [...state.rows, row];
  render();
}

function render(full = true) {
  renderTable();
  renderStats();
  renderSelectedBatch();
  toggleDownloadButtons();
  syncSelectionState();
  syncCalendarActionLabel();
  if (full) {
    syncMetadataInputs();
  }
}

function syncMetadataInputs() {
  elements.calendarTitle.value = state.metadata.calendarTitle;
  elements.timezone.value = state.metadata.timezone;
  elements.patientName.value = state.metadata.patientName;
  elements.patientMrn.value = state.metadata.mrn;
}

function renderTable() {
  if (state.rows.length === 0) {
    elements.rowsBody.innerHTML = `
      <tr class="empty-row">
        <td colspan="9">No sessions extracted yet.</td>
      </tr>
    `;
    return;
  }

  elements.rowsBody.innerHTML = state.rows
    .map(
      (row) => `
        <tr>
          <td class="row-select-cell">
            <input
              class="row-select-checkbox"
              type="checkbox"
              ${row.selected !== false ? "checked" : ""}
              data-select-row="${row.id}"
              aria-label="Select appointment"
            />
          </td>
          <td><input type="date" value="${escapeHtml(row.date)}" data-row-id="${row.id}" data-field="date" /></td>
          <td><input type="time" value="${escapeHtml(row.startTime)}" data-row-id="${row.id}" data-field="startTime" /></td>
          <td><input type="time" value="${escapeHtml(row.endTime)}" data-row-id="${row.id}" data-field="endTime" /></td>
          <td><input type="text" value="${escapeHtml(row.therapist)}" data-row-id="${row.id}" data-field="therapist" /></td>
          <td><input type="text" value="${escapeHtml(row.location)}" data-row-id="${row.id}" data-field="location" /></td>
          <td><input type="text" value="${escapeHtml(row.procedure)}" data-row-id="${row.id}" data-field="procedure" /></td>
          <td><input type="text" value="${escapeHtml(row.summary)}" data-row-id="${row.id}" data-field="summary" /></td>
          <td><button class="remove-row-button" type="button" data-remove-id="${row.id}">Remove</button></td>
        </tr>
      `
    )
    .join("");
}

function renderStats() {
  const validRows = getSortedValidRows();
  elements.sessionsCount.textContent = String(validRows.length);
  elements.firstAppointment.textContent = validRows[0] ? formatAppointmentLabel(validRows[0]) : "-";
  elements.lastAppointment.textContent = validRows.at(-1) ? formatAppointmentLabel(validRows.at(-1)) : "-";
}

function formatAppointmentLabel(row) {
  const date = new Date(`${row.date}T00:00:00`);
  return `${dateFormatter.format(date)} ${row.startTime}`;
}

function formatAppointmentRangeLabel(row) {
  const date = new Date(`${row.date}T00:00:00`);
  return `${dateFormatter.format(date)} ${row.startTime}-${row.endTime}`;
}

function toggleDownloadButtons() {
  const hasRows = getMergedCalendarRows(getSortedValidRows()).length > 0;
  const hasSelectedRows = getMergedCalendarRows(getSelectedValidRows()).length > 0;
  elements.downloadIcsButton.disabled = !hasRows;
  elements.downloadCsvButton.disabled = !hasRows;
  elements.downloadJsonButton.disabled = !hasRows;
  elements.addCalendarButton.disabled = !hasSelectedRows;
}

function getSortedValidRows() {
  return [...state.rows]
    .filter((row) => row.date && row.startTime && row.endTime)
    .sort((left, right) =>
      `${left.date}T${left.startTime}`.localeCompare(`${right.date}T${right.startTime}`)
    );
}

function getSelectedValidRows() {
  return getSortedValidRows().filter((row) => row.selected !== false);
}

function getMergedCalendarRows(rows) {
  if (rows.length === 0) {
    return [];
  }

  const mergedRows = [];
  let current = createMergedCalendarRow(rows[0]);

  for (const row of rows.slice(1)) {
    if (canMergeAdjacentSessions(current, row)) {
      current = mergeCalendarRow(current, row);
      continue;
    }

    mergedRows.push(current);
    current = createMergedCalendarRow(row);
  }

  mergedRows.push(current);
  return mergedRows;
}

function createMergedCalendarRow(row) {
  const sourceRows = [row];
  return {
    ...row,
    sourceRows,
    mergedCount: 1,
    summary: buildMergedSummary(sourceRows),
    therapist: joinUniqueValues(sourceRows.map((item) => item.therapist)),
    location: joinUniqueValues(sourceRows.map((item) => item.location)),
    procedure: joinUniqueValues(sourceRows.map((item) => item.procedure)),
    description: buildMergedDescription(sourceRows)
  };
}

function canMergeAdjacentSessions(entry, nextRow) {
  return entry.date === nextRow.date && entry.endTime === nextRow.startTime;
}

function mergeCalendarRow(entry, nextRow) {
  const sourceRows = [...entry.sourceRows, nextRow];
  return {
    ...entry,
    endTime: nextRow.endTime,
    sourceRows,
    mergedCount: sourceRows.length,
    summary: buildMergedSummary(sourceRows),
    therapist: joinUniqueValues(sourceRows.map((item) => item.therapist)),
    location: joinUniqueValues(sourceRows.map((item) => item.location)),
    procedure: joinUniqueValues(sourceRows.map((item) => item.procedure)),
    description: buildMergedDescription(sourceRows)
  };
}

function buildMergedSummary(rows) {
  const uniqueSummaries = [...new Set(rows.map((row) => row.summary || buildSummary(row)).filter(Boolean))];
  if (uniqueSummaries.length === 1) {
    return uniqueSummaries[0];
  }
  return `Royal Rehab session block (${rows.length})`;
}

function buildMergedDescription(rows) {
  return rows
    .map((row, index) => {
      const fragments = [
        `${index + 1}. ${row.startTime}-${row.endTime}`,
        row.summary || buildSummary(row)
      ];
      if (row.procedure) {
        fragments.push(`Procedure: ${row.procedure}`);
      }
      return fragments.join(" | ");
    })
    .join("\n");
}

function joinUniqueValues(values) {
  const uniqueValues = [...new Set(values.map((value) => String(value ?? "").trim()).filter(Boolean))];
  return uniqueValues.join(" / ");
}

function downloadIcs() {
  const rows = getMergedCalendarRows(getSortedValidRows());
  if (rows.length === 0) {
    setStatus("error", "No rows to export", "Add or fix at least one appointment row first.");
    return;
  }

  const ics = buildIcs(rows, state.metadata);
  saveTextFile(`${buildFileStem()}.ics`, "text/calendar; charset=utf-8", ics);
}

function downloadCsv() {
  const rows = getMergedCalendarRows(getSortedValidRows());
  if (rows.length === 0) {
    setStatus("error", "No rows to export", "Add or fix at least one appointment row first.");
    return;
  }

  const header = ["date", "startTime", "endTime", "therapist", "location", "procedure", "summary", "mergedCount"];
  const lines = [header.join(",")];

  for (const row of rows) {
    lines.push(
      [row.date, row.startTime, row.endTime, row.therapist, row.location, row.procedure, row.summary, row.mergedCount]
        .map(csvEscape)
        .join(",")
    );
  }

  saveTextFile(`${buildFileStem()}.csv`, "text/csv; charset=utf-8", lines.join("\r\n"));
}

function downloadJson() {
  const rows = getMergedCalendarRows(getSortedValidRows());
  if (rows.length === 0) {
    setStatus("error", "No rows to export", "Add or fix at least one appointment row first.");
    return;
  }

  const payload = {
    metadata: state.metadata,
    sessions: rows
  };
  saveTextFile(`${buildFileStem()}.json`, "application/json; charset=utf-8", `${JSON.stringify(payload, null, 2)}\n`);
}

async function addSelectedToCalendar() {
  const rows = getMergedCalendarRows(getSelectedValidRows());
  if (rows.length === 0) {
    setStatus("error", "No selected rows", "Select at least one appointment first.");
    return;
  }

  const provider = elements.calendarProvider.value;
  if (provider === "google" && state.googleCalendar.ready) {
    const connected = await ensureGoogleCalendarConnection();
    if (!connected) {
      return;
    }

    const shouldInsert = window.confirm(
      `Add ${rows.length} merged calendar entr${rows.length === 1 ? "y" : "ies"} to your Google Calendar?`
    );

    if (!shouldInsert) {
      return;
    }

    try {
      setStatus(
        "processing",
        "Adding to Google Calendar",
        `Creating ${rows.length} Google Calendar entr${rows.length === 1 ? "y" : "ies"}.`
      );
      for (const row of rows) {
        const response = await fetch("https://www.googleapis.com/calendar/v3/calendars/primary/events", {
          method: "POST",
          headers: {
            Authorization: `Bearer ${state.googleCalendar.accessToken}`,
            "Content-Type": "application/json"
          },
          body: JSON.stringify(buildGoogleCalendarEvent(row))
        });

        if (!response.ok) {
          const errorPayload = await safeReadJson(response);
          const message = errorPayload?.error?.message || `Google Calendar returned ${response.status}`;
          throw new Error(message);
        }
      }

      setStatus(
        "success",
        "Google Calendar updated",
        `${rows.length} merged calendar entr${rows.length === 1 ? "y has" : "ies have"} been added to your Google Calendar.`
      );
      void logAuditEvent("google_calendar_add", "success", {
        mergedEntries: rows.length,
        patientName: state.metadata.patientName
      });
      return;
    } catch (error) {
      console.error(error);
      void logAuditEvent("google_calendar_add", "error", {
        mergedEntries: rows.length,
        message: error instanceof Error ? error.message : "The Google Calendar request failed."
      });
      setStatus(
        "error",
        "Google Calendar add failed",
        error instanceof Error ? error.message : "The Google Calendar request failed."
      );
      return;
    }
  }

  const providerLabel = provider === "outlook" ? "Outlook Calendar" : "Google Calendar";
  const filename = `${buildFileStem()}-${provider}-selected.ics`;
  const ics = buildIcs(rows, state.metadata);
  const downloaded = await saveTextFile(filename, "text/calendar; charset=utf-8", ics, { silentSuccess: true });

  if (!downloaded) {
    return;
  }

  const importUrl = provider === "outlook"
    ? "https://outlook.office.com/calendar/0/addcalendar"
    : "https://calendar.google.com/calendar/u/0/r/settings/export";
  window.open(importUrl, "_blank", "noopener,noreferrer");

  setStatus(
    "success",
    "Batch calendar ready",
    `${rows.length} calendar entr${rows.length === 1 ? "y was" : "ies were"} packed into one ICS file for ${providerLabel}.`
  );
  void logAuditEvent("calendar_batch_prepared", "success", {
    provider,
    mergedEntries: rows.length,
    patientName: state.metadata.patientName
  });
}

function buildIcs(rows, metadata) {
  const timezone = metadata.timezone || "Australia/Sydney";
  const lines = [
    "BEGIN:VCALENDAR",
    "VERSION:2.0",
    "PRODID:-//Royal Rehab//Schedule Converter//EN",
    "CALSCALE:GREGORIAN",
    "METHOD:PUBLISH",
    `X-WR-CALNAME:${escapeIcsText(metadata.calendarTitle || "Royal Rehab Schedule")}`,
    `X-WR-TIMEZONE:${timezone}`,
    ...buildTimeZoneBlock(timezone)
  ];

  const stamp = formatUtcStamp(new Date());

  for (const row of rows) {
    const start = formatLocalIcsDateTime(row.date, row.startTime);
    const end = formatLocalIcsDateTime(row.date, row.endTime);
    const descriptionLines = [
      metadata.patientName ? `Patient: ${metadata.patientName}` : "",
      metadata.mrn ? `MRN: ${metadata.mrn}` : "",
      row.mergedCount > 1 ? `Merged sessions: ${row.mergedCount}` : "",
      row.therapist ? `Session: ${row.therapist}` : "",
      row.location ? `Location: ${row.location}` : "",
      row.procedure ? `Procedure: ${row.procedure}` : "",
      row.description ? `Details:\n${row.description}` : ""
    ].filter(Boolean);

    lines.push("BEGIN:VEVENT");
    lines.push(`UID:${buildUid(row)}`);
    lines.push(`DTSTAMP:${stamp}`);
    lines.push(`DTSTART;TZID=${timezone}:${start}`);
    lines.push(`DTEND;TZID=${timezone}:${end}`);
    lines.push(`SUMMARY:${escapeIcsText(row.summary || buildSummary(row))}`);
    if (row.location) {
      lines.push(`LOCATION:${escapeIcsText(row.location)}`);
    }
    if (descriptionLines.length > 0) {
      lines.push(`DESCRIPTION:${escapeIcsText(descriptionLines.join("\n"))}`);
    }
    lines.push("END:VEVENT");
  }

  lines.push("END:VCALENDAR");
  return lines.map(foldIcsLine).join("\r\n");
}

function buildTimeZoneBlock(timezone) {
  if (timezone !== "Australia/Sydney") {
    return [];
  }

  return [
    "BEGIN:VTIMEZONE",
    "TZID:Australia/Sydney",
    "X-LIC-LOCATION:Australia/Sydney",
    "BEGIN:DAYLIGHT",
    "TZOFFSETFROM:+1000",
    "TZOFFSETTO:+1100",
    "TZNAME:AEDT",
    "DTSTART:19701004T020000",
    "RRULE:FREQ=YEARLY;BYMONTH=10;BYDAY=1SU",
    "END:DAYLIGHT",
    "BEGIN:STANDARD",
    "TZOFFSETFROM:+1100",
    "TZOFFSETTO:+1000",
    "TZNAME:AEST",
    "DTSTART:19700405T030000",
    "RRULE:FREQ=YEARLY;BYMONTH=4;BYDAY=1SU",
    "END:STANDARD",
    "END:VTIMEZONE"
  ];
}

function foldIcsLine(line) {
  if (line.length <= 74) {
    return line;
  }

  const segments = [];
  let remaining = line;
  while (remaining.length > 74) {
    segments.push(remaining.slice(0, 74));
    remaining = ` ${remaining.slice(74)}`;
  }
  segments.push(remaining);
  return segments.join("\r\n");
}

function buildUid(row) {
  const slug = [row.date, row.startTime, row.summary || row.therapist || "appointment"]
    .join("-")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "");

  return `${slug}@royalrehab.local`;
}

function formatUtcStamp(date) {
  return [
    date.getUTCFullYear(),
    String(date.getUTCMonth() + 1).padStart(2, "0"),
    String(date.getUTCDate()).padStart(2, "0")
  ].join("") +
    "T" +
    [
      String(date.getUTCHours()).padStart(2, "0"),
      String(date.getUTCMinutes()).padStart(2, "0"),
      String(date.getUTCSeconds()).padStart(2, "0")
    ].join("") +
    "Z";
}

function formatLocalIcsDateTime(date, time) {
  const [year, month, day] = date.split("-");
  const [hours, minutes] = time.split(":");
  return `${year}${month}${day}T${hours}${minutes}00`;
}

function escapeIcsText(value) {
  return String(value)
    .replace(/\\/g, "\\\\")
    .replace(/\n/g, "\\n")
    .replace(/,/g, "\\,")
    .replace(/;/g, "\\;");
}

async function saveTextFile(filename, mimeType, content, options = {}) {
  try {
    setStatus("processing", "Preparing download", `Building ${filename} for download.`);
    const response = await fetch("/api/exports", {
      method: "POST",
      headers: {
        "Content-Type": "application/json"
      },
      body: JSON.stringify({ filename, mimeType, content })
    });

    if (!response.ok) {
      throw new Error(`Export request failed with status ${response.status}`);
    }

    const payload = await response.json();
    const anchor = document.createElement("a");
    anchor.href = payload.downloadUrl;
    anchor.download = filename;
    anchor.style.display = "none";
    document.body.append(anchor);
    anchor.click();
    anchor.remove();
    if (!options.silentSuccess) {
      setStatus("success", "Download ready", `${filename} has been handed off to the browser download flow.`);
    }
    void logAuditEvent("download_prepared", "success", {
      filename,
      mimeType,
      sizeBytes: content.length
    });
    return true;
  } catch (error) {
    console.error(error);
    void logAuditEvent("download_prepared", "error", {
      filename,
      mimeType,
      message: error instanceof Error ? error.message : "The export could not be generated."
    });
    setStatus(
      "error",
      "Download failed",
      error instanceof Error ? error.message : "The export could not be generated."
    );
    return false;
  }
}

async function connectGoogleCalendar() {
  if (!state.googleCalendar.ready) {
    setStatus(
      "error",
      "Google Calendar not configured",
      "Set GOOGLE_CLIENT_ID to enable Connect Google and direct calendar insertion."
    );
    return;
  }

  if (isNonLocalGoogleOrigin()) {
    setStatus(
      "processing",
      "Google connection check",
      `If Google shows invalid_request, add ${window.location.origin} to the OAuth client's Authorized JavaScript origins first.`
    );
  }

  const connected = await ensureGoogleCalendarConnection();
  if (connected) {
    setStatus("success", "Google Calendar connected", "Google Calendar is ready for one-step event creation.");
  }
}

async function ensureGoogleCalendarConnection() {
  if (!state.googleCalendar.ready || !state.googleCalendar.tokenClient) {
    syncGoogleCalendarUi();
    return false;
  }

  if (hasUsableGoogleAccessToken()) {
    state.googleCalendar.connected = true;
    syncGoogleCalendarUi();
    return true;
  }

  try {
    await new Promise((resolve, reject) => {
      state.googleCalendar.tokenClient.callback = (response) => {
        if (response?.error) {
          reject(new Error(response.error));
          return;
        }
        state.googleCalendar.accessToken = String(response?.access_token ?? "");
        state.googleCalendar.tokenExpiresAt = Date.now() + (Number(response?.expires_in ?? 0) * 1000);
        state.googleCalendar.connected = true;
        resolve(response);
      };
      state.googleCalendar.tokenClient.requestAccessToken({ prompt: "consent" });
    });

    syncGoogleCalendarUi();
    return true;
  } catch (error) {
    console.error(error);
    state.googleCalendar.connected = false;
    syncGoogleCalendarUi();
    const message = error instanceof Error ? error.message : "Google sign-in was not completed.";
    setStatus(
      "error",
      "Google connection failed",
      message === "invalid_request"
        ? `Google rejected the OAuth request. Add ${window.location.origin} to the OAuth client's Authorized JavaScript origins, then refresh and try again.`
        : message
    );
    return false;
  }
}

function syncGoogleCalendarUi() {
  if (elements.calendarProvider.value !== "google") {
    elements.googleConnectButton.hidden = true;
    elements.googleStatusPill.hidden = true;
    return;
  }

  elements.googleConnectButton.hidden = false;
  elements.googleStatusPill.hidden = false;

  if (!state.googleCalendar.enabled) {
    elements.googleConnectButton.disabled = true;
    elements.googleConnectButton.textContent = "Connect Google";
    elements.googleStatusPill.textContent = "Server setup needed for direct add";
    return;
  }

  elements.googleConnectButton.disabled = false;

  if (hasUsableGoogleAccessToken()) {
    state.googleCalendar.connected = true;
    elements.googleConnectButton.textContent = "Google connected";
    elements.googleStatusPill.textContent = "Google direct add ready";
    return;
  }

  elements.googleConnectButton.textContent = "Connect Google";
  elements.googleStatusPill.textContent = isNonLocalGoogleOrigin()
    ? `Authorize ${window.location.origin} in Google OAuth`
    : "Google direct add available";
}

function syncSelectionState() {
  if (state.rows.length === 0) {
    elements.selectAllCheckbox.checked = false;
    elements.selectAllCheckbox.indeterminate = false;
    elements.selectAllCheckbox.disabled = true;
    return;
  }

  const selectedCount = state.rows.filter((row) => row.selected !== false).length;
  elements.selectAllCheckbox.disabled = false;
  elements.selectAllCheckbox.checked = selectedCount === state.rows.length;
  elements.selectAllCheckbox.indeterminate = selectedCount > 0 && selectedCount < state.rows.length;
}

function syncCalendarActionLabel() {
  const provider = elements.calendarProvider.value;
  if (provider === "google" && state.googleCalendar.enabled) {
    elements.addCalendarButton.textContent = state.googleCalendar.connected
      ? "Add selected to Google Calendar"
      : "Connect Google & add selected";
    elements.selectedProviderHint.textContent = state.googleCalendar.connected
      ? "Google direct add with one confirmation"
      : "Use Connect Google to pick your account, then add everything in one go";
  } else if (provider === "google") {
    elements.addCalendarButton.textContent = "Prepare selected for Google Calendar";
    elements.selectedProviderHint.textContent = "Direct Google add needs server setup; ICS import remains available";
  } else {
    const providerLabel = provider === "outlook" ? "Outlook Calendar" : "Google Calendar";
    elements.addCalendarButton.textContent = `Prepare selected for ${providerLabel}`;
    elements.selectedProviderHint.textContent = `${providerLabel} batch import`;
  }
  syncGoogleCalendarUi();
}

function renderSelectedBatch() {
  const selectedRows = getMergedCalendarRows(getSelectedValidRows());
  elements.selectedCountLabel.textContent = `${selectedRows.length} calendar entr${selectedRows.length === 1 ? "y" : "ies"}`;

  if (selectedRows.length === 0) {
    elements.selectedList.innerHTML = "<li>No appointments selected yet.</li>";
    return;
  }

  elements.selectedList.innerHTML = selectedRows
    .map((row) => {
      const summary = row.summary || buildSummary(row);
      const mergedBadge = row.mergedCount > 1 ? ` (${row.mergedCount} sessions merged)` : "";
      return `<li>${escapeHtml(formatAppointmentRangeLabel(row))} - ${escapeHtml(summary + mergedBadge)}</li>`;
    })
    .join("");
}

function buildFileStem() {
  const patientStem = state.metadata.patientName
    ? state.metadata.patientName.toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/^-+|-+$/g, "")
    : state.sourceName.toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/^-+|-+$/g, "");
  return `${patientStem || "royalrehab-schedule"}-export`;
}

function buildGoogleCalendarEvent(row) {
  return {
    summary: row.summary || buildSummary(row),
    location: row.location || undefined,
    description: buildGoogleCalendarDescription(row),
    start: {
      dateTime: `${row.date}T${row.startTime}:00`,
      timeZone: state.metadata.timezone || "Australia/Sydney"
    },
    end: {
      dateTime: `${row.date}T${row.endTime}:00`,
      timeZone: state.metadata.timezone || "Australia/Sydney"
    }
  };
}

function buildGoogleCalendarDescription(row) {
  return [
    state.metadata.patientName ? `Patient: ${state.metadata.patientName}` : "",
    state.metadata.mrn ? `MRN: ${state.metadata.mrn}` : "",
    row.mergedCount > 1 ? `Merged sessions: ${row.mergedCount}` : "",
    row.therapist ? `Session: ${row.therapist}` : "",
    row.location ? `Location: ${row.location}` : "",
    row.procedure ? `Procedure: ${row.procedure}` : "",
    row.description ? `Details:\n${row.description}` : ""
  ].filter(Boolean).join("\n");
}

function hasUsableGoogleAccessToken() {
  return Boolean(
    state.googleCalendar.accessToken &&
    (!state.googleCalendar.tokenExpiresAt || state.googleCalendar.tokenExpiresAt > Date.now() + 30_000)
  );
}

async function safeReadJson(response) {
  try {
    return await response.json();
  } catch {
    return null;
  }
}

function isNonLocalGoogleOrigin() {
  return !/^https?:\/\/(localhost|127\.0\.0\.1)(:\d+)?$/i.test(window.location.origin);
}

function csvEscape(value) {
  const safeValue = String(value ?? "");
  return `"${safeValue.replace(/"/g, "\"\"")}"`;
}

function sortByYThenX(left, right) {
  return left.y0 === right.y0 ? left.x0 - right.x0 : left.y0 - right.y0;
}

function setStatus(kind, title, text) {
  elements.statusPanel.className = `status-panel ${kind}`;
  elements.statusTitle.textContent = title;
  elements.statusText.textContent = text;
}

function capitalize(text) {
  return text.charAt(0).toUpperCase() + text.slice(1);
}

function escapeHtml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/"/g, "&quot;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}
