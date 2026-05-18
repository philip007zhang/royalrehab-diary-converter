import { randomUUID } from "node:crypto";
import { mkdtemp, readdir, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { extname, join } from "node:path";
import { spawn } from "node:child_process";

import { dedupeRows, parseScheduleFromOcr } from "./public/shared/schedule-parser.js";

export async function getServerOcrStatus() {
  const tesseractReady = await commandWorks("tesseract", ["--version"]);
  const pdftoppmReady = await commandWorks("pdftoppm", ["-v"]);

  return {
    enabled: tesseractReady,
    imageOcr: tesseractReady,
    pdfOcr: tesseractReady && pdftoppmReady,
    tools: {
      tesseract: tesseractReady,
      pdftoppm: pdftoppmReady
    }
  };
}

export async function extractScheduleFromUpload({ buffer, filename, mimeType }) {
  return extractScheduleFromUploadWithProgress({ buffer, filename, mimeType });
}

export async function extractScheduleFromUploadWithProgress({ buffer, filename, mimeType }, onProgress = () => {}) {
  const safeName = sanitiseFilename(filename || "upload");
  const sourceExtension = extname(safeName).toLowerCase() || extensionFromMime(mimeType);
  const tempDir = await mkdtemp(join(tmpdir(), "royalrehab-ocr-"));

  try {
    onProgress({
      phase: "validating",
      progress: 5,
      message: "Checking the uploaded file and OCR tools."
    });

    const sourcePath = join(tempDir, `source${sourceExtension || ".bin"}`);
    await writeFile(sourcePath, buffer);
    onProgress({
      phase: "stored",
      progress: 12,
      message: "The upload has reached the server."
    });

    const isPdf = mimeType === "application/pdf" || /\.pdf$/i.test(safeName);
    const status = await getServerOcrStatus();
    if (!status.imageOcr) {
      throw new Error("Server OCR is not installed. Add the tesseract command on the server first.");
    }
    if (isPdf && !status.pdfOcr) {
      throw new Error("Server PDF OCR is not installed. Add the pdftoppm command on the server first.");
    }

    if (isPdf) {
      onProgress({
        phase: "rendering_pdf",
        progress: 18,
        message: "Rendering the PDF into OCR-ready pages."
      });
    } else {
      onProgress({
        phase: "preparing_image",
        progress: 18,
        message: "Preparing the uploaded image for OCR."
      });
    }

    const pagePaths = isPdf ? await renderPdfPages(sourcePath, tempDir) : [sourcePath];
    onProgress({
      phase: isPdf ? "pdf_ready" : "image_ready",
      progress: 30,
      message: isPdf
        ? `Prepared ${pagePaths.length} page${pagePaths.length === 1 ? "" : "s"} for OCR.`
        : "The uploaded image is ready for OCR."
    });

    const rows = [];
    const ocrParts = [];
    let metadata = {};

    for (const [index, pagePath] of pagePaths.entries()) {
      const scanProgress = 30 + Math.round(((index + 0.35) / pagePaths.length) * 50);
      onProgress({
        phase: "ocr_scanning",
        progress: scanProgress,
        message: `Scanning page ${index + 1} of ${pagePaths.length} with OCR.`
      });

      const pageResult = await recognizePage(pagePath);
      const parseProgress = 30 + Math.round(((index + 0.8) / pagePaths.length) * 50);
      onProgress({
        phase: "parsing_rows",
        progress: parseProgress,
        message: `Parsing appointment rows from page ${index + 1} of ${pagePaths.length}.`
      });
      const parsed = parseScheduleFromOcr(
        { words: pageResult.words, text: pageResult.text, blocks: [] },
        pageResult.width,
        pageResult.height,
        () => randomUUID()
      );

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

      const text = pageResult.text.trim();
      if (text) {
        ocrParts.push(
          pagePaths.length > 1
            ? `--- Page ${index + 1} ---\n${text}`
            : text
        );
      }
    }

    onProgress({
      phase: "finalizing",
      progress: 94,
      message: "Finalizing the extracted rows and metadata."
    });

    return {
      metadata,
      ocrText: ocrParts.join("\n\n").trim(),
      pageCount: pagePaths.length,
      rows: dedupeRows(rows)
    };
  } finally {
    await rm(tempDir, { recursive: true, force: true });
  }
}

async function commandWorks(command, args) {
  try {
    await runCommand(command, args);
    return true;
  } catch {
    return false;
  }
}

async function renderPdfPages(sourcePath, tempDir) {
  const prefix = join(tempDir, "page");
  await runCommand("pdftoppm", ["-png", "-r", "220", sourcePath, prefix]);

  const filenames = (await readdir(tempDir))
    .filter((name) => /^page-\d+\.png$/i.test(name))
    .sort((left, right) => extractPageNumber(left) - extractPageNumber(right));

  if (filenames.length === 0) {
    throw new Error("The PDF could not be rendered into OCR pages on the server.");
  }

  return filenames.map((name) => join(tempDir, name));
}

async function recognizePage(imagePath) {
  const tsv = await runCommand("tesseract", [imagePath, "stdout", "-l", "eng", "--psm", "6", "tsv"]);
  const text = await runCommand("tesseract", [imagePath, "stdout", "-l", "eng", "--psm", "6"]);
  return parseTesseractTsv(tsv.stdout, text.stdout);
}

function parseTesseractTsv(tsvText, plainText) {
  const lines = tsvText.split(/\r?\n/).filter(Boolean);
  const words = [];
  let pageWidth = 0;
  let pageHeight = 0;

  for (const line of lines.slice(1)) {
    const columns = line.split("\t");
    if (columns.length < 12) {
      continue;
    }

    const level = Number.parseInt(columns[0], 10);
    const left = Number.parseFloat(columns[6] || "0");
    const top = Number.parseFloat(columns[7] || "0");
    const width = Number.parseFloat(columns[8] || "0");
    const height = Number.parseFloat(columns[9] || "0");
    const confidence = Number.parseFloat(columns[10] || "0");
    const text = (columns[11] || "").trim();

    if (level === 1) {
      pageWidth = Math.max(pageWidth, width);
      pageHeight = Math.max(pageHeight, height);
      continue;
    }

    if (level !== 5 || !text) {
      continue;
    }

    words.push({
      text,
      confidence,
      x0: left,
      y0: top,
      x1: left + width,
      y1: top + height
    });
  }

  if (!pageWidth || !pageHeight) {
    for (const word of words) {
      pageWidth = Math.max(pageWidth, word.x1);
      pageHeight = Math.max(pageHeight, word.y1);
    }
  }

  return {
    width: Math.ceil(pageWidth || 1200),
    height: Math.ceil(pageHeight || 1600),
    text: plainText.trim(),
    words
  };
}

function runCommand(command, args) {
  return new Promise((resolve, reject) => {
    const child = spawn(command, args, { windowsHide: true });
    const stdoutChunks = [];
    const stderrChunks = [];

    child.stdout.on("data", (chunk) => {
      stdoutChunks.push(chunk);
    });

    child.stderr.on("data", (chunk) => {
      stderrChunks.push(chunk);
    });

    child.on("error", (error) => {
      reject(new Error(`${command} is unavailable: ${error.message}`));
    });

    child.on("close", (code) => {
      const stdout = Buffer.concat(stdoutChunks).toString("utf8");
      const stderr = Buffer.concat(stderrChunks).toString("utf8");

      if (code === 0) {
        resolve({ stdout, stderr });
        return;
      }

      const message = stderr.trim() || stdout.trim() || `${command} exited with code ${code}`;
      reject(new Error(message));
    });
  });
}

function extractPageNumber(filename) {
  const match = filename.match(/-(\d+)\.png$/i);
  return Number.parseInt(match?.[1] ?? "0", 10);
}

function extensionFromMime(mimeType) {
  const value = String(mimeType ?? "").toLowerCase();
  if (value === "application/pdf") {
    return ".pdf";
  }
  if (value === "image/jpeg") {
    return ".jpg";
  }
  if (value === "image/png") {
    return ".png";
  }
  if (value === "image/webp") {
    return ".webp";
  }
  if (value === "image/bmp") {
    return ".bmp";
  }
  return "";
}

function sanitiseFilename(filename) {
  return String(filename ?? "")
    .replace(/[^\w.\-() ]+/g, "_")
    .trim();
}
