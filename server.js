import { appendFile, createReadStream, existsSync, readFileSync, statSync } from "node:fs";
import { randomUUID } from "node:crypto";
import { createServer } from "node:http";
import { extname, join, normalize } from "node:path";
import { Readable } from "node:stream";

import { extractScheduleFromUpload, extractScheduleFromUploadWithProgress, getServerOcrStatus } from "./server-ocr.js";

loadEnvFile(".env.local");
loadEnvFile(".env");

const host = process.env.HOST ?? "0.0.0.0";
const port = Number.parseInt(process.env.PORT ?? "3000", 10);
const rootDir = process.cwd();
const publicDir = join(rootDir, "public");
const auditLogPath = join(rootDir, "audit.log.jsonl");
const exportStore = new Map();
const ocrJobStore = new Map();
const sessionStore = new Map();
const exportTtlMs = 10 * 60 * 1000;
const ocrJobTtlMs = 30 * 60 * 1000;
const sessionTtlMs = 12 * 60 * 60 * 1000;
const googleCalendarConfig = {
  apiKey: process.env.GOOGLE_API_KEY ?? "",
  clientId: process.env.GOOGLE_CLIENT_ID ?? ""
};
const adminAuth = {
  username: process.env.ADMIN_USERNAME ?? "admin",
  password: process.env.ADMIN_PASSWORD ?? "ChangeMe123!"
};

const mimeTypes = {
  ".css": "text/css; charset=utf-8",
  ".html": "text/html; charset=utf-8",
  ".js": "application/javascript; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".pdf": "application/pdf",
  ".png": "image/png",
  ".jpg": "image/jpeg",
  ".jpeg": "image/jpeg",
  ".svg": "image/svg+xml; charset=utf-8",
  ".webp": "image/webp",
  ".txt": "text/plain; charset=utf-8"
};

function loadEnvFile(filename) {
  const filePath = join(process.cwd(), filename);
  if (!existsSync(filePath)) {
    return;
  }

  const raw = readFileSync(filePath, "utf8");
  for (const line of raw.split(/\r?\n/)) {
    const trimmed = line.trim();
    if (!trimmed || trimmed.startsWith("#")) {
      continue;
    }

    const separatorIndex = trimmed.indexOf("=");
    if (separatorIndex === -1) {
      continue;
    }

    const key = trimmed.slice(0, separatorIndex).trim();
    if (!key || key in process.env) {
      continue;
    }

    let value = trimmed.slice(separatorIndex + 1).trim();
    if (
      (value.startsWith("\"") && value.endsWith("\"")) ||
      (value.startsWith("'") && value.endsWith("'"))
    ) {
      value = value.slice(1, -1);
    }

    process.env[key] = value;
  }
}

function safeJoin(baseDir, requestPath) {
  const resolvedPath = normalize(join(baseDir, requestPath));
  return resolvedPath.startsWith(baseDir) ? resolvedPath : null;
}

function sendFile(response, filePath, options = {}) {
  if (!existsSync(filePath) || !statSync(filePath).isFile()) {
    response.writeHead(404, { "Content-Type": "text/plain; charset=utf-8" });
    response.end("Not found");
    return;
  }

  response.writeHead(200, {
    "Content-Type": mimeTypes[extname(filePath).toLowerCase()] ?? "application/octet-stream",
    "Cache-Control": "no-store"
  });

  if (options.head) {
    response.end();
    return;
  }

  createReadStream(filePath).pipe(response);
}

function sendJson(response, statusCode, payload, headers = {}, options = {}) {
  response.writeHead(statusCode, {
    "Content-Type": "application/json; charset=utf-8",
    ...headers
  });
  if (options.head) {
    response.end();
    return;
  }
  response.end(JSON.stringify(payload));
}

function sendHtml(response, statusCode, html, options = {}) {
  response.writeHead(statusCode, {
    "Content-Type": "text/html; charset=utf-8",
    "Cache-Control": "no-store"
  });
  if (options.head) {
    response.end();
    return;
  }
  response.end(html);
}

function redirect(response, location) {
  response.writeHead(302, { Location: location, "Cache-Control": "no-store" });
  response.end();
}

function collectRequestBody(request) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    let totalBytes = 0;

    request.on("data", (chunk) => {
      totalBytes += chunk.length;
      if (totalBytes > 2 * 1024 * 1024) {
        reject(new Error("Request body too large"));
        request.destroy();
        return;
      }
      chunks.push(chunk);
    });

    request.on("end", () => resolve(Buffer.concat(chunks).toString("utf8")));
    request.on("error", reject);
  });
}

async function collectMultipartForm(request, maxBytes = 15 * 1024 * 1024) {
  const contentLength = Number.parseInt(request.headers["content-length"] ?? "0", 10);
  if (Number.isFinite(contentLength) && contentLength > maxBytes) {
    throw new Error("Uploaded file is too large.");
  }

  const multipartRequest = new Request(`http://${request.headers.host ?? `${host}:${port}`}/upload`, {
    method: request.method ?? "POST",
    headers: request.headers,
    body: Readable.toWeb(request),
    duplex: "half"
  });

  const formData = await multipartRequest.formData();
  const sourceFile = formData.get("sourceFile");
  if (!(sourceFile instanceof File)) {
    throw new Error("Please choose an image or PDF to extract.");
  }

  return {
    buffer: Buffer.from(await sourceFile.arrayBuffer()),
    filename: sourceFile.name,
    mimeType: sourceFile.type || mimeTypes[extname(sourceFile.name).toLowerCase()] || "application/octet-stream",
    metadata: {
      calendarTitle: String(formData.get("calendarTitle") ?? "").trim(),
      timezone: String(formData.get("timezone") ?? "").trim()
    }
  };
}

function createStoredExport(filename, mimeType, content) {
  cleanupExpiredExports();
  const id = randomUUID();
  exportStore.set(id, {
    content,
    createdAt: Date.now(),
    filename,
    mimeType
  });

  return {
    id,
    downloadUrl: `/downloads/${id}`,
    filename,
    mimeType
  };
}

function cleanupExpiredExports() {
  const cutoff = Date.now() - exportTtlMs;
  for (const [id, entry] of exportStore.entries()) {
    if (entry.createdAt < cutoff) {
      exportStore.delete(id);
    }
  }
}

function cleanupExpiredOcrJobs() {
  const cutoff = Date.now() - ocrJobTtlMs;
  for (const [id, entry] of ocrJobStore.entries()) {
    if (entry.createdAt < cutoff) {
      ocrJobStore.delete(id);
    }
  }
}

function cleanupExpiredSessions() {
  const cutoff = Date.now() - sessionTtlMs;
  for (const [token, entry] of sessionStore.entries()) {
    if (entry.createdAt < cutoff) {
      sessionStore.delete(token);
    }
  }
}

function parseCookies(request) {
  const rawCookieHeader = request.headers.cookie ?? "";
  const pairs = rawCookieHeader.split(";").map((item) => item.trim()).filter(Boolean);
  const cookies = {};

  for (const pair of pairs) {
    const separatorIndex = pair.indexOf("=");
    if (separatorIndex === -1) {
      continue;
    }
    const name = pair.slice(0, separatorIndex);
    const value = pair.slice(separatorIndex + 1);
    cookies[name] = decodeURIComponent(value);
  }

  return cookies;
}

function getAdminSession(request) {
  cleanupExpiredSessions();
  const cookies = parseCookies(request);
  const token = cookies.audit_session;
  if (!token) {
    return null;
  }

  const session = sessionStore.get(token);
  return session ?? null;
}

function buildAuditEntry(activity, status, details = {}, request = null) {
  return {
    id: randomUUID(),
    timestamp: new Date().toISOString(),
    activity,
    status,
    details,
    actor: request ? {
      ip: request.socket.remoteAddress ?? "",
      userAgent: request.headers["user-agent"] ?? ""
    } : null
  };
}

function appendAuditEntry(entry) {
  appendFile(auditLogPath, `${JSON.stringify(entry)}\n`, (error) => {
    if (error) {
      console.error("Failed to write audit log entry", error);
    }
  });
}

function readAuditEntries() {
  if (!existsSync(auditLogPath)) {
    return [];
  }

  const raw = readFileSync(auditLogPath, "utf8");
  return raw
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean)
    .map((line) => {
      try {
        return JSON.parse(line);
      } catch {
        return null;
      }
    })
    .filter(Boolean)
    .reverse();
}

async function handleExportCreate(request, response) {
  try {
    const body = await collectRequestBody(request);
    const parsed = JSON.parse(body);
    const filename = String(parsed.filename ?? "").trim();
    const mimeType = String(parsed.mimeType ?? "").trim();
    const content = String(parsed.content ?? "");

    if (!filename || !mimeType) {
      sendJson(response, 400, { error: "Missing filename or mimeType" });
      return;
    }

    const storedExport = createStoredExport(filename, mimeType, content);

    appendAuditEntry(buildAuditEntry(
      "export_generated",
      "success",
      { filename, mimeType, sizeBytes: Buffer.byteLength(content, "utf8") },
      request
    ));

    sendJson(response, 201, {
      downloadUrl: storedExport.downloadUrl,
      id: storedExport.id
    });
  } catch (error) {
    appendAuditEntry(buildAuditEntry(
      "export_generated",
      "error",
      { message: error instanceof Error ? error.message : "Invalid export request" },
      request
    ));
    const message = error instanceof Error ? error.message : "Invalid export request";
    sendJson(response, 400, { error: message });
  }
}

function handleExportDownload(response, exportId, options = {}) {
  cleanupExpiredExports();
  const entry = exportStore.get(exportId);

  if (!entry) {
    response.writeHead(404, { "Content-Type": "text/plain; charset=utf-8" });
    response.end("Export not found or expired");
    return;
  }

  response.writeHead(200, {
    "Content-Type": entry.mimeType,
    "Content-Disposition": `attachment; filename="${entry.filename.replace(/"/g, "")}"`,
    "Cache-Control": "no-store"
  });
  if (options.head) {
    response.end();
    return;
  }
  response.end(entry.content);
}

async function handleServerOcrRequest(request, response, options = {}) {
  try {
    const upload = await collectMultipartForm(request);
    const extraction = await extractScheduleFromUpload(upload);
    const metadata = {
      calendarTitle: upload.metadata.calendarTitle || extraction.metadata.calendarTitle || "Royal Rehab Schedule",
      timezone: upload.metadata.timezone || "Australia/Sydney",
      patientName: extraction.metadata.patientName || "",
      mrn: extraction.metadata.mrn || ""
    };
    const payload = {
      ...extraction,
      metadata
    };

    appendAuditEntry(buildAuditEntry(
      options.html ? "server_ocr_fallback" : "server_ocr_api",
      payload.rows.length > 0 ? "success" : "warning",
      {
        fileName: upload.filename,
        fileType: upload.mimeType,
        pageCount: payload.pageCount,
        rowsFound: payload.rows.length
      },
      request
    ));

    if (options.html) {
      const exports = createFallbackExports(payload, upload.filename);
      sendHtml(response, 200, renderNoScriptResultPage({
        extraction: payload,
        exports,
        sourceName: upload.filename
      }));
      return;
    }

    sendJson(response, 200, payload);
  } catch (error) {
    const message = error instanceof Error ? error.message : "Server OCR failed.";
    appendAuditEntry(buildAuditEntry(
      options.html ? "server_ocr_fallback" : "server_ocr_api",
      "error",
      { message },
      request
    ));

    if (options.html) {
      sendHtml(response, 400, renderNoScriptResultPage({ error: message }));
      return;
    }

    sendJson(response, 400, { error: message });
  }
}

function createOcrJob(upload) {
  cleanupExpiredOcrJobs();
  const id = randomUUID();
  const job = {
    id,
    createdAt: Date.now(),
    updatedAt: Date.now(),
    status: "queued",
    phase: "queued",
    progress: 1,
    message: "The upload is queued for server OCR.",
    sourceName: upload.filename,
    sourceType: upload.mimeType,
    sourceKind: upload.mimeType === "application/pdf" || /\.pdf$/i.test(upload.filename) ? "pdf" : "image",
    pageCount: 0,
    result: null,
    error: ""
  };
  ocrJobStore.set(id, job);
  return job;
}

function updateOcrJob(jobId, updates) {
  const job = ocrJobStore.get(jobId);
  if (!job) {
    return null;
  }

  Object.assign(job, updates, { updatedAt: Date.now() });
  if (typeof job.progress === "number") {
    job.progress = Math.max(0, Math.min(100, Math.round(job.progress)));
  }
  return job;
}

function getOcrJobSnapshot(job) {
  return {
    id: job.id,
    status: job.status,
    phase: job.phase,
    progress: job.progress,
    message: job.message,
    sourceName: job.sourceName,
    sourceType: job.sourceType,
    sourceKind: job.sourceKind,
    pageCount: job.pageCount,
    result: job.status === "complete" ? job.result : null,
    error: job.status === "error" ? job.error : ""
  };
}

async function handleServerOcrJobCreate(request, response) {
  try {
    const upload = await collectMultipartForm(request);
    const job = createOcrJob(upload);

    appendAuditEntry(buildAuditEntry(
      "server_ocr_job_created",
      "success",
      {
        jobId: job.id,
        fileName: upload.filename,
        fileType: upload.mimeType
      },
      request
    ));

    void processOcrJob(job.id, upload, request);

    sendJson(response, 202, {
      job: getOcrJobSnapshot(job)
    });
  } catch (error) {
    const message = error instanceof Error ? error.message : "Server OCR job could not be created.";
    appendAuditEntry(buildAuditEntry(
      "server_ocr_job_created",
      "error",
      { message },
      request
    ));
    sendJson(response, 400, { error: message });
  }
}

function handleServerOcrJobStatus(response, jobId) {
  cleanupExpiredOcrJobs();
  const job = ocrJobStore.get(jobId);
  if (!job) {
    sendJson(response, 404, { error: "OCR job not found or expired." });
    return;
  }

  sendJson(response, 200, {
    job: getOcrJobSnapshot(job)
  });
}

async function processOcrJob(jobId, upload, request) {
  updateOcrJob(jobId, {
    status: "processing",
    phase: "starting",
    progress: 3,
    message: "The server has started the OCR task."
  });

  try {
    const extraction = await extractScheduleFromUploadWithProgress(upload, (progressUpdate) => {
      updateOcrJob(jobId, {
        status: "processing",
        phase: progressUpdate.phase,
        progress: progressUpdate.progress,
        message: progressUpdate.message
      });
    });

    const metadata = {
      calendarTitle: upload.metadata.calendarTitle || extraction.metadata.calendarTitle || "Royal Rehab Schedule",
      timezone: upload.metadata.timezone || "Australia/Sydney",
      patientName: extraction.metadata.patientName || "",
      mrn: extraction.metadata.mrn || ""
    };
    const payload = {
      ...extraction,
      metadata
    };

    updateOcrJob(jobId, {
      status: "complete",
      phase: "complete",
      progress: 100,
      message: payload.rows.length > 0
        ? `Server OCR finished with ${payload.rows.length} extracted appointment${payload.rows.length === 1 ? "" : "s"}.`
        : "Server OCR finished, but no appointment rows were confidently parsed.",
      pageCount: payload.pageCount,
      result: payload,
      error: ""
    });

    appendAuditEntry(buildAuditEntry(
      "server_ocr_job_completed",
      payload.rows.length > 0 ? "success" : "warning",
      {
        jobId,
        fileName: upload.filename,
        fileType: upload.mimeType,
        pageCount: payload.pageCount,
        rowsFound: payload.rows.length
      },
      request
    ));
  } catch (error) {
    const message = error instanceof Error ? error.message : "Server OCR failed.";
    updateOcrJob(jobId, {
      status: "error",
      phase: "error",
      progress: 100,
      message,
      error: message
    });

    appendAuditEntry(buildAuditEntry(
      "server_ocr_job_completed",
      "error",
      {
        jobId,
        fileName: upload.filename,
        fileType: upload.mimeType,
        message
      },
      request
    ));
  }
}

function createFallbackExports(extraction, sourceName) {
  const rows = getMergedCalendarRows(getSortedValidRows(extraction.rows));
  const fileStem = buildFileStem(extraction.metadata, sourceName);
  const jsonPayload = `${JSON.stringify({ metadata: extraction.metadata, sessions: rows }, null, 2)}\n`;
  const csvPayload = buildCsv(rows);
  const icsPayload = buildIcs(rows, extraction.metadata);

  return {
    mergedRows: rows,
    ics: createStoredExport(`${fileStem}.ics`, "text/calendar; charset=utf-8", icsPayload),
    csv: createStoredExport(`${fileStem}.csv`, "text/csv; charset=utf-8", csvPayload),
    json: createStoredExport(`${fileStem}.json`, "application/json; charset=utf-8", jsonPayload)
  };
}

function renderNoScriptResultPage({ extraction = null, exports = null, sourceName = "", error = "" }) {
  const hasRows = Boolean(extraction && extraction.rows.length > 0);
  const calendarTitleValue = extraction?.metadata.calendarTitle || "Royal Rehab Schedule";
  const timezoneValue = extraction?.metadata.timezone || "Australia/Sydney";
  const statsMarkup = extraction ? `
    <div class="stats-grid">
      <article class="stat-card">
        <span class="stat-label">Sessions</span>
        <strong>${escapeHtml(String(extraction.rows.length))}</strong>
      </article>
      <article class="stat-card">
        <span class="stat-label">Merged entries</span>
        <strong>${escapeHtml(String(exports?.mergedRows.length ?? 0))}</strong>
      </article>
      <article class="stat-card">
        <span class="stat-label">Source pages</span>
        <strong>${escapeHtml(String(extraction.pageCount ?? 1))}</strong>
      </article>
    </div>
  ` : "";

  const reviewActionsMarkup = exports ? `
    <div class="action-row">
      <a class="primary-button button-link" href="${escapeHtml(exports.ics.downloadUrl)}">Download ICS</a>
      <a class="ghost-link button-link" href="${escapeHtml(exports.csv.downloadUrl)}">Download CSV</a>
      <a class="ghost-link button-link" href="${escapeHtml(exports.json.downloadUrl)}">Download JSON</a>
    </div>
  ` : `<p class="calendar-hint nojs-only">Choose a file above and run server-side extraction to populate this review table.</p>`;

  const tableMarkup = hasRows ? `
    <div class="table-wrap">
      <table>
        <thead>
          <tr>
            <th class="row-select-header">Select</th>
            <th>Date</th>
            <th>Start</th>
            <th>End</th>
            <th>Therapist / Session</th>
            <th>Location</th>
            <th>Procedure</th>
            <th>Calendar title</th>
            <th></th>
          </tr>
        </thead>
        <tbody>
          ${extraction.rows.map((row) => `
            <tr>
              <td class="row-select-cell"><input class="row-select-checkbox" type="checkbox" checked disabled /></td>
              <td>${escapeHtml(row.date)}</td>
              <td>${escapeHtml(row.startTime)}</td>
              <td>${escapeHtml(row.endTime)}</td>
              <td>${escapeHtml(row.therapist)}</td>
              <td>${escapeHtml(row.location)}</td>
              <td>${escapeHtml(row.procedure)}</td>
              <td>${escapeHtml(row.summary)}</td>
              <td></td>
            </tr>
          `).join("")}
        </tbody>
      </table>
    </div>
  ` : `
    <div class="table-wrap">
      <table>
        <thead>
          <tr>
            <th class="row-select-header">Select</th>
            <th>Date</th>
            <th>Start</th>
            <th>End</th>
            <th>Therapist / Session</th>
            <th>Location</th>
            <th>Procedure</th>
            <th>Calendar title</th>
            <th></th>
          </tr>
        </thead>
        <tbody>
          <tr class="empty-row">
            <td colspan="9">${escapeHtml(error ? "No sessions extracted." : "No sessions extracted yet.")}</td>
          </tr>
        </tbody>
      </table>
    </div>
  `;

  const statusKind = error ? "error" : hasRows ? "success" : extraction ? "processing" : "idle";
  const statusTitle = error
    ? "Extraction failed"
    : hasRows
      ? "Extraction complete"
      : extraction
        ? "No sessions found"
        : "Ready";
  const statusText = error
    ? error
    : hasRows
      ? `${extraction.rows.length} appointment rows were extracted from ${sourceName || "the uploaded file"}.`
      : extraction
        ? "OCR ran on the server, but no appointment rows were confidently parsed."
        : "Browser JavaScript is not available, so server-side extraction is enabled automatically.";

  const patientMarkup = `
    <div class="meta-grid noscript-meta-grid">
      <label>
        <span>Calendar title</span>
        <input type="text" name="calendarTitle" value="${escapeHtml(calendarTitleValue)}" />
      </label>
      <label>
        <span>Timezone</span>
        <input type="text" name="timezone" value="${escapeHtml(timezoneValue)}" />
      </label>
      <label>
        <span>Patient name</span>
        <input value="${escapeHtml(extraction?.metadata.patientName || "")}" readonly />
      </label>
      <label>
        <span>MRN</span>
        <input value="${escapeHtml(extraction?.metadata.mrn || "")}" readonly />
      </label>
    </div>
  `;

  return `<!doctype html>
<html lang="en" class="no-js">
  <head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <title>Diary converter (RH)</title>
    <link rel="preconnect" href="https://fonts.googleapis.com" />
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin />
    <link
      href="https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@400;500;700&family=IBM+Plex+Mono:wght@400;500&display=swap"
      rel="stylesheet"
    />
    <link rel="icon" type="image/png" href="/favicon.png" />
    <link rel="stylesheet" href="/styles.css" />
  </head>
  <body>
    <main class="shell">
      <section class="hero card">
        <div class="hero-copy">
          <h1>Diary converter (RH)</h1>
          <p class="lede">Simplify your rehab schedule process by transferring your PDF or image to calendar entries or automate the insert/import to your calendar.</p>
        </div>
      </section>

      <section class="workspace">
        <form class="workspace-form" action="/fallback/extract" method="post" enctype="multipart/form-data">
        <div class="left-column">
          <section class="card">
            <div class="section-head">
              <h2>1. Upload</h2>
            </div>

            <div class="audit-form">
              <input id="noscript-file-input" class="nojs-file-input" type="file" name="sourceFile" accept="image/png,image/jpeg,image/webp,image/bmp,application/pdf,.pdf" required />
              <label class="dropzone nojs-dropzone" for="noscript-file-input">
                <div class="dropzone-copy">
                  <span class="dropzone-title">Drop a schedule image or PDF here</span>
                  <span class="dropzone-subtitle">or click to browse PNG, JPG, WEBP, BMP, or PDF</span>
                </div>
                <div class="dropzone-preview">
                  <span class="nojs-placeholder-title">No file selected yet.</span>
                  <span id="image-placeholder">Click this box to choose the schedule file.</span>
                </div>
              </label>
              <div class="extract-options static">
                <span class="extract-options-label">Extraction mode</span>
                <label class="engine-option">
                  <input type="radio" name="engineDisplay" value="client" disabled />
                  <span>Client side</span>
                </label>
                <label class="engine-option">
                  <input type="radio" name="engineDisplay" value="server" checked />
                  <span>Server side</span>
                </label>
              </div>
              <div class="action-row">
                <button class="primary-button" type="submit">Extract schedule</button>
                <a class="ghost-link button-link" href="/">Refresh page</a>
              </div>
              <div class="status-panel ${statusKind}">
                <strong>${escapeHtml(statusTitle)}</strong>
                <p>${escapeHtml(statusText)}</p>
              </div>
            </div>
          </section>

          <section class="card">
            <div class="section-head">
              <h2>2. Calendar Settings</h2>
              <span class="pill">${escapeHtml(sourceName || "No file loaded")}</span>
            </div>
            ${patientMarkup}
          </section>
        </div>

        <div class="right-column">
          <section class="card">
            <div class="section-head">
              <h2>3. Review Extracted Sessions</h2>
            </div>
            ${reviewActionsMarkup}
            ${statsMarkup}
            ${tableMarkup}
          </section>

          <section class="card footer-card">
            <p class="audit-message">License &amp; powered by EDP Consulting.</p>
            <a class="ghost-link footer-link" href="mailto:enquiries@edp-consult.cn">Contact us: enquiries@edp-consult.cn</a>
          </section>
        </div>
        </form>
      </section>
    </main>
  </body>
</html>`;
}

async function handleAuditEventCreate(request, response) {
  try {
    const body = await collectRequestBody(request);
    const parsed = JSON.parse(body);
    const activity = String(parsed.activity ?? "").trim();
    const status = String(parsed.status ?? "info").trim();
    const details = typeof parsed.details === "object" && parsed.details !== null ? parsed.details : {};

    if (!activity) {
      sendJson(response, 400, { error: "Missing activity" });
      return;
    }

    appendAuditEntry(buildAuditEntry(activity, status, details, request));
    sendJson(response, 201, { ok: true });
  } catch (error) {
    const message = error instanceof Error ? error.message : "Invalid audit event request";
    sendJson(response, 400, { error: message });
  }
}

async function handleAdminLogin(request, response) {
  try {
    const body = await collectRequestBody(request);
    const parsed = JSON.parse(body);
    const username = String(parsed.username ?? "");
    const password = String(parsed.password ?? "");

    if (username !== adminAuth.username || password !== adminAuth.password) {
      appendAuditEntry(buildAuditEntry("admin_login", "failed", { username }, request));
      sendJson(response, 401, { error: "Invalid username or password" });
      return;
    }

    const token = randomUUID();
    sessionStore.set(token, {
      createdAt: Date.now(),
      username
    });

    appendAuditEntry(buildAuditEntry("admin_login", "success", { username }, request));
    sendJson(
      response,
      200,
      { ok: true, username },
      {
        "Set-Cookie": `audit_session=${encodeURIComponent(token)}; HttpOnly; Path=/; SameSite=Lax; Max-Age=${sessionTtlMs / 1000}`
      }
    );
  } catch (error) {
    const message = error instanceof Error ? error.message : "Invalid login request";
    sendJson(response, 400, { error: message });
  }
}

function handleAdminLogout(request, response) {
  const cookies = parseCookies(request);
  if (cookies.audit_session) {
    sessionStore.delete(cookies.audit_session);
  }
  appendAuditEntry(buildAuditEntry("admin_logout", "success", {}, request));
  sendJson(
    response,
    200,
    { ok: true },
    { "Set-Cookie": "audit_session=; HttpOnly; Path=/; SameSite=Lax; Max-Age=0" }
  );
}

function handleAdminStatus(request, response) {
  const session = getAdminSession(request);
  sendJson(response, 200, {
    authenticated: Boolean(session),
    username: session?.username ?? null
  });
}

function handleAuditLogRequest(request, response) {
  const session = getAdminSession(request);
  if (!session) {
    sendJson(response, 401, { error: "Authentication required" });
    return;
  }

  sendJson(response, 200, {
    entries: readAuditEntries(),
    username: session.username
  });
}

function getSortedValidRows(rows) {
  return [...rows]
    .filter((row) => row.date && row.startTime && row.endTime)
    .sort((left, right) =>
      `${left.date}T${left.startTime}`.localeCompare(`${right.date}T${right.startTime}`)
    );
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
  const uniqueSummaries = [...new Set(rows.map((row) => row.summary).filter(Boolean))];
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
        row.summary
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

function buildFileStem(metadata, sourceName) {
  const patientStem = metadata.patientName
    ? metadata.patientName.toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/^-+|-+$/g, "")
    : String(sourceName ?? "").replace(/\.[^.]+$/, "").toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/^-+|-+$/g, "");
  return `${patientStem || "royalrehab-schedule"}-export`;
}

function buildCsv(rows) {
  const header = ["date", "startTime", "endTime", "therapist", "location", "procedure", "summary", "mergedCount"];
  const lines = [header.join(",")];

  for (const row of rows) {
    lines.push(
      [row.date, row.startTime, row.endTime, row.therapist, row.location, row.procedure, row.summary, row.mergedCount]
        .map(csvEscape)
        .join(",")
    );
  }

  return lines.join("\r\n");
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
    lines.push(`SUMMARY:${escapeIcsText(row.summary)}`);
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

function csvEscape(value) {
  const safeValue = String(value ?? "");
  return `"${safeValue.replace(/"/g, "\"\"")}"`;
}

function escapeHtml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/"/g, "&quot;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

function warnIfUsingDefaultAdminPassword() {
  if (process.env.ADMIN_USERNAME || process.env.ADMIN_PASSWORD) {
    return;
  }

  console.log("Audit log admin login enabled with default credentials.");
  console.log(`Username: ${adminAuth.username}`);
  console.log(`Password: ${adminAuth.password}`);
  console.log("Set ADMIN_USERNAME and ADMIN_PASSWORD to change these defaults.");
}

const server = createServer(async (request, response) => {
  const url = new URL(request.url ?? "/", `http://${request.headers.host ?? `${host}:${port}`}`);
  const pathname = decodeURIComponent(url.pathname);
  const isHead = request.method === "HEAD";
  const isGetLike = request.method === "GET" || isHead;

  if (isGetLike && pathname === "/health") {
    sendJson(response, 200, { ok: true }, {}, { head: isHead });
    return;
  }

  if (isGetLike && pathname === "/api/google-config") {
    sendJson(response, 200, {
      enabled: Boolean(googleCalendarConfig.clientId),
      apiKey: googleCalendarConfig.apiKey,
      clientId: googleCalendarConfig.clientId
    }, {}, { head: isHead });
    return;
  }

  if (isGetLike && pathname === "/api/server-ocr-status") {
    const status = await getServerOcrStatus();
    sendJson(response, 200, status, {}, { head: isHead });
    return;
  }

  if (request.method === "POST" && pathname === "/api/server-ocr-jobs") {
    await handleServerOcrJobCreate(request, response);
    return;
  }

  if (isGetLike && pathname.startsWith("/api/server-ocr-jobs/")) {
    const jobId = pathname.slice("/api/server-ocr-jobs/".length);
    handleServerOcrJobStatus(response, jobId);
    return;
  }

  if (request.method === "POST" && pathname === "/api/server-ocr") {
    await handleServerOcrRequest(request, response);
    return;
  }

  if (request.method === "POST" && pathname === "/fallback/extract") {
    await handleServerOcrRequest(request, response, { html: true });
    return;
  }

  if (isGetLike && pathname === "/fallback") {
    sendHtml(response, 200, renderNoScriptResultPage({}), { head: isHead });
    return;
  }

  if (isGetLike && pathname === "/api/admin/status") {
    handleAdminStatus(request, response);
    return;
  }

  if (request.method === "POST" && pathname === "/api/admin/login") {
    await handleAdminLogin(request, response);
    return;
  }

  if (request.method === "POST" && pathname === "/api/admin/logout") {
    handleAdminLogout(request, response);
    return;
  }

  if (isGetLike && pathname === "/api/auditlog") {
    handleAuditLogRequest(request, response);
    return;
  }

  if (request.method === "POST" && pathname === "/api/audit-events") {
    await handleAuditEventCreate(request, response);
    return;
  }

  if (request.method === "POST" && pathname === "/api/exports") {
    await handleExportCreate(request, response);
    return;
  }

  if (isGetLike && pathname.startsWith("/downloads/")) {
    const exportId = pathname.slice("/downloads/".length);
    handleExportDownload(response, exportId, { head: isHead });
    return;
  }

  if (isGetLike && pathname === "/auditlogin") {
    sendFile(response, join(publicDir, "audit-login.html"), { head: isHead });
    return;
  }

  if (isGetLike && pathname === "/auditlog") {
    if (!getAdminSession(request)) {
      redirect(response, "/auditlogin");
      return;
    }
    sendFile(response, join(publicDir, "auditlog.html"), { head: isHead });
    return;
  }

  if (isGetLike && pathname === "/") {
    sendFile(response, join(publicDir, "index.html"), { head: isHead });
    return;
  }

  if (isGetLike && pathname === "/sample/Sample-schedule.png") {
    sendFile(response, join(rootDir, "Sample-schedule.png"), { head: isHead });
    return;
  }

  if (!isGetLike) {
    response.writeHead(405, { "Content-Type": "text/plain; charset=utf-8" });
    response.end("Method not allowed");
    return;
  }

  const requestedFile = safeJoin(publicDir, pathname.slice(1));
  if (!requestedFile) {
    response.writeHead(400, { "Content-Type": "text/plain; charset=utf-8" });
    response.end("Bad request");
    return;
  }

  sendFile(response, requestedFile, { head: isHead });
});

server.listen(port, host, () => {
  console.log(`Royal Rehab schedule converter running at http://${host}:${port}`);
  warnIfUsingDefaultAdminPassword();
});
