import { appendFile, createReadStream, existsSync, readFileSync, statSync } from "node:fs";
import { randomUUID } from "node:crypto";
import { createServer } from "node:http";
import { extname, join, normalize } from "node:path";

loadEnvFile(".env.local");
loadEnvFile(".env");

const host = process.env.HOST ?? "0.0.0.0";
const port = Number.parseInt(process.env.PORT ?? "3000", 10);
const rootDir = process.cwd();
const publicDir = join(rootDir, "public");
const auditLogPath = join(rootDir, "audit.log.jsonl");
const exportStore = new Map();
const sessionStore = new Map();
const exportTtlMs = 10 * 60 * 1000;
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

function sendFile(response, filePath) {
  if (!existsSync(filePath) || !statSync(filePath).isFile()) {
    response.writeHead(404, { "Content-Type": "text/plain; charset=utf-8" });
    response.end("Not found");
    return;
  }

  response.writeHead(200, {
    "Content-Type": mimeTypes[extname(filePath).toLowerCase()] ?? "application/octet-stream",
    "Cache-Control": "no-store"
  });

  createReadStream(filePath).pipe(response);
}

function sendJson(response, statusCode, payload, headers = {}) {
  response.writeHead(statusCode, {
    "Content-Type": "application/json; charset=utf-8",
    ...headers
  });
  response.end(JSON.stringify(payload));
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

function cleanupExpiredExports() {
  const cutoff = Date.now() - exportTtlMs;
  for (const [id, entry] of exportStore.entries()) {
    if (entry.createdAt < cutoff) {
      exportStore.delete(id);
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

    cleanupExpiredExports();

    const id = randomUUID();
    exportStore.set(id, {
      content,
      createdAt: Date.now(),
      filename,
      mimeType
    });

    appendAuditEntry(buildAuditEntry(
      "export_generated",
      "success",
      { filename, mimeType, sizeBytes: Buffer.byteLength(content, "utf8") },
      request
    ));

    sendJson(response, 201, {
      downloadUrl: `/downloads/${id}`,
      id
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

function handleExportDownload(response, exportId) {
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
  response.end(entry.content);
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

  if (request.method === "GET" && pathname === "/health") {
    sendJson(response, 200, { ok: true });
    return;
  }

  if (request.method === "GET" && pathname === "/api/google-config") {
    sendJson(response, 200, {
      enabled: Boolean(googleCalendarConfig.clientId),
      apiKey: googleCalendarConfig.apiKey,
      clientId: googleCalendarConfig.clientId
    });
    return;
  }

  if (request.method === "GET" && pathname === "/api/admin/status") {
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

  if (request.method === "GET" && pathname === "/api/auditlog") {
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

  if (request.method === "GET" && pathname.startsWith("/downloads/")) {
    const exportId = pathname.slice("/downloads/".length);
    handleExportDownload(response, exportId);
    return;
  }

  if (request.method === "GET" && pathname === "/auditlogin") {
    sendFile(response, join(publicDir, "audit-login.html"));
    return;
  }

  if (request.method === "GET" && pathname === "/auditlog") {
    if (!getAdminSession(request)) {
      redirect(response, "/auditlogin");
      return;
    }
    sendFile(response, join(publicDir, "auditlog.html"));
    return;
  }

  if (request.method === "GET" && pathname === "/") {
    sendFile(response, join(publicDir, "index.html"));
    return;
  }

  if (request.method === "GET" && pathname === "/sample/Sample-schedule.png") {
    sendFile(response, join(rootDir, "Sample-schedule.png"));
    return;
  }

  if (request.method !== "GET") {
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

  sendFile(response, requestedFile);
});

server.listen(port, host, () => {
  console.log(`Royal Rehab schedule converter running at http://${host}:${port}`);
  warnIfUsingDefaultAdminPassword();
});
