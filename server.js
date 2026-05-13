import { createReadStream, existsSync, statSync } from "node:fs";
import { extname, join, normalize } from "node:path";
import { createServer } from "node:http";
import { randomUUID } from "node:crypto";

const host = "127.0.0.1";
const port = Number.parseInt(process.env.PORT ?? "3000", 10);
const rootDir = process.cwd();
const publicDir = join(rootDir, "public");
const exportStore = new Map();
const exportTtlMs = 10 * 60 * 1000;
const googleCalendarConfig = {
  apiKey: process.env.GOOGLE_API_KEY ?? "",
  clientId: process.env.GOOGLE_CLIENT_ID ?? ""
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

function sendJson(response, statusCode, payload) {
  response.writeHead(statusCode, { "Content-Type": "application/json; charset=utf-8" });
  response.end(JSON.stringify(payload));
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

    sendJson(response, 201, {
      downloadUrl: `/downloads/${id}`,
      id
    });
  } catch (error) {
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

const server = createServer(async (request, response) => {
  const url = new URL(request.url ?? "/", `http://${request.headers.host ?? `${host}:${port}`}`);
  const pathname = decodeURIComponent(url.pathname);

  if (request.method === "GET" && pathname === "/health") {
    sendJson(response, 200, { ok: true });
    return;
  }

  if (request.method === "GET" && pathname === "/api/google-config") {
    sendJson(response, 200, {
      enabled: Boolean(googleCalendarConfig.apiKey && googleCalendarConfig.clientId),
      apiKey: googleCalendarConfig.apiKey,
      clientId: googleCalendarConfig.clientId
    });
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
});
