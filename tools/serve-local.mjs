import fs from "node:fs";
import http from "node:http";
import path from "node:path";
import { fileURLToPath } from "node:url";

const here = path.dirname(fileURLToPath(import.meta.url));
const root = path.resolve(here, "..");
const port = Number(process.env.PORT || 4176);
const host = "127.0.0.1";

const mimeTypes = {
  ".css": "text/css; charset=utf-8",
  ".html": "text/html; charset=utf-8",
  ".ico": "image/x-icon",
  ".js": "text/javascript; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".map": "application/json; charset=utf-8",
  ".pdf": "application/pdf",
  ".svg": "image/svg+xml",
  ".zip": "application/zip"
};

const server = http.createServer((request, response) => {
  const url = new URL(request.url || "/", `http://${host}:${port}`);
  let pathname = decodeURIComponent(url.pathname);
  if (pathname.endsWith("/")) {
    pathname += "index.html";
  }

  const filePath = path.resolve(root, "." + pathname);
  if (!filePath.startsWith(root + path.sep)) {
    send(response, 403, "Forbidden");
    return;
  }

  fs.readFile(filePath, (error, data) => {
    if (error) {
      send(response, 404, "Not found");
      return;
    }

    const extension = path.extname(filePath).toLowerCase();
    response.writeHead(200, headers(mimeTypes[extension] || "application/octet-stream"));
    response.end(data);
  });
});

server.listen(port, host, () => {
  console.log(`PDF Freely local server: http://${host}:${port}/`);
});

function send(response, status, message) {
  response.writeHead(status, headers("text/plain; charset=utf-8"));
  response.end(message);
}

function headers(contentType) {
  return {
    "Content-Type": contentType,
    "Cross-Origin-Resource-Policy": "same-origin",
    "X-Content-Type-Options": "nosniff",
    "Referrer-Policy": "strict-origin-when-cross-origin",
    "Permissions-Policy": "camera=(), microphone=(), geolocation=(), payment=()"
  };
}
