// Local dev server: adapts Node's http module to the shared fetch-style proxy
// in src/proxy.js and serves the skin straight from public/ and the Fluent UI
// icon package, so edits show up on refresh.

import http from "node:http";
import { readFile } from "node:fs/promises";
import { extname, join, normalize } from "node:path";
import { fileURLToPath } from "node:url";
import { handle, SKIN_PREFIX } from "./src/proxy.js";
import { iconFile } from "./src/icons.js";

const HOST = "127.0.0.1";
const PORT = Number(process.env.PORT) || 4322;
const PUBLIC_DIR = fileURLToPath(new URL("./public/", import.meta.url));
const ICON_PREFIX = `${SKIN_PREFIX}icons/`;
const ICON_DIR = fileURLToPath(new URL("./node_modules/@fluentui/svg-icons/icons/", import.meta.url));

const MIME = {
  ".css": "text/css; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".svg": "image/svg+xml",
  ".png": "image/png",
  ".ico": "image/x-icon",
};

function resolveSkinFile(pathname) {
  if (pathname.startsWith(ICON_PREFIX)) {
    const name = pathname.slice(ICON_PREFIX.length).replace(/\.svg$/, "");
    return /^[a-z0-9_]+$/.test(name) ? join(ICON_DIR, iconFile(name)) : null;
  }
  const rel = normalize(pathname.slice(SKIN_PREFIX.length));
  return rel.startsWith("..") ? null : join(PUBLIC_DIR, rel);
}

async function serveSkin(pathname) {
  const file = resolveSkinFile(pathname);
  if (!file) return new Response("Forbidden", { status: 403 });
  try {
    return new Response(await readFile(file), {
      headers: {
        "content-type": MIME[extname(file)] || "application/octet-stream",
        "cache-control": file.startsWith(ICON_DIR) ? "public, max-age=86400" : "no-store",
      },
    });
  } catch {
    return new Response("Not found", { status: 404 });
  }
}

function readBody(req) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    req.on("data", (c) => chunks.push(c));
    req.on("end", () => resolve(Buffer.concat(chunks)));
    req.on("error", reject);
  });
}

async function onRequest(req, res) {
  const headers = new Headers();
  for (const [k, v] of Object.entries(req.headers)) {
    if (k === "host" || k === "connection") continue;
    headers.set(k, Array.isArray(v) ? v.join(", ") : v);
  }
  const hasBody = !["GET", "HEAD"].includes(req.method);
  const request = new Request(`http://${req.headers.host}${req.url}`, {
    method: req.method,
    headers,
    body: hasBody ? await readBody(req) : undefined,
  });

  const response = await handle(request, { skin: serveSkin });

  const out = {};
  response.headers.forEach((v, k) => {
    if (k !== "set-cookie") out[k] = v;
  });
  const cookies = response.headers.getSetCookie();
  if (cookies.length) out["set-cookie"] = cookies;
  res.writeHead(response.status, out);
  res.end(response.body ? Buffer.from(await response.arrayBuffer()) : undefined);
}

const server = http.createServer((req, res) => {
  onRequest(req, res).catch((err) => {
    console.error(`✗ ${req.method} ${req.url}: ${err.message}`);
    if (!res.headersSent) res.writeHead(502, { "content-type": "text/plain" });
    res.end(`Upstream error: ${err.message}`);
  });
});

server.listen(PORT, HOST, () => {
  console.log(`Krillion (Excel edition) → http://localhost:${PORT}`);
});
