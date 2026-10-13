// Reverse proxy for krillion.io that wraps the real app in the Excel skin.
// Written against the web-standard fetch API so the same code runs in the
// local Node server (server.js) and on Cloudflare Workers (worker.js).
//
// Every page, asset and API call is forwarded upstream, so all of the game's
// features keep working; HTML responses get the skin's stylesheet and script
// injected. Requests under /__excel/ are the skin itself and are served by the
// caller-supplied `skin` function.

export const UPSTREAM = "https://krillion.io";
export const SKIN_PREFIX = "/__excel/";

// Headers that must not be forwarded as-is in either direction.
const HOP_BY_HOP = new Set([
  "connection",
  "keep-alive",
  "transfer-encoding",
  "upgrade",
  "proxy-connection",
  "te",
  "trailer",
]);

// Request headers that describe the client's hop to us, not ours upstream.
const DROP_REQUEST = new Set(["host", "accept-encoding", "content-length", "x-real-ip", "x-forwarded-for", "x-forwarded-proto", "x-forwarded-host"]);

// Upstream security headers that would break the proxied, injected page on
// another origin (CSP pins 'self', HSTS, etc.). Cookies are rewritten apart.
const DROP_RESPONSE = new Set([
  "content-encoding",
  "content-length",
  "content-security-policy",
  "strict-transport-security",
  "x-frame-options",
  "set-cookie",
  "alt-svc",
]);

const NULL_BODY_STATUS = new Set([101, 204, 205, 304]);

const INJECT = `
<link rel="stylesheet" href="${SKIN_PREFIX}excel.css">
<link rel="icon" href="${SKIN_PREFIX}favicon.svg" type="image/svg+xml">
<script src="${SKIN_PREFIX}excel.js" defer></script>
`;

function rewriteCookie(cookie, secure) {
  return cookie
    .split(";")
    .map((p) => p.trim())
    .filter((p) => !/^domain=/i.test(p) && (secure || !/^secure$/i.test(p)))
    .map((p) => (!secure && /^samesite=none$/i.test(p) ? "SameSite=Lax" : p))
    .join("; ");
}

/**
 * @param {Request} request
 * @param {{ skin: (pathname: string) => Promise<Response>, noindex?: boolean }} opts
 */
export async function handle(request, { skin, noindex = false }) {
  const url = new URL(request.url);
  if (url.pathname.startsWith(SKIN_PREFIX)) return skin(url.pathname);

  const target = new URL(url.pathname + url.search, UPSTREAM);
  const headers = new Headers();
  for (const [k, v] of request.headers) {
    if (HOP_BY_HOP.has(k) || DROP_REQUEST.has(k) || k.startsWith("cf-")) continue;
    headers.set(k, v);
  }
  if (headers.has("origin")) headers.set("origin", UPSTREAM);
  if (headers.has("referer")) headers.set("referer", headers.get("referer").replace(url.origin, UPSTREAM));

  const hasBody = !["GET", "HEAD"].includes(request.method);
  const upstream = await fetch(target, {
    method: request.method,
    headers,
    body: hasBody ? await request.arrayBuffer() : undefined,
    redirect: "manual",
  });

  const out = new Headers();
  upstream.headers.forEach((v, k) => {
    if (!HOP_BY_HOP.has(k) && !DROP_RESPONSE.has(k)) out.set(k, v);
  });
  const location = out.get("location");
  if (location) out.set("location", location.replace(/^https:\/\/(www\.)?krillion\.io/, url.origin));
  const secure = url.protocol === "https:";
  for (const c of upstream.headers.getSetCookie?.() ?? []) out.append("set-cookie", rewriteCookie(c, secure));
  if (noindex) out.set("x-robots-tag", "noindex, nofollow");

  if ((upstream.headers.get("content-type") || "").includes("text/html")) {
    let html = await upstream.text();
    html = html.includes("</head>") ? html.replace("</head>", `${INJECT}</head>`) : INJECT + html;
    out.set("cache-control", "no-store");
    return new Response(html, { status: upstream.status, headers: out });
  }

  const body = NULL_BODY_STATUS.has(upstream.status) || request.method === "HEAD" ? null : upstream.body;
  return new Response(body, { status: upstream.status, headers: out });
}
