// Cloudflare Worker for krillion.tomnewson.com: runs the shared proxy and loads
// the skin (CSS, JS, icons) from the GitHub Pages build of this repo.

import { handle } from "./src/proxy.js";

export default {
  async fetch(request, env) {
    const skin = async (pathname) => {
      const res = await fetch(`${env.SKIN_ORIGIN}${pathname}`, { cf: { cacheTtl: 300, cacheEverything: true } });
      const out = new Response(res.body, res);
      out.headers.set("cache-control", "public, max-age=300");
      return out;
    };
    try {
      return await handle(request, { skin, noindex: true });
    } catch (err) {
      return new Response(`Upstream error: ${err.message}`, { status: 502 });
    }
  },
};
