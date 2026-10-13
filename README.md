# Krillion, Excel edition

An Excel-style front end for [krillion.io](https://krillion.io), live at **https://krillion.tomnewson.com**.

```bash
npm install
npm start          # local dev → http://localhost:4322  (PORT=xxxx to change)
```

## How it works

`src/proxy.js` is a small reverse proxy written against the standard fetch API. It forwards every page, asset and `/api/*` call to krillion.io. On HTML responses it injects `public/excel.css` and `public/excel.js`. The real game still does all the work, so every feature behaves as it does on the live site: the daily dive, Unlimited, themed packs, answer corrections and rejections, results, every answer for each round, streaks and stats, sharing, settings, the wardrobe and the dive log.

- **excel.js** draws the Excel chrome around the game. That covers the title bar, ribbon, formula bar, row and column headings, sheet tabs and status bar. The chrome's controls click the game's own buttons and links.
- **excel.css** restyles the game as worksheet content. The full-screen pixel-art ocean is switched off and the game is laid out down column B: the HUD becomes a table, the prompt a heading cell, the answer box the active cell with a countdown cell beside it, and each landed answer a row with conditional formatting by tier. The results screen becomes a report with a column chart, and the game's dialogs look like Excel dialogs.
- `server.js` runs the proxy on Node for local development and serves the skin straight from `public/`.
- `worker.js` runs the same proxy on Cloudflare Workers for krillion.tomnewson.com and loads the skin from GitHub Pages.
- Icons come from Microsoft's Fluent UI set (`@fluentui/svg-icons`) and are served at `/__excel/icons/<name>.svg`.

Ad slots are hidden by the skin because they'd break the disguise. The cookie consent banner still appears once per browser.

## Deployment

The site has two parts:

- **Skin (GitHub Pages).** Every push to `main` runs `.github/workflows/pages.yml`. It builds `dist/` with `npm run build`: the skin files plus only the icons that are used. It publishes to https://tomnewson.com/krillion/.
- **Proxy (Cloudflare Worker).** `npm run deploy:worker` deploys `worker.js` to krillion.tomnewson.com using the custom-domain route in `wrangler.toml`. Run `npx wrangler login` once first. Only redeploy when `src/proxy.js` or `worker.js` changes; skin changes just need a push.

The Worker caches skin files for 5 minutes, so skin changes show up shortly after the Pages deploy finishes. Pages are served with `x-robots-tag: noindex` so the mirror stays out of search results.

## Using it

| Excel control | What it does |
| --- | --- |
| Type anywhere / formula bar | Enters your answer (Enter or ✓ submits) |
| Enter (when no answer box is open) | Begin descent / Descend / Surface |
| Home › Run, Fill, Formulas › Insert Function | Same as Enter: the game's next step |
| Home › Copy, Share | Copies or shares your result |
| Conditional Formatting, Analyze Data, Comments | Jumps to every answer for each round |
| Find & Select, Help | How to play |
| Sort & Filter, Data › Get Data | Themed packs |
| Format, Help › Settings | The game's settings |
| Format Painter, Cell Styles, Themes | The wardrobe (dive buddies) |
| Formulas › Trace Precedents, Data › Queries | The dive log on Unlimited and pack sheets |
| View › Sound / Navigation | Sound on/off, and the game's own menu |
| Sheet tabs | Today, Unlimited, Packs, Archive, My Krillion, Friends, FAQ |
| File | Navigation, account, sharing, options and help |
| View › Gridlines / Headings / Zoom | Work as they do in Excel (saved per browser) |
| Esc Esc, View › Switch Windows | Switches to a boring expense sheet (and back) |
| Status bar | Average / Count / Sum of your round scores |
| Name box | Type a cell reference like `D12` to jump there |

To change the file name shown in the title bar and browser tab, edit `WORKBOOK` at the top of `public/excel.js`.

## Signing in

Krillion signs in with Google through Firebase, which only accepts sign-ins from domains on the project's allow-list, so **Continue with Google** won't complete on krillion.tomnewson.com. Progress is kept in the browser instead: this site (and localhost) each have their own anonymous diver, with their own streak and stats. Purchases can be moved over with **Restore**.
