// Builds dist/ for GitHub Pages: the skin files plus only the Fluent UI icons
// that excel.js references. The Worker serves these under /__excel/.

import { cp, mkdir, readFile, rm, writeFile, access } from "node:fs/promises";
import { fileURLToPath } from "node:url";
import { join } from "node:path";
import { iconFile } from "../src/icons.js";

const root = fileURLToPath(new URL("..", import.meta.url));
const dist = join(root, "dist");
const skinDir = join(dist, "__excel");
const iconSrc = join(root, "node_modules/@fluentui/svg-icons/icons");

await rm(dist, { recursive: true, force: true });
await mkdir(join(skinDir, "icons"), { recursive: true });
await cp(join(root, "public"), skinDir, { recursive: true });

// Any quoted snake_case word in excel.js that names an icon file is an icon.
const js = await readFile(join(root, "public/excel.js"), "utf8");
const names = [...new Set(js.match(/(?<=["'])[a-z][a-z0-9_]*(?=["'])/g))];
let count = 0;
for (const name of names) {
  const file = join(iconSrc, iconFile(name));
  try {
    await access(file);
  } catch {
    continue;
  }
  await cp(file, join(skinDir, "icons", `${name}.svg`));
  count++;
}

await writeFile(join(dist, ".nojekyll"), "");
await writeFile(
  join(dist, "index.html"),
  `<!doctype html><meta charset="utf-8"><title>Krillion, Excel edition</title>
<meta http-equiv="refresh" content="0; url=https://krillion.tomnewson.com/">
<p>This is the asset host for <a href="https://krillion.tomnewson.com/">krillion.tomnewson.com</a>.</p>\n`,
);
console.log(`Built dist/ with ${count} icons`);
