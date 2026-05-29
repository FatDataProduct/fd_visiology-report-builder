/**
 * Build-time prerender script (SSR snapshot).
 *
 * Run order (see package.json "build" script):
 *   1. vite build              → dist/         (client bundle + index.html)
 *   2. vite build --ssr ...    → dist/server/  (Node-compatible server bundle)
 *   3. node scripts/prerender  → patches dist/index.html with rendered HTML
 *
 * The patched index.html is what nginx serves.
 * React on the client then hydrateRoot() onto the existing DOM,
 * giving instant First Contentful Paint and correct SEO markup.
 */

import { readFileSync, writeFileSync } from "node:fs";
import { resolve, dirname } from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";

const __dirname = dirname(fileURLToPath(import.meta.url));
const root = resolve(__dirname, "..");

async function main() {
  // ── 1. Import server bundle ──────────────────────────────────────────────
  const serverEntry = resolve(root, "dist/server/entry-server.js");
  const { render } = await import(pathToFileURL(serverEntry).href);

  // ── 2. Read the Vite-built client template ───────────────────────────────
  const templatePath = resolve(root, "dist/index.html");
  const template = readFileSync(templatePath, "utf-8");

  // ── 3. Render the React tree to an HTML string ───────────────────────────
  const appHtml = render();

  // ── 4. Inject rendered HTML into the SSR outlet placeholder ─────────────
  if (!template.includes("<!--ssr-outlet-->")) {
    console.warn(
      "⚠  <!--ssr-outlet--> placeholder not found in dist/index.html — skipping injection."
    );
    return;
  }

  const html = template.replace("<!--ssr-outlet-->", appHtml);

  // ── 5. Overwrite dist/index.html ─────────────────────────────────────────
  writeFileSync(templatePath, html, "utf-8");

  const kbSaved = ((appHtml.length / 1024)).toFixed(1);
  console.log(`✅ Prerender complete — injected ${kbSaved} KB of HTML into dist/index.html`);
}

main().catch((err) => {
  console.error("❌ Prerender failed:", err);
  process.exit(1);
});
