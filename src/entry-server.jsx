/**
 * Server-side entry point.
 * Used by scripts/prerender.mjs at build time to generate
 * the static HTML snapshot injected into dist/index.html.
 */
import { renderToString } from "react-dom/server";
import { StrictMode } from "react";
import App from "../visiology-report-builder.jsx";

/**
 * Render the full React tree to an HTML string.
 * useEffect / browser APIs are skipped automatically during renderToString.
 */
export function render() {
  return renderToString(
    <StrictMode>
      <App />
    </StrictMode>
  );
}
