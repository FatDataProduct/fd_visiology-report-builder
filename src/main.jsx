import { StrictMode } from "react";
import { hydrateRoot, createRoot } from "react-dom/client";
import App from "../visiology-report-builder.jsx";

const container = document.getElementById("root");

/**
 * If the container has SSR-rendered HTML (pre-render was run during build),
 * use hydrateRoot so React attaches to the existing DOM without a full repaint.
 * Otherwise fall back to a normal createRoot mount (dev / no-prerender builds).
 */
if (container.innerHTML.trim()) {
  hydrateRoot(
    container,
    <StrictMode>
      <App />
    </StrictMode>
  );
} else {
  createRoot(container).render(
    <StrictMode>
      <App />
    </StrictMode>
  );
}
