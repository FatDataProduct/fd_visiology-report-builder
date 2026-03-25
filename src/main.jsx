import { StrictMode } from "react";
import { createRoot } from "react-dom/client";
import App from "../visiology-report-builder.jsx";

createRoot(document.getElementById("root")).render(
  <StrictMode>
    <App />
  </StrictMode>
);
