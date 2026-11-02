import { GlobalWorkerOptions } from "pdfjs-dist";
import pdfWorkerUrl from "pdfjs-dist/build/pdf.worker.mjs?url";
import { StrictMode } from "react";
import { createRoot } from "react-dom/client";

import { App } from "@/App";

import "@/index.css";

// Lets PDF resource conversion (packages/core/src/pdfToSlides.ts) run off
// the main thread. Bundled locally rather than left to pdf.js's own CDN
// fallback, so nothing about forging a lesson depends on reaching the
// internet — everything here runs client-side regardless of network.
GlobalWorkerOptions.workerSrc = pdfWorkerUrl;

const rootElement = document.getElementById("root");
if (!rootElement) throw new Error("Missing #root element");

createRoot(rootElement).render(
  <StrictMode>
    <App />
  </StrictMode>,
);
