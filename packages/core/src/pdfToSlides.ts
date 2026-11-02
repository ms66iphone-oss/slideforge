// Converts a PDF resource into a small, self-contained .pptx — one slide per
// page, each page rendered to a full-bleed image — so it can travel through
// the same embedding path already used for docx/xlsx/pptx resources. A PDF
// isn't an OOXML zip package the way those are, so it can't be embedded
// directly (see oleEmbed.ts); this sidesteps that by turning it into a real
// .pptx first, which the existing embedding code already knows how to embed
// and which PowerPoint already knows how to open — on Mac and Windows alike,
// unlike a direct PDF embed via legacy OLE, which Mac PowerPoint doesn't
// support creating at all.
//
// Rendering a page needs a real 2D canvas context. The browser tab supplies
// one; jsdom (the test environment) does not, so both the PDF.js build and
// the canvas factory are injectable — tests supply the Node-targeted
// equivalents directly rather than this module trying to detect its
// environment.

import PptxGenJS from "pptxgenjs";
// eslint-disable-next-line import/no-unresolved -- pdfjs-dist ships no root type declarations; the matching ambient module below covers what we use.
import * as browserPdfjsLib from "pdfjs-dist";

export interface PdfPageViewport {
  width: number;
  height: number;
}

export interface PdfPage {
  getViewport(params: { scale: number }): PdfPageViewport;
  render(params: { canvasContext: unknown; viewport: PdfPageViewport; canvas?: unknown }): {
    promise: Promise<void>;
  };
}

export interface PdfDocument {
  numPages: number;
  getPage(pageNumber: number): Promise<PdfPage>;
}

export interface PdfjsLibLike {
  getDocument(params: { data: Uint8Array }): { promise: Promise<PdfDocument> };
}

export interface CanvasLike {
  width: number;
  height: number;
  getContext(type: "2d"): unknown;
  toDataURL(type?: string): string;
}

export type CanvasFactory = (width: number, height: number) => CanvasLike;

function browserCanvasFactory(width: number, height: number): CanvasLike {
  const canvas = document.createElement("canvas");
  canvas.width = width;
  canvas.height = height;
  return canvas as unknown as CanvasLike;
}

export interface ConvertPdfDeps {
  pdfjsLib?: PdfjsLibLike;
  canvasFactory?: CanvasFactory;
}

export interface PdfSlidesResult {
  dataUrl: string;
  pageCount: number;
}

const POINTS_PER_INCH = 72;
const RENDER_DPI = 150;
// PowerPoint's own slide-size bounds (1in–56in on each edge) — a corrupt or
// unusual PDF page shouldn't be able to produce a layout PowerPoint refuses.
const MIN_SLIDE_IN = 1;
const MAX_SLIDE_IN = 56;

export function dataUrlToUint8Array(dataUrl: string): Uint8Array {
  const comma = dataUrl.indexOf(",");
  const base64 = comma === -1 ? dataUrl : dataUrl.slice(comma + 1);
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
  return bytes;
}

export function arrayBufferToBase64(buffer: ArrayBuffer): string {
  const bytes = new Uint8Array(buffer);
  const chunkSize = 0x8000;
  let binary = "";
  for (let i = 0; i < bytes.length; i += chunkSize) {
    binary += String.fromCharCode(...bytes.subarray(i, i + chunkSize));
  }
  return btoa(binary);
}

function clampDimension(inches: number): number {
  if (!Number.isFinite(inches) || inches <= 0) return MIN_SLIDE_IN;
  return Math.min(MAX_SLIDE_IN, Math.max(MIN_SLIDE_IN, inches));
}

/** Renders every page of a PDF to an image and assembles them into a new
 * .pptx, one page per slide, sized to the PDF's own first-page aspect ratio.
 * A later page with a different aspect ratio (a one-off landscape page in an
 * otherwise portrait worksheet) is fitted within that slide rather than
 * stretched, so nothing distorts.
 *
 * Returns null on anything that stops a clean conversion — a corrupt file, a
 * zero-page document, a render failure — so the caller can fall back to the
 * plain resource card instead of shipping a half-built file. */
export async function convertPdfToSlidesPptx(
  pdfDataUrl: string,
  deps: ConvertPdfDeps = {},
): Promise<PdfSlidesResult | null> {
  const pdfjsLib = deps.pdfjsLib ?? (browserPdfjsLib as unknown as PdfjsLibLike);
  const canvasFactory = deps.canvasFactory ?? browserCanvasFactory;

  let doc: PdfDocument;
  try {
    const data = dataUrlToUint8Array(pdfDataUrl);
    doc = await pdfjsLib.getDocument({ data }).promise;
  } catch {
    return null;
  }

  if (!doc.numPages || doc.numPages < 1) return null;

  let slideW = 10;
  let slideH = 7.5;
  try {
    const firstPage = await doc.getPage(1);
    const unscaled = firstPage.getViewport({ scale: 1 });
    slideW = clampDimension(unscaled.width / POINTS_PER_INCH);
    slideH = clampDimension(unscaled.height / POINTS_PER_INCH);
  } catch {
    return null;
  }

  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: "SF_PDF_PAGE", width: slideW, height: slideH });
  pptx.layout = "SF_PDF_PAGE";

  try {
    for (let i = 1; i <= doc.numPages; i++) {
      const page = await doc.getPage(i);
      const scale = RENDER_DPI / POINTS_PER_INCH;
      const viewport = page.getViewport({ scale });
      const pageW = viewport.width / RENDER_DPI;
      const pageH = viewport.height / RENDER_DPI;

      const canvas = canvasFactory(Math.ceil(viewport.width), Math.ceil(viewport.height));
      const context = canvas.getContext("2d");
      await page.render({ canvasContext: context, viewport, canvas }).promise;
      const imageDataUrl = canvas.toDataURL("image/png");

      // Fit (not stretch) within the slide, centred, in case this page's
      // aspect ratio differs from the first page that set the slide size.
      const fitScale = Math.min(slideW / pageW, slideH / pageH);
      const w = pageW * fitScale;
      const h = pageH * fitScale;
      const x = (slideW - w) / 2;
      const y = (slideH - h) / 2;

      const slide = pptx.addSlide();
      slide.addImage({ data: imageDataUrl, x, y, w, h });
    }
  } catch {
    return null;
  }

  const buffer = (await pptx.write({ outputType: "arraybuffer" })) as ArrayBuffer;
  const base64 = arrayBufferToBase64(buffer);
  return {
    dataUrl: `data:application/vnd.openxmlformats-officedocument.presentationml.presentation;base64,${base64}`,
    pageCount: doc.numPages,
  };
}
