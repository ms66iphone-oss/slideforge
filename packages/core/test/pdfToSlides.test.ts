// Exercises the real PDF.js + canvas rendering path, not just the plumbing
// around it — jsdom has no working 2D canvas context, so these tests supply
// the Node-targeted PDF.js build and a real (non-DOM) canvas implementation
// explicitly, the same way production code supplies the browser versions by
// default. See pdfToSlides.ts's module comment.

import { createCanvas } from "@napi-rs/canvas";
import JSZip from "jszip";
import * as pdfjsLib from "pdfjs-dist/legacy/build/pdf.mjs";
import { describe, expect, it } from "vitest";

import {
  arrayBufferToBase64,
  convertPdfToSlidesPptx,
  dataUrlToUint8Array,
} from "../src/pdfToSlides.js";
import type { CanvasLike, PdfjsLibLike } from "../src/pdfToSlides.js";
import { buildSyntheticPdfDataUrl } from "./pdfFixture.js";

const nodeCanvasFactory = (width: number, height: number): CanvasLike =>
  createCanvas(width, height) as unknown as CanvasLike;
const nodePdfjsLib = pdfjsLib as unknown as PdfjsLibLike;
const deps = { pdfjsLib: nodePdfjsLib, canvasFactory: nodeCanvasFactory };

describe("dataUrlToUint8Array / arrayBufferToBase64", () => {
  it("round-trips bytes through a data: URL", () => {
    const original = new Uint8Array([0, 1, 2, 250, 251, 252, 253, 254, 255]);
    const base64 = arrayBufferToBase64(original.buffer as ArrayBuffer);
    const result = dataUrlToUint8Array(`data:application/octet-stream;base64,${base64}`);
    expect(Array.from(result)).toEqual(Array.from(original));
  });

  it("handles a buffer larger than one chunk", () => {
    const original = new Uint8Array(200_000).map((_, i) => i % 256);
    const base64 = arrayBufferToBase64(original.buffer as ArrayBuffer);
    const result = dataUrlToUint8Array(`data:application/octet-stream;base64,${base64}`);
    expect(result.length).toBe(original.length);
    expect(Array.from(result)).toEqual(Array.from(original));
  });
});

describe("convertPdfToSlidesPptx", () => {
  it("returns null for bytes that aren't a PDF at all", async () => {
    const result = await convertPdfToSlidesPptx("data:application/pdf;base64,AAAA", deps);
    expect(result).toBeNull();
  });

  it("converts a single-page PDF into a one-slide .pptx", async () => {
    const result = await convertPdfToSlidesPptx(buildSyntheticPdfDataUrl(1), deps);
    expect(result).not.toBeNull();
    expect(result?.pageCount).toBe(1);

    const zip = await JSZip.loadAsync(dataUrlToUint8Array(result!.dataUrl));
    const slideFiles = Object.keys(zip.files).filter((f) => /^ppt\/slides\/slide\d+\.xml$/.test(f));
    expect(slideFiles).toHaveLength(1);
  });

  it("converts a multi-page PDF into one slide per page, in order", async () => {
    const result = await convertPdfToSlidesPptx(buildSyntheticPdfDataUrl(3), deps);
    expect(result).not.toBeNull();
    expect(result?.pageCount).toBe(3);

    const zip = await JSZip.loadAsync(dataUrlToUint8Array(result!.dataUrl));
    const slideFiles = Object.keys(zip.files).filter((f) => /^ppt\/slides\/slide\d+\.xml$/.test(f));
    expect(slideFiles).toHaveLength(3);

    // Each slide carries exactly one full-bleed image of its page.
    const presentationXml = await zip.file("ppt/presentation.xml")?.async("string");
    expect(presentationXml).toContain("sldSz");
  });

  it("sizes the slide to the PDF's own page aspect ratio (US Letter, portrait)", async () => {
    const result = await convertPdfToSlidesPptx(buildSyntheticPdfDataUrl(1), deps);
    const zip = await JSZip.loadAsync(dataUrlToUint8Array(result!.dataUrl));
    const presentationXml = await zip.file("ppt/presentation.xml")?.async("string");
    const match = presentationXml?.match(/<p:sldSz cx="(\d+)" cy="(\d+)"/);
    expect(match).not.toBeNull();
    const [, cx, cy] = match!;
    // US Letter is 612x792pt = 8.5x11in; portrait, so height > width.
    expect(Number(cy)).toBeGreaterThan(Number(cx));
    expect(Number(cx) / 914400).toBeCloseTo(8.5, 1);
    expect(Number(cy) / 914400).toBeCloseTo(11, 1);
  });

  it("returns null for a zero-page PDF", async () => {
    const result = await convertPdfToSlidesPptx(buildSyntheticPdfDataUrl(0), deps);
    expect(result).toBeNull();
  });
});
