import JSZip from "jszip";
import { describe, expect, it } from "vitest";

import { buildResourceBundle } from "../src/bundleResources.js";

const PPTX_BYTES = new TextEncoder().encode("fake pptx bytes").buffer as ArrayBuffer;

describe("buildResourceBundle", () => {
  it("puts the presentation at the top level and each resource under Resources/", async () => {
    const bundle = await buildResourceBundle(PPTX_BYTES, "Lesson 9", [
      { filename: "worksheet.pdf", dataUrl: "data:application/pdf;base64,AAAA" },
      { filename: "slides.pptx", dataUrl: "data:application/octet-stream;base64,BBBB" },
    ]);

    expect(bundle.fileName).toBe("Lesson 9.zip");

    const zip = await JSZip.loadAsync(bundle.bytes);
    expect(Object.keys(zip.files).sort()).toEqual(
      [
        "Lesson 9.pptx",
        "README.txt",
        "Resources/",
        "Resources/slides.pptx",
        "Resources/worksheet.pdf",
      ].sort(),
    );

    const pptxBytes = await zip.file("Lesson 9.pptx")?.async("string");
    expect(pptxBytes).toBe("fake pptx bytes");
  });

  it("lists every bundled resource by name in the README", async () => {
    const bundle = await buildResourceBundle(PPTX_BYTES, "Lesson 9", [
      { filename: "worksheet.pdf", dataUrl: "data:application/pdf;base64,AAAA" },
    ]);
    const zip = await JSZip.loadAsync(bundle.bytes);
    const readme = await zip.file("README.txt")?.async("string");

    expect(readme).toContain("Lesson 9.pptx");
    expect(readme).toContain("worksheet.pdf");
    expect(readme).toContain("Resources");
  });

  it("disambiguates two bundled resources that share a filename", async () => {
    const bundle = await buildResourceBundle(PPTX_BYTES, "Lesson 9", [
      { filename: "handout.pdf", dataUrl: "data:application/pdf;base64,AAAA" },
      { filename: "handout.pdf", dataUrl: "data:application/pdf;base64,BBBB" },
    ]);
    const zip = await JSZip.loadAsync(bundle.bytes);
    expect(Object.keys(zip.files).sort()).toEqual(
      [
        "Lesson 9.pptx",
        "README.txt",
        "Resources/",
        "Resources/handout (2).pdf",
        "Resources/handout.pdf",
      ].sort(),
    );
  });

  it("round-trips a resource's exact bytes through the zip", async () => {
    // "AQIDBA==" is base64 for the bytes [1, 2, 3, 4].
    const bundle = await buildResourceBundle(PPTX_BYTES, "Lesson 9", [
      { filename: "worksheet.pdf", dataUrl: "data:application/pdf;base64,AQIDBA==" },
    ]);
    const zip = await JSZip.loadAsync(bundle.bytes);
    const bytes = await zip.file("Resources/worksheet.pdf")?.async("uint8array");
    expect(Array.from(bytes ?? [])).toEqual([1, 2, 3, 4]);
  });
});
