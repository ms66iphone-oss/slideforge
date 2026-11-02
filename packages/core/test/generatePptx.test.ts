import JSZip from "jszip";
import * as pdfjsLib from "pdfjs-dist/legacy/build/pdf.mjs";
import PptxGenJS from "pptxgenjs";
import { describe, expect, it } from "vitest";

import { addContentTitle, planContentPages, resolveResourceEmbed } from "../src/generatePptx.js";
import type { CanvasLike, PdfjsLibLike } from "../src/pdfToSlides.js";
import type { GenItem, Theme } from "../src/types.js";
import { buildSyntheticPdfDataUrl } from "./pdfFixture.js";

// resolveResourceEmbed's PDF path needs a real canvas and a Node-targeted
// PDF.js build, same as pdfToSlides.test.ts — jsdom has neither by default.
let nodeCanvasFactory: (width: number, height: number) => CanvasLike;
async function pdfDeps() {
  if (!nodeCanvasFactory) {
    const { createCanvas } = await import("@napi-rs/canvas");
    nodeCanvasFactory = (w, h) => createCanvas(w, h) as unknown as CanvasLike;
  }
  return { pdfjsLib: pdfjsLib as unknown as PdfjsLibLike, canvasFactory: nodeCanvasFactory };
}

function paragraph(length: number): GenItem {
  return { kind: "paragraph", paragraph: [{ kind: "plain", text: "A".repeat(length) }] };
}

function image(filename = "diagram.png"): GenItem {
  return {
    kind: "image",
    image: { dataUrl: null, filename, position: "left", size: "medium" },
  };
}

function resource(
  filename = "handout.docx",
  dataUrl: string | null = "data:application/octet-stream;base64,AAAA",
): GenItem {
  return { kind: "resource", resource: { filename, dataUrl } };
}

/** A non-paragraph filler block, so it can take up space without itself
 * being eligible to bond with an adjacent image (see isCaptionPair). */
function bulletFiller(itemCount: number): GenItem {
  return {
    kind: "bulletList",
    bulletList: Array.from({ length: itemCount }, (_, i) => `Item ${i}`),
  };
}

describe("planContentPages", () => {
  it("keeps a caption paragraph and its adjacent image on the same page even when the pair doesn't fit what's left", async () => {
    // Long enough filler leaves just enough room for the caption alone, but
    // not enough for the caption + image together — this reproduces an
    // image getting bumped to a "(cont.)" slide while its caption stayed
    // behind, orphaned, on the slide before it.
    const items: GenItem[] = [paragraph(2200), paragraph(8), image()];

    const pages = await planContentPages(items);

    expect(pages).toHaveLength(2);
    // The filler paragraph stays alone on page 1.
    expect(pages[0]).toHaveLength(1);
    // The caption and the image move together onto page 2, in order.
    expect(pages[1]?.map((b) => b.kind)).toEqual(["paragraph", "imageRow"]);
  });

  it("keeps an image and its following caption together when the image comes first", async () => {
    // Filler is a bullet list here (not a paragraph) specifically so it
    // doesn't itself bond with the image — this isolates the image+caption
    // pairing direction from the paragraph+image pairing covered above.
    const items: GenItem[] = [bulletFiller(11), image(), paragraph(8)];

    const pages = await planContentPages(items);

    expect(pages).toHaveLength(2);
    expect(pages[0]).toHaveLength(1);
    expect(pages[1]?.map((b) => b.kind)).toEqual(["imageRow", "paragraph"]);
  });

  it("does not bond a third adjacent block, so bonding can't swallow a whole slide", async () => {
    const items: GenItem[] = [paragraph(20), image(), paragraph(20)];

    const pages = await planContentPages(items);

    // Everything still fits on one page here; the point is just that the
    // grouping doesn't chain paragraph+image+paragraph into one unbreakable
    // block that would then be free to overflow the slide.
    expect(pages).toHaveLength(1);
    expect(pages[0]?.map((b) => b.kind)).toEqual(["paragraph", "imageRow", "paragraph"]);
  });

  it("still paginates normally between unrelated blocks that overflow", async () => {
    const items: GenItem[] = [paragraph(2200), paragraph(2200)];

    const pages = await planContentPages(items);

    expect(pages).toHaveLength(2);
    expect(pages[0]).toHaveLength(1);
    expect(pages[1]).toHaveLength(1);
  });

  it("returns a single empty page for a slide with no items", async () => {
    const pages = await planContentPages([]);

    expect(pages).toHaveLength(1);
    expect(pages[0]).toEqual([]);
  });

  it("gives a resource card its own page even when the current page has plenty of room", async () => {
    // Bullet-list filler (not a paragraph) on both sides, so neither one
    // bonds to the resource as its caption — this isolates "gets its own
    // page" from the caption-bonding behaviour covered separately below.
    const items: GenItem[] = [bulletFiller(3), resource(), bulletFiller(3)];

    const pages = await planContentPages(items);

    expect(pages).toHaveLength(3);
    expect(pages[0]?.map((b) => b.kind)).toEqual(["bulletList"]);
    expect(pages[1]?.map((b) => b.kind)).toEqual(["resource"]);
    expect(pages[2]?.map((b) => b.kind)).toEqual(["bulletList"]);
  });

  it("keeps a resource card's caption with it on its dedicated page when the caption comes first", async () => {
    const items: GenItem[] = [bulletFiller(11), paragraph(8), resource(), paragraph(20)];

    const pages = await planContentPages(items);

    expect(pages).toHaveLength(3);
    expect(pages[0]?.map((b) => b.kind)).toEqual(["bulletList"]);
    expect(pages[1]?.map((b) => b.kind)).toEqual(["paragraph", "resource"]);
    expect(pages[2]?.map((b) => b.kind)).toEqual(["paragraph"]);
  });

  it("keeps a resource card's caption with it on its dedicated page when the caption comes after", async () => {
    const items: GenItem[] = [bulletFiller(11), resource(), paragraph(8), paragraph(20)];

    const pages = await planContentPages(items);

    expect(pages).toHaveLength(3);
    expect(pages[0]?.map((b) => b.kind)).toEqual(["bulletList"]);
    expect(pages[1]?.map((b) => b.kind)).toEqual(["resource", "paragraph"]);
    expect(pages[2]?.map((b) => b.kind)).toEqual(["paragraph"]);
  });

  it("gives each of two back-to-back resource cards its own page", async () => {
    const items: GenItem[] = [resource("a.docx"), resource("b.docx")];

    const pages = await planContentPages(items);

    expect(pages).toHaveLength(2);
    expect(pages[0]?.map((b) => b.kind)).toEqual(["resource"]);
    expect(pages[1]?.map((b) => b.kind)).toEqual(["resource"]);
  });
});

const TEST_THEME: Theme = {
  navy: "1B2A4A",
  gold: "C9A227",
  background: "FFFFFF",
  border: "D9D9D9",
};

/** Renders a title with addContentTitle onto a fresh slide and returns the
 * resulting slide XML, so the test can check what shapes actually landed in
 * the .pptx rather than relying on pptxgenjs's internal object shape. */
async function renderTitleXml(title: string): Promise<string> {
  const pptx = new PptxGenJS();
  const slide = pptx.addSlide();
  addContentTitle(slide, TEST_THEME, title);

  const bytes = (await pptx.write({ outputType: "arraybuffer" })) as ArrayBuffer;
  const zip = await JSZip.loadAsync(bytes);
  const slideXml = await zip.file("ppt/slides/slide1.xml")?.async("string");
  if (!slideXml) throw new Error("slide1.xml was not produced");
  return slideXml;
}

describe("addContentTitle", () => {
  it("renders a plain single-line title as one text shape", async () => {
    const xml = await renderTitleXml("Understanding Scams");

    expect((xml.match(/<p:sp>/g) ?? []).length).toBe(1);
    expect(xml).toContain("Understanding Scams");
  });

  it("splits a merged multi-line title into a larger title and a smaller subtitle", async () => {
    const xml = await renderTitleXml("Statement of Inquiry:\nA quote here.");

    expect((xml.match(/<p:sp>/g) ?? []).length).toBe(2);
    expect(xml).toContain("Statement of Inquiry:");
    expect(xml).toContain("A quote here.");
    expect(xml).toContain('sz="2800"'); // the real title keeps its original 28pt size
    expect(xml).toContain('sz="1800"'); // the subtitle line is smaller
  });

  it("joins three or more merged heading lines into a single subtitle line", async () => {
    const xml = await renderTitleXml("Line One\nLine Two\nLine Three");

    expect((xml.match(/<p:sp>/g) ?? []).length).toBe(2);
    expect(xml).toContain("Line One");
    expect(xml).toContain("Line Two Line Three");
  });
});

describe("resolveResourceEmbed", () => {
  it("embeds a docx/xlsx/pptx resource directly, unchanged", async () => {
    const result = await resolveResourceEmbed({
      filename: "handout.docx",
      dataUrl: "data:application/octet-stream;base64,AAAA",
    });
    expect(result?.kind.progId).toBe("Word.Document.12");
    expect(result?.dataUrl).toBe("data:application/octet-stream;base64,AAAA");
  });

  it("converts a convertible PDF into an embeddable pptx", async () => {
    const result = await resolveResourceEmbed(
      { filename: "worksheet.pdf", dataUrl: buildSyntheticPdfDataUrl(2) },
      await pdfDeps(),
    );
    expect(result?.kind.progId).toBe("PowerPoint.Show.12");
    expect(result?.dataUrl).toContain(
      "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    );
  });

  it("falls back to null for a PDF that can't be parsed", async () => {
    const result = await resolveResourceEmbed(
      { filename: "corrupt.pdf", dataUrl: "data:application/pdf;base64,AAAA" },
      await pdfDeps(),
    );
    expect(result).toBeNull();
  });

  it("returns null for a missing resource", async () => {
    const result = await resolveResourceEmbed({ filename: "worksheet.pdf", dataUrl: null });
    expect(result).toBeNull();
  });

  it("returns null for a file type that's neither a package nor a PDF", async () => {
    const result = await resolveResourceEmbed({
      filename: "notes.txt",
      dataUrl: "data:text/plain;base64,AAAA",
    });
    expect(result).toBeNull();
  });
});
