// Same fixed-theme, deterministic layout rules as the original Node/pptxgenjs
// generator, ported to run directly in the browser tab:
//   - images come in as data: URLs (already read into memory by zipImport.ts)
//     — pptxgenjs's addImage() accepts a `data` URI directly.
//   - image dimensions are read via a browser Image() load.
//   - pptx.writeFile() in a browser context triggers a normal browser
//     download instead of writing to disk — nothing here ever leaves the tab.

import PptxGenJS from "pptxgenjs";

import { embedOleObjects, resolveEmbeddableKind } from "./oleEmbed.js";
import type { EmbeddableKind, PendingOleEmbed } from "./oleEmbed.js";
import { convertPdfToSlidesPptx } from "./pdfToSlides.js";
import type { ConvertPdfDeps } from "./pdfToSlides.js";
import { RESOURCE_ICON_PNG_BASE64 } from "./resourceIcon.js";
import type { GenerationInput, GenItem, GenResource, GenSlide, TextRun, Theme } from "./types.js";

/** A PDF isn't an OOXML package, so it can't be embedded the way
 * docx/xlsx/pptx are (see oleEmbed.ts) — it's first converted into a small
 * generated .pptx (one slide per page) and embedded as that instead. Done
 * once here so the height this resource claims in the page-flow pass and the
 * bytes actually embedded in the render pass always agree. Falls back to the
 * plain resource card (null) if the PDF can't be parsed or rendered.
 *
 * `pdfDeps` only exists so tests can supply a Node-targeted PDF.js build and
 * canvas (jsdom has neither) — production code never passes it, and gets the
 * browser defaults from pdfToSlides.ts. */
export async function resolveResourceEmbed(
  resource: GenResource,
  pdfDeps?: ConvertPdfDeps,
): Promise<{ kind: EmbeddableKind; dataUrl: string } | null> {
  if (!resource.dataUrl) return null;

  const directKind = resolveEmbeddableKind(resource.filename);
  if (directKind) return { kind: directKind, dataUrl: resource.dataUrl };

  if (!/\.pdf$/i.test(resource.filename)) return null;
  const converted = await convertPdfToSlidesPptx(resource.dataUrl, pdfDeps);
  if (!converted) return null;
  const pptxKind = resolveEmbeddableKind("converted.pptx");
  if (!pptxKind) return null; // Unreachable — pptx is always a recognised kind.
  return { kind: pptxKind, dataUrl: converted.dataUrl };
}

// pptxgenjs's shipped types don't model every option shape used below
// (fill/line objects, the hyperlink run option, percentage-string
// coordinates); rather than fight that per call site, run options are typed
// loosely at this one boundary and left strict everywhere else in this file.
type PptxSlide = PptxGenJS.Slide;
type TextRunOptions = Record<string, unknown>;
type TextItem = { text: string; options?: TextRunOptions };

const THEME_DEFAULTS: Theme = {
  navy: "1B2A4A",
  gold: "C9A227",
  background: "FFFFFF",
  border: "D9D9D9",
};

const SLIDE_W = 13.333;
const SLIDE_H = 7.5;

// Seqta's bracket tags use British "centre"; both spellings are mapped here.
const POSITION_X: Record<string, number> = {
  left: 0.5,
  center: (SLIDE_W - 4) / 2,
  centre: (SLIDE_W - 4) / 2,
  right: SLIDE_W - 4.5,
};
const SIZE_WIDTH: Record<string, number> = { small: 3, medium: 4.5, large: 6 };

function imageBox(position: string | null, size: string | null): { x: number; w: number } {
  const w = (size && SIZE_WIDTH[size]) || SIZE_WIDTH.medium;
  const x = (position && POSITION_X[position]) ?? POSITION_X.right;
  return { x: x as number, w: w as number };
}

/** Reads natural pixel dimensions of a data: URL image via a throwaway <img>. */
function loadImageDimensions(dataUrl: string): Promise<{ width: number; height: number } | null> {
  return new Promise((resolve) => {
    const img = new Image();
    img.onload = () => resolve({ width: img.naturalWidth, height: img.naturalHeight });
    img.onerror = () => resolve(null);
    img.src = dataUrl;
  });
}

interface ImageBox {
  x: number;
  w: number;
  h: number;
}

/** Estimates an image's rendered box (x, w, h), using the real file's aspect
 * ratio when available and falling back to a fixed ratio for a missing
 * resource (still needs a placeholder-sized box in the flow). */
async function estimateImageBox(image: {
  position: string | null;
  size: string | null;
  dataUrl: string | null;
}): Promise<ImageBox> {
  const box = imageBox(image.position, image.size);
  let h = box.w * 0.66;
  if (image.dataUrl) {
    const dims = await loadImageDimensions(image.dataUrl);
    if (dims && dims.width && dims.height) h = box.w * (dims.height / dims.width);
  }
  return { x: box.x, w: box.w, h };
}

// --- Overflow-safe text flow (identical formula to the original generator) ---
const CHAR_WIDTH_IN = 0.105;
const LINE_HEIGHT_IN = 0.24;
const CONTENT_TOP = 1.5;
const CONTENT_BOTTOM = SLIDE_H - 0.5;

// --- Embedded-resource card layout ---
const RESOURCE_ICON_W = 0.7;
const RESOURCE_ICON_H = 0.875; // matches the icon asset's 240:300 aspect ratio
const RESOURCE_ROW_HEIGHT = 1.1;
const RESOURCE_CARD_HEIGHT = 0.8; // the plain text-card fallback, unchanged from before

function estimateParagraphHeight(text: string, textW: number): number {
  const charsPerLine = Math.max(10, Math.floor(textW / CHAR_WIDTH_IN));
  const lines = Math.max(1, Math.ceil(text.length / charsPerLine));
  return Math.max(0.4, lines * LINE_HEIGHT_IN + 0.1);
}

function estimateBulletListHeight(items: string[]): number {
  return 0.4 * items.length + 0.2;
}

function estimateTableHeight(rows: string[][]): number {
  return 0.5 * rows.length + 0.3;
}

type ImageItem = Extract<GenItem, { kind: "image" }>;

type FlowBlock =
  | { kind: "paragraph"; data: TextRun[]; height: number }
  | { kind: "bulletList"; data: string[]; height: number }
  | { kind: "table"; data: string[][]; height: number }
  | {
      kind: "resource";
      data: GenResource;
      embed: { kind: EmbeddableKind; dataUrl: string } | null;
      height: number;
    }
  | { kind: "imageRow"; data: ImageItem[]; boxes: ImageBox[]; height: number };

type GroupedItem = GenItem | { kind: "imageRow"; images: ImageItem[] };

const TEXT_W = SLIDE_W - 1;

/** Builds the ordered flow queue for a slide's items — this is what keeps
 * content in the same sequence it appeared in the lesson (paragraph -> its
 * list -> its table -> an inline image -> next paragraph, etc). Consecutive
 * image items (nothing else between them) collapse into one row block first,
 * so they render side by side instead of each claiming a full-height row of
 * its own — matching how they actually sat together. */
async function buildFlowBlocks(items: GenItem[]): Promise<FlowBlock[]> {
  const groupedItems: GroupedItem[] = [];
  for (const item of items) {
    const prev = groupedItems[groupedItems.length - 1];
    if (item.kind === "image" && prev && "images" in prev) {
      prev.images.push(item);
    } else if (item.kind === "image") {
      groupedItems.push({ kind: "imageRow", images: [item] });
    } else {
      groupedItems.push(item);
    }
  }

  const flowBlocks: FlowBlock[] = [];
  for (const item of groupedItems) {
    if (item.kind === "paragraph") {
      const text = item.paragraph.map((run) => run.text).join("");
      flowBlocks.push({
        kind: "paragraph",
        data: item.paragraph,
        height: estimateParagraphHeight(text, TEXT_W),
      });
    } else if (item.kind === "bulletList") {
      flowBlocks.push({
        kind: "bulletList",
        data: item.bulletList,
        height: estimateBulletListHeight(item.bulletList),
      });
    } else if (item.kind === "table") {
      flowBlocks.push({ kind: "table", data: item.table, height: estimateTableHeight(item.table) });
    } else if (item.kind === "resource") {
      const embed = await resolveResourceEmbed(item.resource);
      const height = embed ? RESOURCE_ROW_HEIGHT : RESOURCE_CARD_HEIGHT;
      flowBlocks.push({ kind: "resource", data: item.resource, embed, height });
    } else if ("images" in item) {
      const boxes = await Promise.all(item.images.map((i) => estimateImageBox(i.image)));
      const height = Math.max(...boxes.map((b) => b.h)) + 0.15;
      flowBlocks.push({ kind: "imageRow", data: item.images, boxes, height });
    }
  }
  return flowBlocks;
}

/** A paragraph sitting directly next to an image or a resource card is
 * almost always its caption (or the descriptive text it was inline with) —
 * Seqta's `[[image:...]]`/`[[resource:...]]` tags routinely split one
 * authored paragraph into "text before" / tag / "text after" blocks.
 * Pairing them here is what stops that text from being separated from what
 * it's describing. */
function isCaptionPair(a: FlowBlock, b: FlowBlock): boolean {
  const isMedia = (block: FlowBlock) => block.kind === "imageRow" || block.kind === "resource";
  return (a.kind === "paragraph" && isMedia(b)) || (isMedia(a) && b.kind === "paragraph");
}

/** Groups adjacent flow blocks that must never be split across a slide
 * break. Each group is at most a single bonded pair (a caption paragraph
 * plus its image or resource card) — a third adjacent block always starts
 * its own group, so bonding can't chain into one unbreakable run of the
 * whole slide. */
function groupFlowBlocks(blocks: FlowBlock[]): FlowBlock[][] {
  const groups: FlowBlock[][] = [];
  for (const block of blocks) {
    const prevGroup = groups[groups.length - 1];
    const prevBlock = prevGroup?.[prevGroup.length - 1];
    if (prevGroup && prevGroup.length === 1 && prevBlock && isCaptionPair(prevBlock, block)) {
      prevGroup.push(block);
    } else {
      groups.push([block]);
    }
  }
  return groups;
}

function groupHeight(group: FlowBlock[]): number {
  return group.reduce((sum, block) => sum + block.height, 0);
}

/** True for a group that carries a resource card (with or without its bonded
 * caption) — these always get a dedicated page of their own, the same way a
 * video already gets pulled onto its own slide, rather than sharing space
 * with unrelated content. */
function groupIsResource(group: FlowBlock[]): boolean {
  return group.some((block) => block.kind === "resource");
}

/** Plans how a content slide's items paginate across one or more actual
 * slides, honoring the same overflow threshold as before, but breaking
 * between bonded groups (see `groupFlowBlocks`) rather than between any two
 * blocks — so a caption and its image or resource card always land on the
 * same page. A resource card's group always gets a page to itself (see
 * `groupIsResource`), regardless of how much room was left on the page
 * before it. Always returns at least one (possibly empty) page, matching
 * the previous behaviour of a content slide with no items still getting a
 * title slide. */
export async function planContentPages(items: GenItem[]): Promise<FlowBlock[][]> {
  const flowBlocks = await buildFlowBlocks(items);
  const groups = groupFlowBlocks(flowBlocks);

  const pages: FlowBlock[][] = [];
  let current: FlowBlock[] = [];
  let currentHeight = CONTENT_TOP;

  for (const group of groups) {
    if (groupIsResource(group)) {
      if (current.length > 0) pages.push(current);
      pages.push(group);
      current = [];
      currentHeight = CONTENT_TOP;
      continue;
    }

    const h = groupHeight(group);
    if (current.length > 0 && currentHeight + h > CONTENT_BOTTOM) {
      pages.push(current);
      current = [];
      currentHeight = CONTENT_TOP;
    }
    current.push(...group);
    currentHeight += h;
  }
  if (current.length > 0 || pages.length === 0) pages.push(current);
  return pages;
}

/** Renders a content slide's title. A plain single-line title is one text
 * box, unchanged from before. A title that slideAssembler merged from
 * multiple consecutive headings (joined with "\n") renders as two: the
 * first heading as the real, larger title, and the rest as a smaller
 * subtitle line underneath — splitting one merged block into two separate
 * shapes is, empirically, what lets PowerPoint's Designer suggest a layout
 * for the slide; a single two-line title box doesn't trigger it. */
export function addContentTitle(s: PptxSlide, theme: Theme, title: string): void {
  const lines = title.split("\n");
  if (lines.length <= 1) {
    s.addText(title, {
      x: 0.5,
      y: 0.4,
      w: SLIDE_W - 1,
      h: 0.9,
      fontSize: 28,
      bold: true,
      color: theme.navy,
      fontFace: "Arial",
    });
    return;
  }

  const [primary, ...rest] = lines;
  s.addText(primary ?? title, {
    x: 0.5,
    y: 0.35,
    w: SLIDE_W - 1,
    h: 0.55,
    fontSize: 28,
    bold: true,
    color: theme.navy,
    fontFace: "Arial",
  });
  s.addText(rest.join(" "), {
    x: 0.5,
    y: 0.95,
    w: SLIDE_W - 1,
    h: 0.45,
    fontSize: 18,
    color: theme.navy,
    fontFace: "Arial",
  });
}

async function addContentSlide(
  pptx: PptxGenJS,
  theme: Theme,
  slide: Extract<GenSlide, { type: "content" }>,
  startSlideNumber: number,
  pendingEmbeds: PendingOleEmbed[],
): Promise<number> {
  const textW = TEXT_W;
  const pages = await planContentPages(slide.items);
  let slideNumber = startSlideNumber;

  for (let pageIndex = 0; pageIndex < pages.length; pageIndex++) {
    slideNumber += 1;
    const s: PptxSlide = pptx.addSlide();
    s.background = { color: theme.background };

    const title = pageIndex === 0 ? slide.title : `${slide.title} (cont.)`;
    addContentTitle(s, theme, title);

    let textY = CONTENT_TOP;

    // Consecutive paragraph blocks that land on the same slide are buffered
    // and flushed as ONE shape (one paragraph per line via breakLine) rather
    // than each getting its own addText() call.
    let paragraphBuffer: { runs: TextItem[]; height: number }[] = [];
    let paragraphBufferY: number | null = null;

    function flushParagraphBuffer() {
      if (paragraphBuffer.length === 0) return;
      const runs: TextItem[] = [];
      paragraphBuffer.forEach((p, idx) => {
        // p.runs entries are already {text, options} — options (hyperlink/
        // color/underline) were computed once when the paragraph was first
        // queued below; copy that forward rather than re-deriving it here.
        p.runs.forEach((run, runIdx) => {
          const options: TextRunOptions = { ...run.options };
          if (runIdx === p.runs.length - 1 && idx < paragraphBuffer.length - 1) {
            options.breakLine = true; // start a new line for the next paragraph
          }
          runs.push({ text: run.text, options });
        });
      });
      const totalHeight = paragraphBuffer.reduce((sum, p) => sum + p.height, 0);
      s.addText(runs as never, {
        x: 0.5,
        y: paragraphBufferY ?? CONTENT_TOP,
        w: textW,
        h: totalHeight,
        fontSize: 14,
        color: "333333",
        fontFace: "Arial",
        valign: "top",
      });
      paragraphBuffer = [];
      paragraphBufferY = null;
    }

    for (const block of pages[pageIndex] ?? []) {
      const currentSlide: PptxSlide = s;

      if (block.kind === "paragraph") {
        const runs: TextItem[] = block.data.map((run) => {
          const options: TextRunOptions = {};
          if (run.kind === "link") {
            options.hyperlink = { url: run.url };
            options.color = "0563C1";
            options.underline = true;
          }
          return { text: run.text, options };
        });
        if (paragraphBuffer.length === 0) paragraphBufferY = textY;
        paragraphBuffer.push({ runs, height: block.height });
      } else if (block.kind === "bulletList") {
        flushParagraphBuffer();
        const bulletItems: TextItem[] = block.data.map((item) => ({
          text: item,
          options: { bullet: true },
        }));
        currentSlide.addText(bulletItems as never, {
          x: 0.5,
          y: textY,
          w: textW,
          h: block.height,
          fontSize: 14,
          color: "333333",
          fontFace: "Arial",
        });
      } else if (block.kind === "table") {
        flushParagraphBuffer();
        const rows = block.data.map((row) =>
          row.map((cell) => ({ text: cell, options: { fontSize: 12 } })),
        );
        currentSlide.addTable(
          rows as never,
          {
            x: 0.5,
            y: textY,
            w: textW,
            border: { type: "solid", color: theme.border, pt: 1 },
            fill: { color: "FFFFFF" },
          } as never,
        );
      } else if (block.kind === "resource") {
        flushParagraphBuffer();
        const resource = block.data;
        const embed = block.embed;

        if (embed) {
          // The visible icon + caption are rendered normally here; the
          // actual OLE embedding (what makes double-clicking it open the
          // real file) is patched into the .pptx afterwards in generatePptx().
          currentSlide.addImage({
            data: `data:image/png;base64,${RESOURCE_ICON_PNG_BASE64}`,
            x: 0.5,
            y: textY,
            w: RESOURCE_ICON_W,
            h: RESOURCE_ICON_H,
          });
          currentSlide.addText(
            [
              { text: resource.filename, options: { bold: true, breakLine: true } },
              { text: "Double-click to open", options: { italic: true, color: "666666" } },
            ] as never,
            {
              x: 1.35,
              y: textY,
              w: textW - 0.85,
              h: RESOURCE_ICON_H,
              fontSize: 12,
              color: theme.navy,
              fontFace: "Arial",
              valign: "middle",
            } as never,
          );
          pendingEmbeds.push({
            slideNumber,
            filename: resource.filename,
            dataUrl: embed.dataUrl,
            kind: embed.kind,
            x: 0.5,
            y: textY,
            w: RESOURCE_ICON_W,
            h: RESOURCE_ICON_H,
          });
        } else if (resource.dataUrl) {
          currentSlide.addText(`Attached resource\n${resource.filename}`, {
            x: 0.5,
            y: textY,
            w: 4,
            h: RESOURCE_CARD_HEIGHT,
            fontSize: 12,
            color: theme.navy,
            fontFace: "Arial",
            fill: { color: "FFFFFF" },
            line: { color: theme.border, width: 1 },
            align: "center",
            valign: "middle",
          } as never);
        } else {
          currentSlide.addText(`Missing resource: ${resource.filename} — attach manually`, {
            x: 0.5,
            y: textY,
            w: 4,
            h: RESOURCE_CARD_HEIGHT,
            fontSize: 12,
            color: "B00020",
            italic: true,
            fontFace: "Arial",
            fill: { color: "FFFFFF" },
            line: { color: "B00020", width: 1 },
            align: "center",
            valign: "middle",
          } as never);
        }
      } else if (block.kind === "imageRow") {
        flushParagraphBuffer();
        // A single image keeps its own left/right/centre position from the
        // bracket tag. Two or more consecutive images ignore individual
        // positions and just lay out left to right from the margin.
        const gap = 0.25;
        const firstBox = block.boxes[0];
        let x = block.data.length === 1 && firstBox ? firstBox.x : 0.5;
        for (let i = 0; i < block.data.length; i++) {
          const item = block.data[i];
          const box = block.boxes[i];
          if (!item || item.kind !== "image" || !box) continue;
          const image = item.image;
          if (!image.dataUrl) {
            currentSlide.addText(
              `Missing resource: ${image.filename || "unknown"} — insert manually`,
              {
                x,
                y: textY,
                w: box.w,
                h: box.h,
                fontSize: 11,
                color: "B00020",
                italic: true,
                align: "center",
                valign: "middle",
                line: { color: "B00020", width: 1 },
              } as never,
            );
          } else {
            currentSlide.addImage({ data: image.dataUrl, x, y: textY, w: box.w, h: box.h });
          }
          x += box.w + gap;
        }
      }

      textY += block.height;
    }

    flushParagraphBuffer(); // render whatever paragraphs were still buffered at the end of this page
  }

  return slideNumber;
}

function addVideoSlide(
  pptx: PptxGenJS,
  theme: Theme,
  slide: Extract<GenSlide, { type: "video" }>,
): void {
  const s = pptx.addSlide();
  s.background = { color: theme.navy };

  s.addText(slide.title, {
    x: 0.5,
    y: 0.6,
    w: SLIDE_W - 1,
    h: 0.9,
    fontSize: 26,
    bold: true,
    color: "FFFFFF",
    fontFace: "Arial",
  });

  s.addText("▶  Watch Video", {
    x: "31%",
    y: "40%",
    w: "38%",
    h: "16%",
    fontSize: 20,
    bold: true,
    color: theme.navy,
    fill: { color: theme.gold },
    align: "center",
    valign: "middle",
    hyperlink: { url: slide.url },
  } as never);

  s.addText(slide.url, {
    x: "5%",
    y: "60%",
    w: "90%",
    h: "7%",
    fontSize: 11,
    color: "CCCCCC",
    align: "center",
  });
}

/** Triggers a normal browser download for pptx bytes we've already got in
 * memory — the same mechanism pptx.writeFile() uses internally, needed here
 * because writing the OLE-patched bytes bypasses that helper. */
function downloadPptx(bytes: ArrayBuffer, fileName: string): void {
  const blob = new Blob([bytes], {
    type: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
  });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = fileName;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
}

/** Resolves once the browser download has been triggered. */
export async function generatePptx(
  data: GenerationInput,
  theme: Partial<Theme> = {},
): Promise<void> {
  const resolvedTheme: Theme = { ...THEME_DEFAULTS, ...theme };

  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: "SLIDEFORGE_16x9", width: SLIDE_W, height: SLIDE_H });
  pptx.layout = "SLIDEFORGE_16x9";

  const pendingEmbeds: PendingOleEmbed[] = [];
  let slideNumber = 0;

  for (const slide of data.slides) {
    if (slide.type === "video") {
      slideNumber += 1;
      addVideoSlide(pptx, resolvedTheme, slide);
    } else {
      slideNumber = await addContentSlide(pptx, resolvedTheme, slide, slideNumber, pendingEmbeds);
    }
  }

  const fileName = `${data.outputName || "Presentation"}.pptx`;

  if (pendingEmbeds.length === 0) {
    await pptx.writeFile({ fileName });
    return;
  }

  const rawBytes = (await pptx.write({ outputType: "arraybuffer" })) as ArrayBuffer;
  const patchedBytes = await embedOleObjects(rawBytes, pendingEmbeds);
  downloadPptx(patchedBytes, fileName);
}
