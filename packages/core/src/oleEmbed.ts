// Turns a resource card into a real embedded OLE object in the generated
// .pptx — the same mechanism PowerPoint's own "Insert > Object > Create from
// File" uses, so the attached file travels inside the deck itself instead of
// depending on a link back to wherever it was originally hosted.
//
// pptxgenjs has no support for this, so the .pptx it writes (itself a zip)
// is patched directly with JSZip afterwards. A native OOXML file is already
// a zip package in its own right, so it can be embedded as-is via the
// "package" relationship type — no legacy OLE1/compound-file wrapping
// needed.
//
// Only docx/xlsx (and their macro-enabled variants) are listed here. A pptx
// embeds just as validly by the same mechanism and PowerPoint's UI accepts
// it (selection, context menu, "Open" submenu all work), but double-clicking
// to activate it does nothing on Mac PowerPoint — confirmed both for a
// generated embed and for an already-existing, unrelated genuine .pptx
// attachment, so this isn't specific to how the file was produced. Rather
// than ship an embed that looks right but silently doesn't open, pptx/pptm
// (and anything else outside this list, including PDFs, which were never
// embeddable at all) are bundled alongside the forged presentation instead
// — see bundleResources.ts — and the caller falls back to a plain label.

import JSZip from "jszip";

import { RESOURCE_ICON_PNG_BASE64 } from "./resourceIcon.js";

export interface EmbeddableKind {
  extension: string;
  contentType: string;
  progId: string;
}

const EMBEDDABLE_KINDS: Record<string, EmbeddableKind> = {
  docx: {
    extension: "docx",
    contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    progId: "Word.Document.12",
  },
  docm: {
    extension: "docm",
    contentType: "application/vnd.ms-word.document.macroEnabled.12",
    progId: "Word.Document.12",
  },
  xlsx: {
    extension: "xlsx",
    contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    progId: "Excel.Sheet.12",
  },
  xlsm: {
    extension: "xlsm",
    contentType: "application/vnd.ms-excel.sheet.macroEnabled.12",
    progId: "Excel.Sheet.12",
  },
};

/** True for a file type known to actually open on double-click once
 * embedded (see the module comment for why pptx isn't here despite being
 * structurally embeddable the same way). Anything else (pdf, pptx/pptm,
 * images-as-resources, plain docs, …) is out of scope for embedding — the
 * resource card bundles those alongside the presentation instead. */
export function resolveEmbeddableKind(filename: string): EmbeddableKind | null {
  const dot = filename.lastIndexOf(".");
  if (dot === -1) return null;
  const ext = filename.slice(dot + 1).toLowerCase();
  return EMBEDDABLE_KINDS[ext] ?? null;
}

export interface PendingOleEmbed {
  /** The 1-based slide number pptxgenjs will write this to (ppt/slides/slideN.xml) — matches addSlide() call order. */
  slideNumber: number;
  filename: string;
  dataUrl: string;
  kind: EmbeddableKind;
  /** Position/size in inches, matching where the resource sits in the slide's content flow. */
  x: number;
  y: number;
  w: number;
  h: number;
}

const EMU_PER_INCH = 914400;
const OLE_GRAPHIC_URI = "http://schemas.openxmlformats.org/presentationml/2006/ole";
const PACKAGE_REL_TYPE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package";
const IMAGE_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
const SHARED_ICON_PATH = "ppt/media/sfResourceIcon.png";
const SHARED_ICON_TARGET = "../media/sfResourceIcon.png";

export function dataUrlToBase64(dataUrl: string): string {
  const comma = dataUrl.indexOf(",");
  return comma === -1 ? dataUrl : dataUrl.slice(comma + 1);
}

/** The next unused "rIdN" in a .rels file's contents, so a new relationship
 * never collides with one pptxgenjs already wrote. */
export function nextRelationshipId(relsXml: string): number {
  const ids = Array.from(relsXml.matchAll(/Id="rId(\d+)"/g)).map((m) => Number(m[1]));
  return (ids.length > 0 ? Math.max(...ids) : 0) + 1;
}

export function addRelationship(relsXml: string, id: string, type: string, target: string): string {
  const entry = `<Relationship Id="${id}" Type="${type}" Target="${target}"/>`;
  return relsXml.replace("</Relationships>", `${entry}</Relationships>`);
}

export function ensureDefaultContentType(
  typesXml: string,
  extension: string,
  contentType: string,
): string {
  if (typesXml.includes(`Extension="${extension}"`)) return typesXml;
  const entry = `<Default Extension="${extension}" ContentType="${contentType}"/>`;
  return typesXml.replace("</Types>", `${entry}</Types>`);
}

function escapeXmlAttr(value: string): string {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

/** Builds the `<p:graphicFrame>` for one embedded resource: an OLE object
 * (CT_OleObject) whose required fallback picture is the shared document
 * icon. Reused verbatim from what PowerPoint's own "Insert Object" produces,
 * minus the legacy VML/AlternateContent wrapper older versions needed. */
export function buildOleGraphicFrameXml(params: {
  shapeId: number;
  displayName: string;
  packageRelId: string;
  iconRelId: string;
  progId: string;
  x: number;
  y: number;
  w: number;
  h: number;
}): string {
  const offX = Math.round(params.x * EMU_PER_INCH);
  const offY = Math.round(params.y * EMU_PER_INCH);
  const extCx = Math.round(params.w * EMU_PER_INCH);
  const extCy = Math.round(params.h * EMU_PER_INCH);
  const picId = params.shapeId + 1;
  const name = escapeXmlAttr(params.displayName);

  return (
    `<p:graphicFrame>` +
    `<p:nvGraphicFramePr>` +
    `<p:cNvPr id="${params.shapeId}" name="${name}"/>` +
    `<p:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></p:cNvGraphicFramePr>` +
    `<p:nvPr/>` +
    `</p:nvGraphicFramePr>` +
    `<p:xfrm><a:off x="${offX}" y="${offY}"/><a:ext cx="${extCx}" cy="${extCy}"/></p:xfrm>` +
    `<a:graphic><a:graphicData uri="${OLE_GRAPHIC_URI}">` +
    `<p:oleObj name="${name}" r:id="${params.packageRelId}" imgW="${extCx}" imgH="${extCy}" progId="${params.progId}" showAsIcon="1">` +
    `<p:embed/>` +
    `<p:pic>` +
    `<p:nvPicPr><p:cNvPr id="${picId}" name="${name} Icon"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>` +
    `<p:blipFill><a:blip r:embed="${params.iconRelId}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>` +
    `<p:spPr>` +
    `<a:xfrm><a:off x="${offX}" y="${offY}"/><a:ext cx="${extCx}" cy="${extCy}"/></a:xfrm>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>` +
    `</p:spPr>` +
    `</p:pic>` +
    `</p:oleObj>` +
    `</a:graphicData></a:graphic>` +
    `</p:graphicFrame>`
  );
}

/** Patches a generated .pptx (as returned by pptxgenjs's
 * `write({ outputType: "arraybuffer" })`) so each pending resource becomes a
 * real embedded OLE object placed at its reserved spot on its slide. */
export async function embedOleObjects(
  pptxBytes: ArrayBuffer,
  embeds: PendingOleEmbed[],
): Promise<ArrayBuffer> {
  if (embeds.length === 0) return pptxBytes;

  const zip = await JSZip.loadAsync(pptxBytes);

  const typesFile = zip.file("[Content_Types].xml");
  if (!typesFile) return pptxBytes; // Not a shape we recognize — leave untouched.
  let typesXml = await typesFile.async("string");

  if (!zip.file(SHARED_ICON_PATH)) {
    zip.file(SHARED_ICON_PATH, RESOURCE_ICON_PNG_BASE64, { base64: true });
    typesXml = ensureDefaultContentType(typesXml, "png", "image/png");
  }

  const extensionsAdded = new Set<string>();
  let shapeIdCounter = 900001;
  let embeddingCounter = 1;

  const bySlide = new Map<number, PendingOleEmbed[]>();
  for (const embed of embeds) {
    const list = bySlide.get(embed.slideNumber) ?? [];
    list.push(embed);
    bySlide.set(embed.slideNumber, list);
  }

  for (const [slideNumber, slideEmbeds] of bySlide) {
    const slidePath = `ppt/slides/slide${slideNumber}.xml`;
    const relsPath = `ppt/slides/_rels/slide${slideNumber}.xml.rels`;
    const slideFile = zip.file(slidePath);
    const relsFile = zip.file(relsPath);
    if (!slideFile || !relsFile) continue;

    let slideXml = await slideFile.async("string");
    let relsXml = await relsFile.async("string");
    let nextRel = nextRelationshipId(relsXml);
    let insertedXml = "";

    for (const embed of slideEmbeds) {
      if (!extensionsAdded.has(embed.kind.extension)) {
        typesXml = ensureDefaultContentType(typesXml, embed.kind.extension, embed.kind.contentType);
        extensionsAdded.add(embed.kind.extension);
      }

      const embeddingPath = `ppt/embeddings/oleObject${embeddingCounter++}.${embed.kind.extension}`;
      zip.file(embeddingPath, dataUrlToBase64(embed.dataUrl), { base64: true });

      const packageRelId = `rId${nextRel++}`;
      const iconRelId = `rId${nextRel++}`;
      relsXml = addRelationship(
        relsXml,
        packageRelId,
        PACKAGE_REL_TYPE,
        `../embeddings/${embeddingPath.split("/").pop()}`,
      );
      relsXml = addRelationship(relsXml, iconRelId, IMAGE_REL_TYPE, SHARED_ICON_TARGET);

      const shapeId = shapeIdCounter;
      shapeIdCounter += 2;

      insertedXml += buildOleGraphicFrameXml({
        shapeId,
        displayName: `Embedded resource: ${embed.filename}`,
        packageRelId,
        iconRelId,
        progId: embed.kind.progId,
        x: embed.x,
        y: embed.y,
        w: embed.w,
        h: embed.h,
      });
    }

    if (insertedXml) {
      slideXml = slideXml.replace("</p:spTree>", `${insertedXml}</p:spTree>`);
      zip.file(slidePath, slideXml);
      zip.file(relsPath, relsXml);
    }
  }

  zip.file("[Content_Types].xml", typesXml);

  return zip.generateAsync({ type: "arraybuffer" });
}
