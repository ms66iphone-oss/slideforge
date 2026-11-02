import JSZip from "jszip";
import { describe, expect, it } from "vitest";

import {
  addRelationship,
  buildOleGraphicFrameXml,
  dataUrlToBase64,
  embedOleObjects,
  ensureDefaultContentType,
  nextRelationshipId,
  resolveEmbeddableKind,
} from "../src/oleEmbed.js";
import type { PendingOleEmbed } from "../src/oleEmbed.js";

describe("resolveEmbeddableKind", () => {
  it("recognizes docx/xlsx and their macro-enabled variants, case-insensitively", () => {
    expect(resolveEmbeddableKind("Handout.DOCX")?.progId).toBe("Word.Document.12");
    expect(resolveEmbeddableKind("macro.docm")?.progId).toBe("Word.Document.12");
    expect(resolveEmbeddableKind("budget.xlsx")?.progId).toBe("Excel.Sheet.12");
    expect(resolveEmbeddableKind("macro.xlsm")?.progId).toBe("Excel.Sheet.12");
  });

  it("returns null for pptx/pptm — structurally embeddable, but Mac PowerPoint can't activate it", () => {
    expect(resolveEmbeddableKind("slides.pptx")).toBeNull();
    expect(resolveEmbeddableKind("slides.pptm")).toBeNull();
  });

  it("returns null for anything outside the embeddable family", () => {
    expect(resolveEmbeddableKind("worksheet.pdf")).toBeNull();
    expect(resolveEmbeddableKind("photo.png")).toBeNull();
    expect(resolveEmbeddableKind("no-extension")).toBeNull();
  });
});

describe("dataUrlToBase64", () => {
  it("strips the data: URL prefix", () => {
    expect(dataUrlToBase64("data:application/octet-stream;base64,AAAA")).toBe("AAAA");
  });

  it("returns the input unchanged if there's no comma", () => {
    expect(dataUrlToBase64("AAAA")).toBe("AAAA");
  });
});

describe("nextRelationshipId", () => {
  it("returns 1 for a rels file with no relationships", () => {
    expect(nextRelationshipId("<Relationships></Relationships>")).toBe(1);
  });

  it("returns one past the highest existing rId", () => {
    const rels =
      '<Relationships><Relationship Id="rId1"/><Relationship Id="rId3"/></Relationships>';
    expect(nextRelationshipId(rels)).toBe(4);
  });
});

describe("addRelationship", () => {
  it("inserts a new relationship before the closing tag", () => {
    const result = addRelationship(
      '<Relationships><Relationship Id="rId1"/></Relationships>',
      "rId2",
      "http://example.com/type",
      "../media/icon.png",
    );
    expect(result).toBe(
      '<Relationships><Relationship Id="rId1"/><Relationship Id="rId2" Type="http://example.com/type" Target="../media/icon.png"/></Relationships>',
    );
  });
});

describe("ensureDefaultContentType", () => {
  it("adds a Default entry when the extension isn't registered yet", () => {
    const result = ensureDefaultContentType("<Types></Types>", "docx", "application/word");
    expect(result).toBe(
      '<Types><Default Extension="docx" ContentType="application/word"/></Types>',
    );
  });

  it("leaves the document untouched when the extension is already registered", () => {
    const original = '<Types><Default Extension="docx" ContentType="application/word"/></Types>';
    expect(ensureDefaultContentType(original, "docx", "application/word")).toBe(original);
  });
});

describe("buildOleGraphicFrameXml", () => {
  it("embeds (not links) the object and converts inches to EMU", () => {
    const xml = buildOleGraphicFrameXml({
      shapeId: 900001,
      displayName: "Embedded resource: handout.docx",
      packageRelId: "rId5",
      iconRelId: "rId6",
      progId: "Word.Document.12",
      x: 0.5,
      y: 1.5,
      w: 0.7,
      h: 0.875,
    });

    expect(xml).toContain("<p:embed/>");
    expect(xml).not.toContain("<p:link");
    expect(xml).toContain('r:id="rId5"');
    expect(xml).toContain('r:embed="rId6"');
    expect(xml).toContain('progId="Word.Document.12"');
    expect(xml).toContain('<a:off x="457200" y="1371600"/>'); // 0.5in, 1.5in in EMU
    expect(xml).toContain('<a:ext cx="640080" cy="800100"/>'); // 0.7in, 0.875in in EMU
  });
});

function minimalPptxZip(): JSZip {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"></Types>',
  );
  zip.file(
    "ppt/slides/slide1.xml",
    '<?xml version="1.0"?><p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:cSld><p:spTree></p:spTree></p:cSld></p:sld>',
  );
  zip.file(
    "ppt/slides/_rels/slide1.xml.rels",
    '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/></Relationships>',
  );
  return zip;
}

describe("embedOleObjects", () => {
  it("returns the bytes unchanged when there's nothing to embed", async () => {
    const zip = minimalPptxZip();
    const bytes = await zip.generateAsync({ type: "arraybuffer" });

    const result = await embedOleObjects(bytes, []);

    expect(result).toBe(bytes);
  });

  it("patches the slide, its rels, and the content types with a real embedded object", async () => {
    const zip = minimalPptxZip();
    const bytes = await zip.generateAsync({ type: "arraybuffer" });

    const embed: PendingOleEmbed = {
      slideNumber: 1,
      filename: "handout.docx",
      dataUrl:
        "data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64,UEsD",
      kind: {
        extension: "docx",
        contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        progId: "Word.Document.12",
      },
      x: 0.5,
      y: 1.5,
      w: 0.7,
      h: 0.875,
    };

    const result = await embedOleObjects(bytes, [embed]);
    const patched = await JSZip.loadAsync(result);

    const embeddedFile = patched.file("ppt/embeddings/oleObject1.docx");
    expect(embeddedFile).not.toBeNull();
    expect(await embeddedFile?.async("string")).toBe(atob("UEsD"));

    expect(patched.file("ppt/media/sfResourceIcon.png")).not.toBeNull();

    const relsXml = await patched.file("ppt/slides/_rels/slide1.xml.rels")?.async("string");
    expect(relsXml).toContain(
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package",
    );
    expect(relsXml).toContain("../embeddings/oleObject1.docx");
    expect(relsXml).toContain("../media/sfResourceIcon.png");

    const slideXml = await patched.file("ppt/slides/slide1.xml")?.async("string");
    expect(slideXml).toContain("<p:graphicFrame>");
    expect(slideXml).toContain('progId="Word.Document.12"');

    const typesXml = await patched.file("[Content_Types].xml")?.async("string");
    expect(typesXml).toContain(
      'Extension="docx" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document"',
    );
    expect(typesXml).toContain('Extension="png"');
  });

  it("leaves resources untouched when their slide/rels can't be found", async () => {
    const zip = minimalPptxZip();
    const bytes = await zip.generateAsync({ type: "arraybuffer" });

    const embed: PendingOleEmbed = {
      slideNumber: 99, // no such slide in this fixture
      filename: "handout.docx",
      dataUrl: "data:application/octet-stream;base64,AAAA",
      kind: {
        extension: "docx",
        contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        progId: "Word.Document.12",
      },
      x: 0.5,
      y: 1.5,
      w: 0.7,
      h: 0.875,
    };

    const result = await embedOleObjects(bytes, [embed]);
    const patched = await JSZip.loadAsync(result);

    expect(patched.file("ppt/embeddings/oleObject1.docx")).toBeNull();
  });
});
