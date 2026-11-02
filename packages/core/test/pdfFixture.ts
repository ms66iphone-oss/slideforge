// A minimal, valid, hand-built multi-page PDF for tests — deliberately not a
// real lesson resource (see AGENTS.md: no real lesson content or student
// data in tests or fixtures). Builds its own correct xref table rather than
// relying on PDF.js's broken-file recovery path, so a test failure here
// means the conversion code is wrong, not that the fixture is malformed.

function toBytes(text: string): Uint8Array {
  return new TextEncoder().encode(text);
}

export function buildSyntheticPdf(pageCount: number): Uint8Array {
  const objects = new Map<number, string>();

  const pagesObjNum = 2;
  const fontObjNum = 3;
  const firstPageObjNum = 4;
  // Each page gets a "Page N" object followed immediately by its content stream object.
  const pageObjNums = Array.from({ length: pageCount }, (_, i) => firstPageObjNum + i * 2);
  const contentObjNums = pageObjNums.map((n) => n + 1);

  objects.set(1, `<< /Type /Catalog /Pages ${pagesObjNum} 0 R >>`);
  objects.set(
    pagesObjNum,
    `<< /Type /Pages /Kids [${pageObjNums.map((n) => `${n} 0 R`).join(" ")}] /Count ${pageCount} >>`,
  );
  objects.set(fontObjNum, `<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>`);

  for (let i = 0; i < pageCount; i++) {
    const pageObjNum = pageObjNums[i]!;
    const contentObjNum = contentObjNums[i]!;
    objects.set(
      pageObjNum,
      `<< /Type /Page /Parent ${pagesObjNum} 0 R /MediaBox [0 0 612 792] ` +
        `/Resources << /Font << /F1 ${fontObjNum} 0 R >> >> /Contents ${contentObjNum} 0 R >>`,
    );
    const streamText = `BT /F1 24 Tf 72 700 Td (Synthetic test page ${i + 1}) Tj ET`;
    objects.set(
      contentObjNum,
      `<< /Length ${streamText.length} >>\nstream\n${streamText}\nendstream`,
    );
  }

  const objectCount = pageCount === 0 ? fontObjNum + 1 : (contentObjNums.at(-1) ?? fontObjNum) + 1;

  let body = "%PDF-1.4\n";
  const offsets: number[] = Array.from({ length: objectCount }, () => 0);
  for (let n = 1; n < objectCount; n++) {
    offsets[n] = toBytes(body).length;
    body += `${n} 0 obj\n${objects.get(n)}\nendobj\n`;
  }

  const xrefOffset = toBytes(body).length;
  let xref = `xref\n0 ${objectCount}\n0000000000 65535 f \n`;
  for (let n = 1; n < objectCount; n++) {
    xref += `${String(offsets[n]).padStart(10, "0")} 00000 n \n`;
  }
  body += xref;
  body += `trailer\n<< /Size ${objectCount} /Root 1 0 R >>\nstartxref\n${xrefOffset}\n%%EOF`;

  return toBytes(body);
}

export function buildSyntheticPdfDataUrl(pageCount: number): string {
  const bytes = buildSyntheticPdf(pageCount);
  let binary = "";
  for (const byte of bytes) binary += String.fromCharCode(byte);
  return `data:application/pdf;base64,${btoa(binary)}`;
}
