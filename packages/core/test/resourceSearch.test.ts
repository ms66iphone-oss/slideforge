import { describe, expect, it } from "vitest";

import { findResourceCandidates } from "../src/resourceSearch.js";
import type { ResourceEntry, ResourceMap } from "../src/types.js";

function entry(relativePath: string): ResourceEntry {
  return { relativePath, blob: new Blob(), dataUrl: `data:${relativePath}`, bytes: 0 };
}

function resources(...paths: string[]): ResourceMap {
  return new Map(paths.map((path) => [path.toLowerCase(), entry(path)]));
}

describe("findResourceCandidates", () => {
  it("ranks filenames sharing more words with the lesson text first", () => {
    const candidates = findResourceCandidates(
      "Today we cover alternate forms of acceleration and constant acceleration.",
      resources(
        "78_Kinematics notes Graphs.pdf",
        "80_alternate forms of acceleration.pdf",
        "81_Kinematics Constant Acceleration.pdf",
        "12_Unrelated topic entirely.pdf",
      ),
    );

    expect(candidates.map((c) => c.entry.relativePath)).toEqual([
      "80_alternate forms of acceleration.pdf",
      "81_Kinematics Constant Acceleration.pdf",
    ]);
    expect(candidates[0]?.score).toBeGreaterThanOrEqual(candidates[1]?.score ?? 0);
  });

  it("ignores the Seqta numeric export prefix when matching", () => {
    const candidates = findResourceCandidates(
      "Kinematics graphs lesson",
      resources("99_Kinematics Graphs.pdf"),
    );

    expect(candidates).toHaveLength(1);
    expect(candidates[0]?.entry.relativePath).toBe("99_Kinematics Graphs.pdf");
  });

  it("returns nothing when the lesson text has no usable words", () => {
    expect(findResourceCandidates("", resources("handout.pdf"))).toEqual([]);
    expect(findResourceCandidates("the and for", resources("handout.pdf"))).toEqual([]);
  });

  it("returns nothing when no filename shares a word with the lesson text", () => {
    const candidates = findResourceCandidates(
      "Photosynthesis in plants",
      resources("Quadratic Equations Worksheet.pdf"),
    );

    expect(candidates).toEqual([]);
  });

  it("caps the result at the given limit", () => {
    const paths = Array.from({ length: 12 }, (_, i) => `${i}_Kinematics variant ${i}.pdf`);
    const candidates = findResourceCandidates("Kinematics lesson", resources(...paths), 5);

    expect(candidates).toHaveLength(5);
  });

  it("breaks ties alphabetically by path for a predictable order", () => {
    const candidates = findResourceCandidates(
      "Kinematics lesson",
      resources("b_Kinematics.pdf", "a_Kinematics.pdf"),
    );

    expect(candidates.map((c) => c.entry.relativePath)).toEqual([
      "a_Kinematics.pdf",
      "b_Kinematics.pdf",
    ]);
  });
});
