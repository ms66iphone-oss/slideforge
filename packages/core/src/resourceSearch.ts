// Narrows a teacher's full resources list down to a short, likely-matching
// set for lessons whose Seqta HTML has no resource tags at all — either the
// HTML isn't available, or its own "Resources" block didn't paste through
// with usable filenames. There's nothing to pattern-match against in that
// case, so instead we compare the plain words already on the lesson's own
// pasted text against the words in each resource's filename.
//
// Deterministic word overlap only — no AI, no OCR, no per-lesson judgement
// calls (see AGENTS.md) — and the teacher always ticks the final list by
// hand; nothing here is auto-included in the output.

import { stripLeadingNumericPrefix } from "./resourceMatcher.js";
import type { ResourceEntry, ResourceMap } from "./types.js";

export interface ResourceCandidate {
  entry: ResourceEntry;
  score: number;
}

// Common English words that would otherwise match almost every filename and
// every lesson, without actually telling two lessons apart.
const STOPWORDS = new Set([
  "the",
  "and",
  "for",
  "with",
  "from",
  "this",
  "that",
  "your",
  "you",
  "are",
  "was",
  "were",
  "will",
  "have",
  "has",
  "had",
  "not",
  "but",
  "all",
  "can",
  "lesson",
  "lessons",
  "class",
  "notes",
  "note",
  "sheet",
  "file",
  "pdf",
  "docx",
  "pptx",
  "doc",
  "ppt",
]);

function tokenize(text: string): string[] {
  return text
    .toLowerCase()
    .split(/[^a-z0-9]+/)
    .filter((word) => word.length > 2 && !STOPWORDS.has(word));
}

function filenameWithoutExtension(relativePath: string): string {
  const name = relativePath.split("/").pop() ?? relativePath;
  const dot = name.lastIndexOf(".");
  return dot === -1 ? name : name.slice(0, dot);
}

/** Ranks every resource by how many distinct words its filename shares with
 * `lessonText`, highest first, and returns at most `limit` resources that
 * share at least one word. Ties break alphabetically so the order stays
 * predictable between searches. */
export function findResourceCandidates(
  lessonText: string,
  resources: ResourceMap,
  limit = 8,
): ResourceCandidate[] {
  const queryWords = new Set(tokenize(lessonText));
  if (queryWords.size === 0) return [];

  const candidates: ResourceCandidate[] = [];

  for (const entry of resources.values()) {
    const nameWords = new Set(
      tokenize(stripLeadingNumericPrefix(filenameWithoutExtension(entry.relativePath))),
    );

    let score = 0;
    for (const word of nameWords) {
      if (queryWords.has(word)) score += 1;
    }

    if (score > 0) candidates.push({ entry, score });
  }

  candidates.sort(
    (a, b) => b.score - a.score || a.entry.relativePath.localeCompare(b.entry.relativePath),
  );

  return candidates.slice(0, limit);
}
