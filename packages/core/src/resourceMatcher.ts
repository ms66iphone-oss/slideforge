// Matches a filename referenced in lesson HTML ([[image:...]]/[[resource:...]])
// against what's actually in the resources zip. `resources` here is a
// Map<normalizedRelativePath, entry> built by zipImport.ts — already
// flattened across every folder depth in the zip, so no separate recursive
// walk is needed (JSZip gives every entry's full path up front).
//
// Confirmed against a real Seqta export: the lesson HTML references the
// original filename, but Seqta's own resource export prefixes every
// extracted filename with a numeric resource ID + underscore, e.g.
// "33_Screen Shot 2021-12-20 at 2.50.53 pm.png". Exact match is tried first
// so already-correct filenames are unaffected.

import type { ResourceEntry, ResourceMap } from "./types.js";

export function matchResource(filename: string, resources: ResourceMap): ResourceEntry | null {
  const target = normalize(filename);

  // 1. Exact match (already-correct filenames).
  for (const entry of resources.values()) {
    if (normalize(lastPathComponent(entry.relativePath)) === target) return entry;
  }

  // 2. Seqta's "<resourceID>_<originalFilename>" export naming.
  for (const entry of resources.values()) {
    if (normalize(stripLeadingNumericPrefix(lastPathComponent(entry.relativePath))) === target) {
      return entry;
    }
  }

  return null;
}

function lastPathComponent(path: string): string {
  const parts = path.split("/");
  return parts[parts.length - 1] ?? path;
}

/** Strips a leading "<digits>_" prefix, e.g. "33_photo.png" -> "photo.png". */
export function stripLeadingNumericPrefix(name: string): string {
  const underscoreIndex = name.indexOf("_");
  if (underscoreIndex === -1) return name;
  const prefix = name.slice(0, underscoreIndex);
  if (!prefix || !/^[0-9]+$/.test(prefix)) return name;
  return name.slice(underscoreIndex + 1);
}

function normalize(raw: string): string {
  return raw.trim().toLowerCase();
}
