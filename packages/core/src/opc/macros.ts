/**
 * Macro-enabled package classification.
 *
 * A `.docm` / `.dotm` package is a normal OPC package whose main-document
 * content type ends in `macroEnabled.main+xml` and which carries a VBA project
 * part (`word/vbaProject.bin`, plus optionally `vbaData.xml` and
 * `vbaProjectSignature.bin`). office-open never executes macros: it classifies
 * the package so the compiler can keep the project byte-for-byte on a
 * macro-enabled save and drop it from a macro-free save (a `.docx` must not
 * carry a VBA project).
 *
 * @module
 */

import type { ContentTypesInput } from "./content-types-input";

/** Basenames of the parts that make up a VBA project (case-insensitive). */
const MACRO_PART_BASENAMES: ReadonlySet<string> = new Set([
  "vbaproject.bin",
  "vbadata.xml",
  "vbaprojectsignature.bin",
]);

/**
 * Whether a main-document content type declares a macro-enabled package.
 * `docm` and `dotm` both end in `macroEnabled.main+xml`.
 */
export function isMacroEnabledContentType(contentType: string | undefined): boolean {
  return contentType !== undefined && /macroenabled\.main\+xml$/i.test(contentType);
}

/** Whether a package part path is a VBA project part (`word/vbaProject.bin`, …). */
export function isMacroPartPath(path: string): boolean {
  const slash = path.lastIndexOf("/");
  return MACRO_PART_BASENAMES.has(path.slice(slash + 1).toLowerCase());
}

/**
 * The main document part's declared content type from a parsed
 * `[Content_Types].xml`, or `undefined` when the table has no override for
 * `/word/document.xml` (a fresh compile, or a source that relied on a Default).
 */
export function mainDocumentContentType(
  contentTypes: ContentTypesInput | undefined,
): string | undefined {
  return contentTypes?.overrides.find(
    (override) => override.partName.toLowerCase() === "/word/document.xml",
  )?.contentType;
}

/** Whether a parsed package's content-type table declares a macro-enabled main part. */
export function isMacroEnabledPackage(contentTypes: ContentTypesInput | undefined): boolean {
  return isMacroEnabledContentType(mainDocumentContentType(contentTypes));
}

/**
 * Drop VBA project parts from a passthrough list. The compiler calls this when
 * the output's main-document type is macro-free, so a `.docx` save of a
 * macro-enabled source cannot carry an orphaned `vbaProject.bin`.
 */
export function withoutMacroParts<T extends { path: string }>(parts: readonly T[]): T[] {
  return parts.filter((part) => !isMacroPartPath(part.path));
}
