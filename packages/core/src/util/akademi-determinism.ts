/**
 * Akademi determinism seam. When @akademi-ink/docx opens a generation scope
 * under `globalThis[Symbol.for("akademi.ooxml.determinism")]`, identifier and
 * clock reads resolve deterministically for that generation (the same input
 * yields the same bytes) instead of touching process-global counters, crypto
 * randomness or the wall clock. Outside a scope nothing changes.
 *
 * @module
 */

export interface AkademiGenerationScope {
  nextToken(): string;
  nextUuid(): string;
  nextVml(): number;
  /** Fixed ISO timestamp for docProps created/modified. */
  readonly now?: string;
}

const SCOPE_KEY = Symbol.for("akademi.ooxml.determinism");
const FORCE_JS_DEFLATE_KEY = Symbol.for("akademi.ooxml.force-js-deflate");

export const akademiGenerationScope = (): AkademiGenerationScope | undefined =>
  (globalThis as Record<symbol, unknown>)[SCOPE_KEY] as AkademiGenerationScope | undefined;

/** A reproducibility test may force the pure-JS (fflate) DEFLATE path on Node. */
export const akademiForcesJsDeflate = (): boolean =>
  Boolean((globalThis as Record<symbol, unknown>)[FORCE_JS_DEFLATE_KEY]);

/** Fixed DOS-epoch timestamp for reproducible archives. */
export const AKADEMI_ZIP_MTIME = new Date(1980, 0, 1);
