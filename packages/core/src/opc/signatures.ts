/**
 * OPC digital-signature detection.
 *
 * A signed OOXML package carries `_xmlsignatures/`: an origin part
 * (`origin.sigs`), one or more signature parts (`sigN.xml`), their companion
 * `.rels`, and package/origin relationships pointing at them. office-open does
 * not sign or verify. It detects the parts so `parseDocument` can report the
 * signed state, the passthrough collection can carry every signature part
 * byte-for-byte, and `compileDocument` can warn that regeneration changes the
 * signed content and therefore invalidates the package's signatures.
 *
 * @module
 */

import { attr } from "@office-open/xml";
import type { Element } from "@office-open/xml";

import type { ParsedArchive } from "./parser";
import { resolveRelationshipTarget } from "./relationships";

/** Package relationship type identifying the digital-signature origin part. */
export const DIGITAL_SIGNATURE_ORIGIN_REL =
  "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/origin";

/** Signature-part relationship type used by `_xmlsignatures/_rels/origin.sigs.rels`. */
export const DIGITAL_SIGNATURE_SIGNATURE_REL =
  "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/signature";

/** Directory every OPC digital-signature part lives in. */
export const DIGITAL_SIGNATURE_DIRECTORY = "_xmlsignatures/";

/** One relationship that references a signature part. */
export interface DigitalSignatureRelationship {
  /** Owner part path; `""` for the package root `_rels/.rels`. */
  source: string;
  relationshipType: string;
  /** Target exactly as written in the source `.rels`. */
  target: string;
  /** Target resolved against the owner part (the package part path). */
  resolvedTarget: string;
  /** Relationship id as written in the source `.rels`. */
  rId: string;
}

/** Digital-signature parts and the relationships that reference them. */
export interface DigitalSignatureInfo {
  /** Every `_xmlsignatures/` part, in package order (bytes are preserved verbatim). */
  parts: string[];
  /** Signature origin part path (`_xmlsignatures/origin.sigs` when present). */
  origin?: string;
  /** Signature part paths (`_xmlsignatures/sigN.xml`), package order. */
  signatures: string[];
  /** Package-root → origin and origin → signature relationships. */
  relationships: DigitalSignatureRelationship[];
}

/** Whether a package part path belongs to the `_xmlsignatures/` directory. */
export function isDigitalSignaturePartPath(path: string): boolean {
  return path.toLowerCase().startsWith(DIGITAL_SIGNATURE_DIRECTORY);
}

/** Read every signature-referencing relationship from a `.rels` part. */
function readSignatureRelationships(
  archive: ParsedArchive,
  relsPath: string,
  source: string,
): DigitalSignatureRelationship[] {
  // Signature .rels are untrusted input like every other part; a malformed one
  // must not fail the whole document parse — it simply contributes no entries.
  let relsEl: Element | undefined;
  try {
    relsEl = archive.get(relsPath);
  } catch {
    return [];
  }
  if (!relsEl) return [];
  const relationships: DigitalSignatureRelationship[] = [];
  for (const rel of relsEl.elements ?? []) {
    if (rel.name !== "Relationship") continue;
    const relationshipType = attr(rel, "Type");
    const target = attr(rel, "Target");
    const rId = attr(rel, "Id");
    if (!relationshipType || !target || !rId) continue;
    if (!relationshipType.includes("/digital-signature/")) continue;
    relationships.push({
      source,
      relationshipType,
      target,
      resolvedTarget: resolveRelationshipTarget(source, target),
      rId,
    });
  }
  return relationships;
}

/**
 * Detect digital-signature parts and their relationships.
 *
 * Returns `undefined` for an unsigned package. `parts` is every part under
 * `_xmlsignatures/` (origin, signatures, and their `.rels`); relationships are
 * read from `_rels/.rels` (origin) and every `_xmlsignatures/_rels/*.rels`
 * (signature entries). Matching is case-insensitive, as OPC part names are.
 */
export function detectDigitalSignatures(archive: ParsedArchive): DigitalSignatureInfo | undefined {
  // OPC part names are case-insensitive; keep the archive's spelling for the
  // byte-preserving passthrough while matching on the lowercased form.
  const parts: Array<{ path: string; lower: string }> = [];
  for (const path of archive.keys()) {
    if (isDigitalSignaturePartPath(path)) parts.push({ path, lower: path.toLowerCase() });
  }
  if (parts.length === 0) return undefined;

  const origin = parts.find((part) => part.lower.endsWith("/origin.sigs"))?.path;
  const signatures = parts
    .filter(
      (part) =>
        part.path !== origin && !part.lower.includes("/_rels/") && part.lower.endsWith(".xml"),
    )
    .map((part) => part.path);

  const relationships = readSignatureRelationships(archive, "_rels/.rels", "");
  for (const part of parts) {
    if (!part.lower.includes("/_rels/")) continue;
    // A companion .rels is owned by the part it describes, not by the .rels
    // path itself — resolve and report against that owner.
    const owner = part.path.replace(/\/_rels\/(.+)\.rels$/, "/$1");
    relationships.push(...readSignatureRelationships(archive, part.path, owner));
  }

  return { parts: parts.map((part) => part.path), origin, signatures, relationships };
}

/**
 * Explain why regenerating a signed package invalidates its signatures.
 * Returned as a string so callers can log, dialog, or snapshot it.
 */
export function digitalSignatureInvalidationWarning(info: DigitalSignatureInfo): string {
  const count = Math.max(1, info.signatures.length);
  return (
    `This document carries ${count} digital signature${count === 1 ? "" : "s"} ` +
    `(${DIGITAL_SIGNATURE_DIRECTORY}). Regeneration changes the signed content, so the ` +
    "generated package is no longer validly signed; the signature parts are preserved " +
    "byte-for-byte, but signing and verification are not performed."
  );
}

/**
 * Warn (once per compile) that generating a package from parsed signed
 * options invalidates the source signatures. Uses the host console when one
 * exists, so Node, browsers, and workers all take the same path.
 */
export function warnDigitalSignatureInvalidation(info: DigitalSignatureInfo | undefined): void {
  if (!info) return;
  globalThis.console?.warn(`[office-open] ${digitalSignatureInvalidationWarning(info)}`);
}
