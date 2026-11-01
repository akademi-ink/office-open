/**
 * Encrypted OOXML container detection and passthrough helpers.
 *
 * Password-protected OOXML files (ECMA-376 agile/standard encryption) wrap
 * the OPC package inside an OLE2/CFB compound file. The plaintext is not
 * recoverable without the password, so parse() carries the original bytes on
 * the options and generate() re-emits them verbatim instead of compiling.
 *
 * @module
 */

import type { DataType } from "../util/data-type";
import { toUint8Array } from "../util/data-type";
import type { OutputByType, OutputType } from "./output";
import { convertOutput } from "./output";

/** OLE2/CFB compound-file signature shared by every encrypted OOXML file. */
const CFB_SIGNATURE = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1] as const;

/**
 * ECMA-376 encryption flavor detected from a CFB envelope.
 *
 * - `agile` — ECMA-376 Part 4 §4.1.4 (AES-CBC + SHA-512). Word's default since
 *   Office 2010; the editor can decrypt it with a password.
 * - `standard` — the older CryptoAPI envelope (ECMA-376 Part 4 §4.2, versions
 *   3.2/4.2). Detection only: no decrypt path here.
 * - `unknown` — a compound file that is not an OOXML encrypted package
 *   (legacy `.doc`, an OLE container with other streams) or whose
 *   EncryptionInfo version is not one of the two known envelopes.
 */
export type EncryptionFlavor = "agile" | "standard" | "unknown";

/**
 * Precise status of a CFB-encrypted (password-protected) source package.
 *
 * Detection is best-effort by design: office-open never decrypts, so it reads
 * only the CFB framing (the `EncryptionInfo`/`EncryptedPackage` stream names
 * and the EncryptionInfo version header) to classify the envelope. A caller
 * that can decrypt (`@akademi-ink/docx`) uses this to decide whether to prompt,
 * carry the bytes verbatim, or refuse with a precise message.
 */
export interface EncryptedContainerStatus {
  /** Envelope flavor read from the EncryptionInfo version header. */
  flavor: EncryptionFlavor;
  /** Whether the compound file declares both OOXML envelope streams. */
  ooxmlPackage: boolean;
  /** Human-readable status — safe to surface instead of a generic archive error. */
  message: string;
}

/**
 * Verbatim payload of a source file whose OOXML package is encrypted.
 * Round-trip only: the bytes are the only faithful representation until the
 * file is decrypted; every other options field stays empty.
 */
export interface EncryptedContainerOptions {
  /** Complete source file bytes (OLE2/CFB container). */
  data: DataType;
  /** Envelope classification captured at parse time (absent when hand-authored). */
  status?: EncryptedContainerStatus;
}

/** ECMA-376 agile EncryptionInfo version header (4.4 + reserved 0x40). */
const AGILE_VERSION = [0x04, 0x00, 0x04, 0x00, 0x40, 0x00, 0x00, 0x00] as const;

/** Standard/CryptoAPI EncryptionInfo version headers (3.2 and 4.2). */
const STANDARD_VERSIONS = [
  [0x03, 0x00, 0x02, 0x00],
  [0x04, 0x00, 0x02, 0x00],
] as const;

/** Find a byte sequence in `data`; returns its offset or -1. */
function indexOfBytes(data: Uint8Array, needle: readonly number[], from = 0): number {
  const last = data.length - needle.length;
  outer: for (let i = from; i <= last; i++) {
    for (let j = 0; j < needle.length; j++) {
      if (data[i + j] !== needle[j]) continue outer;
    }
    return i;
  }
  return -1;
}

/** Whether a UTF-16LE encoding of `text` occurs in `data` (CFB directory names). */
function hasUtf16LeName(data: Uint8Array, text: string): boolean {
  const needle = new Uint8Array(text.length * 2);
  for (let i = 0; i < text.length; i++) {
    needle[i * 2] = text.charCodeAt(i) & 0xff;
    needle[i * 2 + 1] = text.charCodeAt(i) >> 8;
  }
  return indexOfBytes(data, Array.from(needle)) !== -1;
}

function encryptionStatusMessage(flavor: EncryptionFlavor, ooxmlPackage: boolean): string {
  if (!ooxmlPackage) {
    return (
      "This is an OLE2/CFB compound file but not an encrypted OOXML package " +
      "(no EncryptionInfo/EncryptedPackage streams) — legacy .doc and other OLE files are not readable here."
    );
  }
  switch (flavor) {
    case "agile":
      return (
        "This document is encrypted with ECMA-376 agile encryption (AES-CBC/SHA-512). " +
        "office-open does not decrypt; the original bytes are preserved verbatim for round-trip."
      );
    case "standard":
      return (
        "This document is encrypted with the Standard/CryptoAPI envelope, which is not supported here. " +
        "The original bytes are preserved verbatim; open it in Word and save a decrypted copy to edit."
      );
    default:
      return (
        "This document is an encrypted OOXML package with an unrecognized encryption version. " +
        "It is not supported here; the original bytes are preserved verbatim for round-trip."
      );
  }
}

/**
 * Classify a CFB-encrypted source package without decrypting it.
 *
 * - `undefined` for anything that is not a CFB container (a normal ZIP, or
 *   malformed input shorter than the signature).
 * - Otherwise a status whose `flavor` distinguishes the ECMA-376 agile
 *   envelope from the Standard/CryptoAPI one, and whose `message` names the
 *   supported path precisely. A legacy `.doc` (a CFB without the two OOXML
 *   envelope streams) reports `ooxmlPackage: false` instead of pretending to
 *   be an encrypted package.
 *
 * The CFB directory is not walked (office-open carries no compound-file
 * reader); detection uses the stream names that the CFB directory stores as
 * UTF-16LE and the EncryptionInfo version header. Synthetic headers are
 * enough to exercise the matrix; malformed or segmented envelopes fall back
 * to `unknown`, never to a wrong "decryptable" claim.
 */
export function detectEncryptedContainer(data: Uint8Array): EncryptedContainerStatus | undefined {
  if (!isEncryptedContainer(data)) return undefined;
  const ooxmlPackage =
    hasUtf16LeName(data, "EncryptionInfo") && hasUtf16LeName(data, "EncryptedPackage");
  let flavor: EncryptionFlavor = "unknown";
  if (ooxmlPackage) {
    if (indexOfBytes(data, AGILE_VERSION) !== -1) {
      flavor = "agile";
    } else if (STANDARD_VERSIONS.some((version) => indexOfBytes(data, version) !== -1)) {
      flavor = "standard";
    }
  }
  return { flavor, ooxmlPackage, message: encryptionStatusMessage(flavor, ooxmlPackage) };
}

/** Thrown by {@link assertNotEncryptedContainer} with the precise status. */
export class EncryptedPackageError extends Error {
  readonly code = "ENCRYPTED_PACKAGE" as const;

  constructor(readonly status: EncryptedContainerStatus) {
    super(status.message);
    this.name = "EncryptedPackageError";
  }
}

/**
 * Reject a CFB container before any ZIP/archive work runs, with the precise
 * status instead of the generic "not a zip" error the archive reader would
 * otherwise surface. Round-trip callers (`parseDocument`) detect first and
 * keep the bytes; this guard is for the raw archive entry points.
 */
export function assertNotEncryptedContainer(data: Uint8Array): void {
  const status = detectEncryptedContainer(data);
  if (status) throw new EncryptedPackageError(status);
}

/**
 * Whether `data` is an OLE2/CFB container. Encrypted OOXML files and legacy
 * `.doc` files share this signature; {@link detectEncryptedContainer}
 * distinguishes an OOXML envelope from a generic compound file.
 */
export function isEncryptedContainer(data: Uint8Array): boolean {
  return data.length >= CFB_SIGNATURE.length && CFB_SIGNATURE.every((b, i) => data[i] === b);
}

/**
 * Converted output for an encrypted-container file, or `undefined` when the
 * options carry no encrypted payload (the normal compile path applies).
 */
export function encryptedContainerOutput<T extends OutputType>(
  file: { encrypted?: EncryptedContainerOptions },
  type: T,
  mimeType: string,
): OutputByType[T] | undefined {
  const { encrypted } = file;
  if (!encrypted) return undefined;
  return convertOutput(toUint8Array(encrypted.data), type, mimeType);
}

/** Streaming variant of {@link encryptedContainerOutput}. */
export function encryptedContainerStream(file: {
  encrypted?: EncryptedContainerOptions;
}): ReadableStream<Uint8Array> | undefined {
  const { encrypted } = file;
  if (!encrypted) return undefined;
  const bytes = toUint8Array(encrypted.data);
  return new ReadableStream<Uint8Array>({
    start(c) {
      c.enqueue(bytes);
      c.close();
    },
  });
}

/**
 * Reject encrypted passthrough mixed with real content: generate() re-emits
 * the encrypted bytes, so every other field would be silently dropped.
 *
 * `hasContent` is the package's primary content collection being non-empty
 * (e.g. `sections.length > 0`) — callers know their own shape.
 */
export function assertEncryptedExclusive(
  file: { encrypted?: EncryptedContainerOptions },
  hasContent: boolean,
): void {
  if (file.encrypted && hasContent) {
    throw new Error(
      "Encrypted passthrough carries no editable content — clear the content fields or drop `encrypted`.",
    );
  }
}
