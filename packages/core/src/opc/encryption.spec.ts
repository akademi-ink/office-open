import { describe, expect, it } from "vitest";

import {
  assertEncryptedExclusive,
  assertNotEncryptedContainer,
  detectEncryptedContainer,
  EncryptedPackageError,
  encryptedContainerOutput,
  encryptedContainerStream,
  isEncryptedContainer,
} from "./encryption";
import { OoxmlMimeType } from "./output";
import { parseArchive } from "./parser";

const CFB_BYTES = new Uint8Array([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1, 0x00, 0x00]);

const CFB_SIGNATURE = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1] as const;

/** UTF-16LE bytes of a CFB directory entry name. */
function utf16le(text: string): Uint8Array {
  const out = new Uint8Array(text.length * 2);
  for (let i = 0; i < text.length; i++) {
    out[i * 2] = text.charCodeAt(i) & 0xff;
    out[i * 2 + 1] = text.charCodeAt(i) >> 8;
  }
  return out;
}

/**
 * Synthetic CFB envelope: a 512-byte header carrying the compound-file
 * signature, the OOXML envelope stream names as UTF-16LE directory entries,
 * and the EncryptionInfo version header of the requested flavor. Enough to
 * exercise the detection matrix without building a real compound file.
 */
function cfbEnvelope(
  flavor: "agile" | "standard" | "none",
  options: { streams?: boolean } = {},
): Uint8Array {
  const chunks: Uint8Array[] = [new Uint8Array(512)];
  chunks[0]!.set(CFB_SIGNATURE, 0);
  if (options.streams !== false) {
    chunks.push(utf16le("EncryptionInfo"), utf16le("EncryptedPackage"));
  }
  if (flavor === "agile") {
    chunks.push(Uint8Array.from([0x04, 0x00, 0x04, 0x00, 0x40, 0x00, 0x00, 0x00]));
    chunks.push(new TextEncoder().encode("<encryption/>"));
  } else if (flavor === "standard") {
    chunks.push(Uint8Array.from([0x03, 0x00, 0x02, 0x00, 0x00, 0x00, 0x00, 0x00]));
  }
  const total = chunks.reduce((n, chunk) => n + chunk.length, 0);
  const bytes = new Uint8Array(total);
  let offset = 0;
  for (const chunk of chunks) {
    bytes.set(chunk, offset);
    offset += chunk.length;
  }
  return bytes;
}

describe("isEncryptedContainer", () => {
  it("matches the OLE2/CFB signature", () => {
    expect(isEncryptedContainer(CFB_BYTES)).toBe(true);
  });

  it("rejects ZIP packages and truncated inputs", () => {
    expect(isEncryptedContainer(new Uint8Array([0x50, 0x4b, 0x03, 0x04]))).toBe(false);
    expect(isEncryptedContainer(CFB_BYTES.subarray(0, 7))).toBe(false);
    expect(isEncryptedContainer(new Uint8Array())).toBe(false);
  });
});

describe("encryptedContainerOutput", () => {
  it("converts the verbatim payload to the requested output type", async () => {
    const out = encryptedContainerOutput(
      { encrypted: { data: CFB_BYTES } },
      "uint8array",
      OoxmlMimeType.DOCX,
    );
    expect(out).toBe(CFB_BYTES);

    const base64 = encryptedContainerOutput(
      { encrypted: { data: Array.from(CFB_BYTES) } },
      "base64",
      OoxmlMimeType.XLSX,
    );
    expect(typeof base64).toBe("string");
  });

  it("returns undefined without an encrypted payload", () => {
    expect(encryptedContainerOutput({}, "uint8array", OoxmlMimeType.DOCX)).toBeUndefined();
  });
});

describe("assertEncryptedExclusive", () => {
  it("rejects encrypted passthrough mixed with real content", () => {
    expect(() => assertEncryptedExclusive({ encrypted: { data: CFB_BYTES } }, true)).toThrow(
      /Encrypted passthrough/,
    );
  });

  it("allows the empty passthrough shape and normal files", () => {
    expect(() => assertEncryptedExclusive({ encrypted: { data: CFB_BYTES } }, false)).not.toThrow();
    expect(() => assertEncryptedExclusive({}, true)).not.toThrow();
  });
});

describe("encryptedContainerStream", () => {
  it("streams the payload unchanged", async () => {
    const stream = encryptedContainerStream({ encrypted: { data: CFB_BYTES } });
    expect(stream).toBeDefined();
    const chunks: Uint8Array[] = [];
    const reader = stream!.getReader();
    for (;;) {
      const { done, value } = await reader.read();
      if (done) break;
      chunks.push(value);
    }
    const total = chunks.reduce((n, c) => n + c.length, 0);
    const bytes = new Uint8Array(total);
    let off = 0;
    for (const c of chunks) {
      bytes.set(c, off);
      off += c.length;
    }
    expect(bytes).toEqual(CFB_BYTES);
    expect(encryptedContainerStream({})).toBeUndefined();
  });
});

describe("detectEncryptedContainer — encryption matrix", () => {
  it("classifies the ECMA-376 agile envelope", () => {
    const status = detectEncryptedContainer(cfbEnvelope("agile"));
    expect(status).toEqual({
      flavor: "agile",
      ooxmlPackage: true,
      message: expect.stringMatching(/agile encryption/i),
    });
  });

  it("classifies the Standard/CryptoAPI envelope as unsupported here", () => {
    const status = detectEncryptedContainer(cfbEnvelope("standard"));
    expect(status?.flavor).toBe("standard");
    expect(status?.ooxmlPackage).toBe(true);
    expect(status?.message).toMatch(/Standard\/CryptoAPI/);
    expect(status?.message).toMatch(/not supported here/);
    expect(status?.message).toMatch(/decrypted copy/);
  });

  it("does not mistake a legacy .doc compound file for an encrypted OOXML package", () => {
    const status = detectEncryptedContainer(cfbEnvelope("none", { streams: false }));
    expect(status).toEqual({
      flavor: "unknown",
      ooxmlPackage: false,
      message: expect.stringMatching(/not an encrypted OOXML package/),
    });
  });

  it("reports an unrecognized envelope version as unknown, never as decryptable", () => {
    const status = detectEncryptedContainer(cfbEnvelope("none"));
    expect(status?.flavor).toBe("unknown");
    expect(status?.ooxmlPackage).toBe(true);
    expect(status?.message).toMatch(/unrecognized encryption version/);
  });

  it("returns undefined for ZIP packages and short input", () => {
    expect(detectEncryptedContainer(new Uint8Array([0x50, 0x4b, 0x03, 0x04]))).toBeUndefined();
    expect(detectEncryptedContainer(CFB_BYTES.subarray(0, 7))).toBeUndefined();
  });
});

describe("assertNotEncryptedContainer", () => {
  it("throws the precise status instead of a generic archive error", () => {
    let thrown: unknown;
    try {
      assertNotEncryptedContainer(cfbEnvelope("agile"));
    } catch (error) {
      thrown = error;
    }
    expect(thrown).toBeInstanceOf(EncryptedPackageError);
    const error = thrown as EncryptedPackageError;
    expect(error.code).toBe("ENCRYPTED_PACKAGE");
    expect(error.status.flavor).toBe("agile");
    expect(error.message).toMatch(/agile encryption/);
  });

  it("lets a ZIP package through", () => {
    expect(() =>
      assertNotEncryptedContainer(new Uint8Array([0x50, 0x4b, 0x03, 0x04])),
    ).not.toThrow();
  });

  it("makes parseArchive report the encrypted status, not a zip rejection", () => {
    expect(() => parseArchive(cfbEnvelope("standard"))).toThrow(EncryptedPackageError);
  });
});
