import { describe, expect, it } from "vite-plus/test";

import { generateDocumentSync } from "../generate";
import { parseDocument, parseDocumentSync } from "../parse";

const CFB_SIGNATURE = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1] as const;

function utf16le(text: string): Uint8Array {
  const out = new Uint8Array(text.length * 2);
  for (let i = 0; i < text.length; i++) {
    out[i * 2] = text.charCodeAt(i) & 0xff;
    out[i * 2 + 1] = text.charCodeAt(i) >> 8;
  }
  return out;
}

/** Synthetic CFB wrapper: header + OOXML envelope stream names + version header. */
function cfbEnvelope(flavor: "agile" | "standard" | "legacy"): Uint8Array {
  const chunks: Uint8Array[] = [new Uint8Array(512)];
  chunks[0]!.set(CFB_SIGNATURE, 0);
  if (flavor !== "legacy") {
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

describe("encrypted packages — precise status", () => {
  it("reports the agile envelope and preserves the bytes verbatim", async () => {
    const source = cfbEnvelope("agile");
    const parsed = await parseDocument(source);
    expect(parsed.sections).toEqual([]);
    expect(parsed.encrypted?.status?.flavor).toBe("agile");
    expect(parsed.encrypted?.status?.ooxmlPackage).toBe(true);
    expect(parsed.encrypted?.status?.message).toMatch(/agile encryption/i);

    // office-open is a passthrough here: generate re-emits the source bytes.
    expect(generateDocumentSync(parsed, { type: "uint8array" })).toEqual(source);
  });

  it("reports the Standard/CryptoAPI envelope as unsupported, not as a zip error", () => {
    const parsed = parseDocumentSync(cfbEnvelope("standard"));
    expect(parsed.encrypted?.status?.flavor).toBe("standard");
    expect(parsed.encrypted?.status?.message).toMatch(/Standard\/CryptoAPI/);
    expect(parsed.encrypted?.status?.message).toMatch(/not supported here/);
  });

  it("does not claim a legacy .doc compound file is an encrypted OOXML package", () => {
    const parsed = parseDocumentSync(cfbEnvelope("legacy"));
    expect(parsed.encrypted?.status).toEqual({
      flavor: "unknown",
      ooxmlPackage: false,
      message: expect.stringMatching(/not an encrypted OOXML package/),
    });
  });

  it("keeps encrypted passthrough exclusive of real content", async () => {
    const parsed = await parseDocument(cfbEnvelope("agile"));
    expect(() =>
      generateDocumentSync({ ...parsed, sections: [{ children: [{ paragraph: { text: "x" } }] }] }),
    ).toThrow(/Encrypted passthrough/);
  });
});
