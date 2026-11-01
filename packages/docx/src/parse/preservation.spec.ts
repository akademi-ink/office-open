import { unzipSync, zipSync } from "fflate";
import { describe, expect, it, vi } from "vite-plus/test";

import { generateDocumentSync } from "../generate";
import { parseDocument } from "../parse";
import type { DocumentOptions } from "../parts/core-properties";

const RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships";
const DOCUMENT_REL =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
const SIGNATURE_ORIGIN_REL =
  "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/origin";
const SIGNATURE_REL =
  "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/signature";
const DOCX_MAIN =
  "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
const DOCM_MAIN = "application/vnd.ms-word.document.macroEnabled.main+xml";
const DOTM_MAIN = "application/vnd.ms-word.template.macroEnabled.main+xml";

const DOCUMENT_XML =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
  "<w:body><w:p><w:r><w:t>Signed content</w:t></w:r></w:p></w:body></w:document>";

/** Original signature part bytes — deliberately carrying an obsolete namespace
 *  URI so the passthrough normalization exemption is observable. */
const ORIGIN_SIGS = new TextEncoder().encode(
  '<SignatureOrigin xmlns="http://schemas.openxmlformats.org/package/2006/digital-signature" ' +
    'probe="http://schemas.microsoft.com/office/word/2010/11/wordml"/>',
);
const SIGNATURE_PART = new TextEncoder().encode(
  '<Signature xmlns="http://www.w3.org/2000/09/xmldsig#"><SignedInfo>probe</SignedInfo></Signature>',
);
const ORIGIN_RELS = new TextEncoder().encode(
  `<?xml version="1.0"?><Relationships xmlns="${RELS_NS}">` +
    `<Relationship Id="rId1" Type="${SIGNATURE_REL}" Target="sig1.xml"/>` +
    `</Relationships>`,
);

const VBA = Uint8Array.from({ length: 96 }, (_, index) => (index * 13 + 7) % 256);

function encode(files: Record<string, string | Uint8Array>): Record<string, Uint8Array> {
  const out: Record<string, Uint8Array> = {};
  for (const [path, data] of Object.entries(files)) {
    out[path] = typeof data === "string" ? new TextEncoder().encode(data) : data;
  }
  return out;
}

function contentTypesXml(overrides: string, defaults = ""): string {
  return (
    '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    defaults +
    overrides +
    "</Types>"
  );
}

function rootRels(extra = ""): string {
  return (
    `<?xml version="1.0"?><Relationships xmlns="${RELS_NS}">` +
    `<Relationship Id="rId1" Type="${DOCUMENT_REL}" Target="word/document.xml"/>` +
    extra +
    "</Relationships>"
  );
}

/** A signed package: origin + one signature + their relationships. */
function signedPackage(): Uint8Array {
  const files = encode({
    "[Content_Types].xml": contentTypesXml(
      `<Override PartName="/word/document.xml" ContentType="${DOCX_MAIN}"/>` +
        '<Override PartName="/_xmlsignatures/sig1.xml" ContentType="application/vnd.openxmlformats-package.digital-signature-xmlsignature+xml"/>',
      '<Default Extension="sigs" ContentType="application/vnd.openxmlformats-package.digital-signature-origin"/>',
    ),
    "_rels/.rels": rootRels(
      `<Relationship Id="rId2" Type="${SIGNATURE_ORIGIN_REL}" Target="_xmlsignatures/origin.sigs"/>`,
    ),
    "word/document.xml": DOCUMENT_XML,
    "word/_rels/document.xml.rels": `<Relationships xmlns="${RELS_NS}"/>`,
    "_xmlsignatures/origin.sigs": ORIGIN_SIGS,
    "_xmlsignatures/_rels/origin.sigs.rels": ORIGIN_RELS,
    "_xmlsignatures/sig1.xml": SIGNATURE_PART,
  });
  return zipSync(files as never);
}

/** A macro-enabled package whose passthrough set carries a VBA project. */
function macroPackage(mainType: string = DOCM_MAIN): Uint8Array {
  const files = encode({
    "[Content_Types].xml": contentTypesXml(
      `<Override PartName="/word/document.xml" ContentType="${mainType}"/>`,
      '<Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/>',
    ),
    "_rels/.rels": rootRels(),
    "word/document.xml": DOCUMENT_XML,
    "word/_rels/document.xml.rels": `<Relationships xmlns="${RELS_NS}"/>`,
    "word/vbaProject.bin": VBA,
    "word/vbaData.xml": "<w:vbaData/>",
  });
  return zipSync(files as never);
}

function zipOf(bytes: Uint8Array): Record<string, Uint8Array> {
  return unzipSync(bytes) as Record<string, Uint8Array>;
}

function decode(bytes: Uint8Array | undefined): string {
  return bytes ? new TextDecoder().decode(bytes) : "";
}

describe("digital signatures — preservation", () => {
  it("detects the signature parts and their relationships", async () => {
    const parsed = await parseDocument(signedPackage());
    expect(parsed.digitalSignatures?.origin).toBe("_xmlsignatures/origin.sigs");
    expect(parsed.digitalSignatures?.signatures).toEqual(["_xmlsignatures/sig1.xml"]);
    expect(parsed.rawParts?.map((part) => part.path)).toEqual(
      expect.arrayContaining([
        "_xmlsignatures/origin.sigs",
        "_xmlsignatures/_rels/origin.sigs.rels",
        "_xmlsignatures/sig1.xml",
      ]),
    );
    expect(parsed.digitalSignatures?.relationships.map((rel) => rel.relationshipType)).toEqual([
      SIGNATURE_ORIGIN_REL,
      SIGNATURE_REL,
    ]);
  });

  it("re-emits every signature part byte-for-byte and keeps the declarations", async () => {
    const source = signedPackage();
    const write = vi.spyOn(console, "warn").mockImplementation(() => {});
    try {
      const generated = generateDocumentSync(await parseDocument(source)) as Uint8Array;
      const zip = zipOf(generated);

      // Byte equality for XML that the generic passthrough pass would otherwise
      // have normalized (the obsolete namespace URI in origin.sigs).
      expect(zip["_xmlsignatures/origin.sigs"]).toEqual(ORIGIN_SIGS);
      expect(zip["_xmlsignatures/_rels/origin.sigs.rels"]).toEqual(ORIGIN_RELS);
      expect(zip["_xmlsignatures/sig1.xml"]).toEqual(SIGNATURE_PART);

      // The signature declarations survive in the rebuilt package tables.
      expect(decode(zip["[Content_Types].xml"])).toContain('Extension="sigs"');
      expect(decode(zip["[Content_Types].xml"])).toContain('PartName="/_xmlsignatures/sig1.xml"');
      expect(decode(zip["_rels/.rels"])).toContain(`Type="${SIGNATURE_ORIGIN_REL}"`);
      expect(decode(zip["_rels/.rels"])).toContain('Target="_xmlsignatures/origin.sigs"');

      // A second parse/generate keeps the same bytes (transitional-package safe).
      const second = zipOf(generateDocumentSync(await parseDocument(generated)) as Uint8Array);
      expect(second["_xmlsignatures/origin.sigs"]).toEqual(ORIGIN_SIGS);
      expect(second["_xmlsignatures/sig1.xml"]).toEqual(SIGNATURE_PART);
    } finally {
      write.mockRestore();
    }
  });

  it("warns that regeneration invalidates the signature", async () => {
    const parsed = await parseDocument(signedPackage());
    const write = vi.spyOn(console, "warn").mockImplementation(() => {});
    try {
      generateDocumentSync(parsed);
      expect(write).toHaveBeenCalledTimes(1);
      const [message] = write.mock.calls[0] as [string];
      expect(message).toMatch(/digital signature/);
      expect(message).toMatch(/no longer validly signed/);
    } finally {
      write.mockRestore();
    }
  });

  it("keeps the signature bytes while the signed content changes", async () => {
    const parsed = await parseDocument(signedPackage());
    const write = vi.spyOn(console, "warn").mockImplementation(() => {});
    try {
      const zip = zipOf(
        generateDocumentSync({
          ...parsed,
          sections: [...parsed.sections, { children: [{ paragraph: { text: "edited" } }] }],
        }) as Uint8Array,
      );
      // The signature envelope is intact, but the document bytes it covers are
      // regenerated (and now include the edit) — the digests cannot match.
      expect(zip["_xmlsignatures/origin.sigs"]).toEqual(ORIGIN_SIGS);
      expect(decode(zip["word/document.xml"])).toContain("edited");
      expect(decode(zip["word/document.xml"])).not.toBe(DOCUMENT_XML);
    } finally {
      write.mockRestore();
    }
  });

  it("does not warn for an unsigned package", () => {
    const write = vi.spyOn(console, "warn").mockImplementation(() => {});
    try {
      const bytes = generateDocumentSync({
        sections: [{ children: [{ paragraph: { text: "plain" } }] }],
      }) as Uint8Array;
      expect(write).not.toHaveBeenCalled();
      expect(zipOf(bytes)["_xmlsignatures/sig1.xml"]).toBeUndefined();
    } finally {
      write.mockRestore();
    }
  });
});

describe("macros — byte-exact passthrough", () => {
  for (const [name, mainType] of [
    ["docm", DOCM_MAIN],
    ["dotm", DOTM_MAIN],
  ] as const) {
    it(`preserves vbaProject.bin and vbaData.xml for a .${name} save`, async () => {
      const parsed = await parseDocument(macroPackage(mainType));
      expect(parsed.rawParts?.some((part) => part.path === "word/vbaProject.bin")).toBe(true);

      const zip = zipOf(generateDocumentSync(parsed) as Uint8Array);
      expect(zip["word/vbaProject.bin"]).toEqual(VBA);
      expect(zip["word/vbaData.xml"]).toEqual(new TextEncoder().encode("<w:vbaData/>"));
      const contentTypes = decode(zip["[Content_Types].xml"]);
      expect(contentTypes).toContain(mainType);
      expect(contentTypes).toContain('ContentType="application/vnd.ms-office.vbaProject"');
    });
  }

  it("drops the VBA project when the output is a macro-free save", async () => {
    const parsed = await parseDocument(macroPackage());
    const macroFree: DocumentOptions = {
      ...parsed,
      // A Save As to .docx flips the main part's declared type; office-open has
      // no variant parameter, so the content type is the macro-enabled signal.
      contentTypes: {
        defaults: parsed.contentTypes?.defaults ?? [],
        overrides: (parsed.contentTypes?.overrides ?? []).map((override) =>
          override.partName.toLowerCase() === "/word/document.xml"
            ? { ...override, contentType: DOCX_MAIN }
            : override,
        ),
      },
    };

    const zip = zipOf(generateDocumentSync(macroFree) as Uint8Array);
    expect(zip["word/vbaProject.bin"]).toBeUndefined();
    expect(zip["word/vbaData.xml"]).toBeUndefined();
    const contentTypes = decode(zip["[Content_Types].xml"]);
    expect(contentTypes).not.toContain("macroEnabled");
    expect(contentTypes).not.toContain("/word/vbaProject.bin");
    expect(contentTypes).not.toContain("/word/vbaData.xml");
  });

  it("never emits a macro part in a fresh macro-free save", () => {
    const bytes = generateDocumentSync({
      sections: [{ children: [{ paragraph: { text: "no macros" } }] }],
    }) as Uint8Array;
    const zip = zipOf(bytes);
    expect(zip["word/vbaProject.bin"]).toBeUndefined();
    expect(decode(zip["[Content_Types].xml"])).not.toContain("macroEnabled");
  });
});
