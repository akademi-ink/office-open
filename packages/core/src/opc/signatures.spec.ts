import { describe, expect, it } from "vitest";

import { zipSync } from "./packer";
import { parseArchive } from "./parser";
import {
  detectDigitalSignatures,
  digitalSignatureInvalidationWarning,
  DIGITAL_SIGNATURE_ORIGIN_REL,
  DIGITAL_SIGNATURE_SIGNATURE_REL,
  isDigitalSignaturePartPath,
} from "./signatures";

function archiveOf(files: Record<string, string | Uint8Array>): ReturnType<typeof parseArchive> {
  const zippable: Record<string, Uint8Array> = {};
  for (const [path, data] of Object.entries(files)) {
    zippable[path] = typeof data === "string" ? new TextEncoder().encode(data) : data;
  }
  return parseArchive(zipSync(zippable as never));
}

const CONTENT_TYPES =
  `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
  `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
  `<Default Extension="xml" ContentType="application/xml"/>` +
  `<Default Extension="sigs" ContentType="application/vnd.openxmlformats-package.digital-signature-origin"/>` +
  `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>` +
  `<Override PartName="/_xmlsignatures/sig1.xml" ContentType="application/vnd.openxmlformats-package.digital-signature-xmlsignature+xml"/>` +
  `</Types>`;

const ROOT_RELS =
  `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
  `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>` +
  `<Relationship Id="rId2" Type="${DIGITAL_SIGNATURE_ORIGIN_REL}" Target="_xmlsignatures/origin.sigs"/>` +
  `</Relationships>`;

const ORIGIN_RELS =
  `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
  `<Relationship Id="rId1" Type="${DIGITAL_SIGNATURE_SIGNATURE_REL}" Target="sig1.xml"/>` +
  `</Relationships>`;

function signedPackage(): Record<string, string | Uint8Array> {
  return {
    "[Content_Types].xml": CONTENT_TYPES,
    "_rels/.rels": ROOT_RELS,
    "word/document.xml": "<w:document/>",
    "word/_rels/document.xml.rels": "<Relationships/>",
    "_xmlsignatures/origin.sigs": "<SignatureOrigin/>",
    "_xmlsignatures/_rels/origin.sigs.rels": ORIGIN_RELS,
    "_xmlsignatures/sig1.xml": "<Signature/>",
  };
}

describe("isDigitalSignaturePartPath", () => {
  it("matches the _xmlsignatures directory case-insensitively", () => {
    expect(isDigitalSignaturePartPath("_xmlsignatures/origin.sigs")).toBe(true);
    expect(isDigitalSignaturePartPath("_XMLSignatures/sig1.xml")).toBe(true);
    expect(isDigitalSignaturePartPath("word/document.xml")).toBe(false);
  });
});

describe("detectDigitalSignatures", () => {
  it("returns undefined for an unsigned package", () => {
    const files = signedPackage();
    delete files["_xmlsignatures/origin.sigs"];
    delete files["_xmlsignatures/_rels/origin.sigs.rels"];
    delete files["_xmlsignatures/sig1.xml"];
    expect(detectDigitalSignatures(archiveOf(files))).toBeUndefined();
  });

  it("collects the origin, signature parts, and both relationship hops", () => {
    const info = detectDigitalSignatures(archiveOf(signedPackage()));
    expect(info).toBeDefined();
    expect(info?.origin).toBe("_xmlsignatures/origin.sigs");
    expect(info?.signatures).toEqual(["_xmlsignatures/sig1.xml"]);
    expect(info?.parts).toEqual([
      "_xmlsignatures/origin.sigs",
      "_xmlsignatures/_rels/origin.sigs.rels",
      "_xmlsignatures/sig1.xml",
    ]);
    expect(info?.relationships).toEqual([
      {
        source: "",
        relationshipType: DIGITAL_SIGNATURE_ORIGIN_REL,
        target: "_xmlsignatures/origin.sigs",
        resolvedTarget: "_xmlsignatures/origin.sigs",
        rId: "rId2",
      },
      {
        source: "_xmlsignatures/origin.sigs",
        relationshipType: DIGITAL_SIGNATURE_SIGNATURE_REL,
        target: "sig1.xml",
        resolvedTarget: "_xmlsignatures/sig1.xml",
        rId: "rId1",
      },
    ]);
  });

  it("keeps multiple signature parts in package order", () => {
    const files = signedPackage();
    files["_xmlsignatures/sig2.xml"] = "<Signature2/>";
    const info = detectDigitalSignatures(archiveOf(files));
    expect(info?.signatures).toEqual(["_xmlsignatures/sig1.xml", "_xmlsignatures/sig2.xml"]);
  });

  it("does not fail the parse on a malformed signature .rels", () => {
    const files = signedPackage();
    files["_xmlsignatures/_rels/origin.sigs.rels"] = "<Relationships";
    const info = detectDigitalSignatures(archiveOf(files));
    expect(info?.parts).toContain("_xmlsignatures/sig1.xml");
    expect(info?.relationships).toEqual([
      expect.objectContaining({ source: "", relationshipType: DIGITAL_SIGNATURE_ORIGIN_REL }),
    ]);
  });

  it("explains that regeneration invalidates signatures", () => {
    const info = detectDigitalSignatures(archiveOf(signedPackage()))!;
    expect(digitalSignatureInvalidationWarning(info)).toMatch(/1 digital signature/);
    expect(digitalSignatureInvalidationWarning(info)).toMatch(/no longer validly signed/);
  });
});
