import { strToU8, zipSync } from "fflate";
import { describe, expect, it } from "vite-plus/test";

import { generateDocument } from "../generate";
import { parseDocumentSync } from "../parse";

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const BIB_NS = "http://schemas.openxmlformats.org/officeDocument/2006/bibliography";
const BIB_2004_NS = "http://schemas.microsoft.com/office/word/2004/10/bibliography";
const CUSTOM_XML_REL =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml";
const BIBLIOGRAPHY_REL =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/bibliography";

/** A minimal package with a document body and the supplied extra parts. */
function packageOf(parts: Record<string, string>): Uint8Array {
  const files: Record<string, Uint8Array> = {
    "[Content_Types].xml": strToU8(
      '<?xml version="1.0" encoding="UTF-8"?>' +
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
        '<Default Extension="xml" ContentType="application/xml"/>' +
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
        "</Types>",
    ),
    "_rels/.rels": strToU8(
      '<?xml version="1.0" encoding="UTF-8"?>' +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
        "</Relationships>",
    ),
    "word/document.xml": strToU8(
      `<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="${W_NS}"><w:body><w:p/></w:body></w:document>`,
    ),
  };
  for (const [path, xml] of Object.entries(parts)) files[path] = strToU8(xml);
  return zipSync(files);
}

function documentRels(...relationships: string[]): string {
  return (
    '<?xml version="1.0" encoding="UTF-8"?>' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    relationships.join("") +
    "</Relationships>"
  );
}

function customXmlPart(xml: string): Record<string, string> {
  return {
    "word/_rels/document.xml.rels": documentRels(
      `<Relationship Id="rId1" Type="${CUSTOM_XML_REL}" Target="customXml/item1.xml"/>`,
    ),
    "word/customXml/item1.xml": xml,
  };
}

describe("bibliography part detection", () => {
  it("parses a default-namespace customXml bibliography part", () => {
    const parsed = parseDocumentSync(
      packageOf(
        customXmlPart(
          `<Sources xmlns="${BIB_NS}"><Source><SourceType>Book</SourceType>` +
            "<Title>Default Namespace</Title></Source></Sources>",
        ),
      ),
    );
    expect(parsed.bibliography?.sources).toHaveLength(1);
    expect(parsed.bibliography?.sources[0]?.sourceType).toBe("Book");
    expect(parsed.bibliography?.sources[0]?.title).toBe("Default Namespace");
  });

  it("parses a customXml part with a non-canonical prefix", () => {
    const parsed = parseDocumentSync(
      packageOf(
        customXmlPart(
          `<bib:Sources xmlns:bib="${BIB_NS}"><bib:Source>` +
            "<bib:Title>Alternate Prefix</bib:Title></bib:Source></bib:Sources>",
        ),
      ),
    );
    expect(parsed.bibliography?.sources[0]?.title).toBe("Alternate Prefix");
  });

  it("parses the Word 2004 bibliography namespace", () => {
    const parsed = parseDocumentSync(
      packageOf(
        customXmlPart(
          `<Sources xmlns="${BIB_2004_NS}"><Source>` +
            "<Title>Word 2004</Title></Source></Sources>",
        ),
      ),
    );
    expect(parsed.bibliography?.sources[0]?.title).toBe("Word 2004");
  });

  it("preserves b:Sources/@Version through the package parse", () => {
    const parsed = parseDocumentSync(
      packageOf(
        customXmlPart(
          `<b:Sources xmlns:b="${BIB_NS}" Version="6">` +
            "<b:Source><b:Title>Versioned</b:Title></b:Source></b:Sources>",
        ),
      ),
    );
    expect(parsed.bibliography?.version).toBe("6");
    expect(parsed.bibliography?.sources[0]?.title).toBe("Versioned");
  });

  it("still parses the legacy word/bibliography.xml relationship shape", () => {
    const parsed = parseDocumentSync(
      packageOf({
        "word/_rels/document.xml.rels": documentRels(
          `<Relationship Id="rId1" Type="${BIBLIOGRAPHY_REL}" Target="bibliography.xml"/>`,
        ),
        "word/bibliography.xml":
          `<b:Sources xmlns:b="${BIB_NS}">` +
          "<b:Source><b:Title>Legacy Part</b:Title></b:Source></b:Sources>",
      }),
    );
    expect(parsed.bibliography?.sources).toHaveLength(1);
    expect(parsed.bibliography?.sources[0]?.title).toBe("Legacy Part");
  });

  it("parses back its own generated customXml bibliography part", async () => {
    const output = await generateDocument(
      {
        sections: [{ children: [{ paragraph: { children: [{ text: "Body" }] } }] }],
        bibliography: {
          sources: [{ sourceType: "Book", title: "Generated", tag: "G1" }],
          version: "6",
        },
      },
      { type: "uint8array" },
    );
    const parsed = parseDocumentSync(output);
    expect(parsed.bibliography?.sources[0]?.title).toBe("Generated");
    expect(parsed.bibliography?.version).toBe("6");
  });

  it("ignores a customXml part whose Sources root is in another namespace", () => {
    const parsed = parseDocumentSync(
      packageOf(
        customXmlPart(
          '<Sources xmlns="http://example.test/not-bibliography"><Source>' +
            "<Title>Not Bibliography</Title></Source></Sources>",
        ),
      ),
    );
    expect(parsed.bibliography).toBeUndefined();
  });
});
