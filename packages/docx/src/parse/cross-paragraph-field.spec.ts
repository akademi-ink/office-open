import { strToU8, unzipSync, zipSync } from "fflate";
import { describe, expect, it } from "vite-plus/test";

import { generateDocument } from "../generate";
import { parseDocumentSync } from "../parse";

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

/** A minimal package whose body is `bodyXml`. */
function packageOf(bodyXml: string): Uint8Array {
  return zipSync({
    "[Content_Types].xml": strToU8(
      '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>',
    ),
    "_rels/.rels": strToU8(
      '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
    ),
    "word/document.xml": strToU8(
      `<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="${W_NS}"><w:body>${bodyXml}</w:body></w:document>`,
    ),
  });
}

// Mendeley's bibliography: one ADDIN field opens in the first entry and closes
// in the last, so every entry paragraph is crossed by a field it cannot hold.
const BIBLIOGRAPHY =
  '<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
  '<w:r><w:instrText xml:space="preserve">ADDIN Mendeley Bibliography CSL_BIBLIOGRAPHY </w:instrText></w:r>' +
  '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
  "<w:r><w:t>Ahmad, A. (2020). Judul pertama.</w:t></w:r></w:p>" +
  "<w:p><w:r><w:t>Budi, B. (2021). Judul kedua.</w:t></w:r></w:p>" +
  "<w:p><w:r><w:t>Citra, C. (2022). Judul ketiga.</w:t></w:r>" +
  '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>';

function documentXml(output: Uint8Array): string {
  return new TextDecoder().decode(unzipSync(output)["word/document.xml"]);
}

const RESULTLESS_FIELD =
  '<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
  '<w:r><w:instrText xml:space="preserve">PAGE</w:instrText></w:r>' +
  '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
  '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>';

describe("a result-less complex field", () => {
  it("keeps its separate marker through a generate fixed point", async () => {
    const parsed = parseDocumentSync(packageOf(RESULTLESS_FIELD));
    const children = parsed.sections?.[0]?.children ?? [];
    expect(children[0]).toMatchObject({
      paragraph: { children: [{ complexField: { instruction: "PAGE", hasSeparate: true } }] },
    });

    const output = await generateDocument(parsed, { type: "uint8array" });
    const xml = documentXml(output);
    expect(xml.match(/w:fldCharType="begin"/g)).toHaveLength(1);
    expect(xml.match(/w:fldCharType="separate"/g)).toHaveLength(1);
    expect(xml.match(/w:fldCharType="end"/g)).toHaveLength(1);
    expect(parseDocumentSync(output).sections?.[0]?.children).toEqual(parsed.sections?.[0]?.children);
  });
});

describe("paragraphs crossed by a multi-paragraph complex field", () => {
  it("parse as editable paragraphs whose marker runs stay verbatim", () => {
    const children = parseDocumentSync(packageOf(BIBLIOGRAPHY)).sections?.[0]?.children ?? [];
    expect(children.map((child) => Object.keys(child)[0])).toEqual([
      "paragraph",
      "paragraph",
      "paragraph",
    ]);
    const first = (children[0] as { paragraph: { children?: unknown[] } }).paragraph.children ?? [];
    expect(first.slice(0, 3).every((run) => typeof (run as { rawXml?: unknown }).rawXml === "string")).toBe(true);
    expect(first[3]).toMatchObject({ text: "Ahmad, A. (2020). Judul pertama." });
    const last = (children[2] as { paragraph: { children?: unknown[] } }).paragraph.children ?? [];
    expect(last[0]).toMatchObject({ text: "Citra, C. (2022). Judul ketiga." });
    expect((last[1] as { rawXml: string }).rawXml).toContain('w:fldCharType="end"');
  });

  it("round-trips the field once and reaches a parse fixed point", async () => {
    const parsed = parseDocumentSync(packageOf(BIBLIOGRAPHY));
    const output = await generateDocument(parsed, { type: "uint8array" });
    const xml = documentXml(output);
    expect(xml.match(/w:fldCharType="begin"/g)).toHaveLength(1);
    expect(xml.match(/w:fldCharType="separate"/g)).toHaveLength(1);
    expect(xml.match(/w:fldCharType="end"/g)).toHaveLength(1);
    expect(xml).toContain("ADDIN Mendeley Bibliography CSL_BIBLIOGRAPHY");
    expect(parseDocumentSync(output).sections?.[0]?.children).toEqual(parsed.sections?.[0]?.children);
  });
});
