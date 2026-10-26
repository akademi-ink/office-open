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

const INS = (id: number, runs: string): string =>
  `<w:ins w:id="${id}" w:author="A" w:date="2024-07-10T13:31:00Z">${runs}</w:ins>`;
const begin = '<w:r><w:fldChar w:fldCharType="begin"/></w:r>';
const separate = '<w:r><w:fldChar w:fldCharType="separate"/></w:r>';
const end = '<w:r><w:fldChar w:fldCharType="end"/></w:r>';
const instr = (text: string): string =>
  `<w:r><w:instrText xml:space="preserve">${text}</w:instrText></w:r>`;
const text = (value: string): string => `<w:r><w:t>${value}</w:t></w:r>`;

// A TOC whose entries a reviewer inserted with track changes on: each entry's
// HYPERLINK and PAGEREF open inside one w:ins and close inside the next, with
// the PAGEREF separator between them, outside any wrapper.
function entry(id: number, bookmark: string, title: string, page: string): string {
  return (
    INS(
      id,
      begin +
        instr(` HYPERLINK \\l "${bookmark}" `) +
        separate +
        text(title) +
        begin +
        instr(` PAGEREF ${bookmark} \\h `),
    ) +
    separate +
    INS(id + 1, text(page) + end + end)
  );
}

const TOC =
  '<w:p><w:pPr><w:pStyle w:val="TOC1"/></w:pPr>' +
  begin +
  instr(' TOC \\o "1-3" \\h \\z \\u ') +
  separate +
  entry(1, "_Toc1", "PENDAHULUAN", "1") +
  "</w:p>" +
  '<w:p><w:pPr><w:pStyle w:val="TOC1"/></w:pPr>' +
  entry(3, "_Toc2", "METODE", "5") +
  "</w:p>" +
  `<w:p>${end}</w:p>`;

function documentXml(output: Uint8Array): string {
  return new TextDecoder().decode(unzipSync(output)["word/document.xml"]);
}

function instructions(xml: string): string[] {
  return [...xml.matchAll(/<w:instrText[^>]*>([^<]*)</g)].map((m) => m[1]!.trim());
}

function markers(xml: string): string[] {
  return [...xml.matchAll(/w:fldCharType="(\w+)"/g)].map((m) => m[1]!);
}

describe("fields split by tracked insertions", () => {
  it("keep every instruction and marker of a tracked TOC entry, in order", async () => {
    const source = documentXml(packageOf(TOC));
    const output = documentXml(
      await generateDocument(parseDocumentSync(packageOf(TOC)), { type: "uint8array" }),
    );
    expect(instructions(output)).toEqual(instructions(source));
    expect(markers(output)).toEqual(markers(source));
    expect(output).toContain("PENDAHULUAN");
    expect(output).toContain("METODE");
  });

  it("reach a parse fixed point", async () => {
    const once = await generateDocument(parseDocumentSync(packageOf(TOC)), { type: "uint8array" });
    const twice = await generateDocument(parseDocumentSync(once), { type: "uint8array" });
    expect(documentXml(twice)).toBe(documentXml(once));
  });

  it("leave a field closed inside one insertion collapsed", () => {
    const paragraph = `<w:p>${INS(1, begin + instr(" PAGE ") + separate + text("3") + end)}</w:p>`;
    const children = parseDocumentSync(packageOf(paragraph)).sections?.[0]?.children ?? [];
    const insertion = (
      children[0] as { paragraph: { children: { insertion?: { children: unknown[] } }[] } }
    ).paragraph.children[0]!.insertion!;
    expect(insertion.children).toEqual([
      expect.objectContaining({ complexField: expect.objectContaining({ instruction: " PAGE " }) }),
    ]);
  });
});
