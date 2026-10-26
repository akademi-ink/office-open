import { strToU8, unzipSync, zipSync } from "fflate";
import { describe, expect, it } from "vite-plus/test";

import { generateDocument } from "../generate";
import { parseDocumentSync } from "../parse";

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

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

function documentXml(output: Uint8Array): string {
  return new TextDecoder().decode(unzipSync(output)["word/document.xml"]);
}

const begin = '<w:r><w:fldChar w:fldCharType="begin"/></w:r>';
const separate = '<w:r><w:fldChar w:fldCharType="separate"/></w:r>';
const end = '<w:r><w:fldChar w:fldCharType="end"/></w:r>';
const instr = (text: string): string =>
  `<w:r><w:instrText xml:space="preserve">${text}</w:instrText></w:r>`;
const text = (value: string): string => `<w:r><w:t>${value}</w:t></w:r>`;
const entry = (value: string): string =>
  `<w:p><w:pPr><w:pStyle w:val="TOC1"/></w:pPr>${text(value)}</w:p>`;
const tocOpen = `<w:p>${begin}${instr(' TOC \\o "1-4" ')}${separate}${text("A")}</w:p>`;

function markers(xml: string): string[] {
  return [...xml.matchAll(/w:fldCharType="(\w+)"/g)].map((match) => match[1]!);
}

function decodeXml(value: string): string {
  return value
    .replaceAll("&quot;", '"')
    .replaceAll("&amp;", "&")
    .replaceAll("&lt;", "<")
    .replaceAll("&gt;", ">")
    .replaceAll(/&#(\d+);/g, (_, code: string) => String.fromCodePoint(Number(code)));
}

function instructions(xml: string): string[] {
  return [...xml.matchAll(/<w:instrText[^>]*>([^<]*)</g)].map((match) => decodeXml(match[1]!.trim()));
}

async function roundTrip(bodyXml: string): Promise<Uint8Array> {
  return generateDocument(parseDocumentSync(packageOf(bodyXml)), { type: "uint8array" });
}

describe("TOC closing paragraph order", () => {
  it("keeps the outer end before a following empty TOC field and reaches a fixed point", async () => {
    const sourceBody =
      tocOpen +
      entry("B") +
      `<w:p>${end}${begin}${instr(' TOC \\o "1-4" ')}${end}</w:p>`;
    const source = documentXml(packageOf(sourceBody));
    const once = await roundTrip(sourceBody);
    const onceXml = documentXml(once);

    expect(markers(onceXml)).toEqual(markers(source));
    expect(instructions(onceXml)).toEqual(instructions(source));

    const twice = await generateDocument(parseDocumentSync(once), { type: "uint8array" });
    expect(documentXml(twice)).toBe(onceXml);
  });

  it("keeps the order when the closing paragraph also ends a section", async () => {
    const sourceBody =
      tocOpen +
      entry("B") +
      `<w:p><w:pPr><w:pStyle w:val="TOC1"/><w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgNumType w:fmt="lowerRoman"/></w:sectPr></w:pPr>` +
      `${end}${begin}${instr(' TOC \\o "1-4" ')}${end}</w:p>` +
      '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgNumType w:start="1"/></w:sectPr>';
    const source = documentXml(packageOf(sourceBody));
    const onceBytes = await roundTrip(sourceBody);
    const once = documentXml(onceBytes);

    expect(markers(once)).toEqual(markers(source));
    expect(instructions(once)).toEqual(instructions(source));
    const twice = await generateDocument(parseDocumentSync(onceBytes), { type: "uint8array" });
    expect(documentXml(twice)).toBe(once);
  });

  it("keeps a closing paragraph that only carries the end as one end", async () => {
    const sourceBody = tocOpen + entry("B") + `<w:p>${end}</w:p>`;
    const source = documentXml(packageOf(sourceBody));
    const onceBytes = await roundTrip(sourceBody);
    const once = documentXml(onceBytes);

    expect(markers(once).filter((marker) => marker === "end")).toHaveLength(1);
    expect(markers(once)).toEqual(markers(source));
    const twice = await generateDocument(parseDocumentSync(onceBytes), { type: "uint8array" });
    expect(documentXml(twice)).toBe(once);
  });
});
