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

const count = (xml: string): number => (xml.match(/<w:sectPr\b/g) ?? []).length;

async function roundTrip(bodyXml: string): Promise<string> {
  const out = await generateDocument(parseDocumentSync(packageOf(bodyXml)), {
    type: "uint8array",
  });
  return new TextDecoder().decode(unzipSync(out)["word/document.xml"]);
}

const sectPr = (start: number): string =>
  `<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgNumType w:start="${start}"/></w:sectPr>`;

describe("a section break inside block content the section model does not split", () => {
  it("keeps sectPr carried by paragraphs inside w:sdtContent", async () => {
    const body =
      `<w:p><w:r><w:t>Before</w:t></w:r></w:p>` +
      `<w:sdt><w:sdtPr><w:docPartObj><w:docPartGallery w:val="Bibliographies"/></w:docPartObj></w:sdtPr>` +
      `<w:sdtContent><w:p><w:pPr>${sectPr(3)}</w:pPr><w:r><w:t>Ref</w:t></w:r></w:p>` +
      `<w:p><w:pPr><w:jc w:val="center"/>${sectPr(4)}</w:pPr><w:r><w:t>More</w:t></w:r></w:p></w:sdtContent></w:sdt>` +
      `<w:p><w:r><w:t>After</w:t></w:r></w:p>` +
      sectPr(1);
    const once = await roundTrip(body);
    expect(count(once)).toBe(3);
    expect(once).toContain('<w:pgNumType w:start="3"/>');
    expect(once).toContain('<w:pgNumType w:start="4"/>');
    // A second pass is a fixed point.
    const twice = await roundTrip(once.replace(/^.*<w:body>|<\/w:body>.*$/g, ""));
    expect(count(twice)).toBe(3);
  });

  it("keeps sectPr carried by a paragraph inside a table cell", async () => {
    const body =
      `<w:tbl><w:tr><w:tc><w:p><w:pPr>${sectPr(7)}</w:pPr></w:p></w:tc></w:tr></w:tbl>` +
      `<w:p><w:r><w:t>x</w:t></w:r></w:p>` +
      sectPr(1);
    expect(count(await roundTrip(body))).toBe(2);
  });
});
