import { strToU8, unzipSync, zipSync } from "fflate";
import { describe, expect, it } from "vite-plus/test";

import { generateDocument } from "../generate";
import { parseDocumentSync } from "../parse";

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math";

function packageOf(bodyXml: string): Uint8Array {
  return zipSync({
    "[Content_Types].xml": strToU8(
      '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>',
    ),
    "_rels/.rels": strToU8(
      '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
    ),
    "word/document.xml": strToU8(
      `<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="${W_NS}" xmlns:m="${M_NS}"><w:body>${bodyXml}</w:body></w:document>`,
    ),
  });
}

const nary = (chr: string): string =>
  `<w:p><m:oMathPara><m:oMath><m:nary><m:naryPr><m:chr m:val="${chr}"/></m:naryPr><m:sub><m:r><m:t>i</m:t></m:r></m:sub><m:sup/><m:e><m:r><m:t>x</m:t></m:r></m:e></m:nary></m:oMath></m:oMathPara></w:p>`;

function mathOf(chr: string): Record<string, unknown> {
  const children = parseDocumentSync(packageOf(nary(chr))).sections?.[0]?.children ?? [];
  const para = (children[0] as { paragraph: { children: { math: { children: Record<string, unknown>[] } }[] } })
    .paragraph;
  return para.children[0]!.math.children[0]!;
}

describe("m:nary operator character", () => {
  it("keeps ∑ as `sum` and ∫ as `integral` without an override", () => {
    expect(mathOf("∑")).toHaveProperty("sum");
    expect(mathOf("∫")).toHaveProperty("integral");
    expect((mathOf("∫").integral as { properties?: unknown }).properties).toBeUndefined();
  });

  it("carries any other operator (∏, ∮, ⋃) as an override", () => {
    for (const chr of ["∏", "∮", "⋃"]) {
      const node = mathOf(chr) as { integral: { properties: { character: string } } };
      expect(node.integral.properties.character).toBe(chr);
    }
  });

  it("writes the override back as m:chr", async () => {
    const out = await generateDocument(parseDocumentSync(packageOf(nary("∏"))), { type: "uint8array" });
    const xml = new TextDecoder().decode(unzipSync(out)["word/document.xml"]);
    expect(xml).toContain('<m:chr m:val="∏"/>');
  });
});
