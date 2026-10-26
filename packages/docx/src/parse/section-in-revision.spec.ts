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

const sectPr = (start: number): string =>
  `<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgNumType w:start="${start}"/></w:sectPr>`;

const REVISION =
  '<w:pPrChange w:id="1" w:author="A" w:date="2024-01-01T00:00:00Z">' +
  '<w:pPr><w:pStyle w:val="Heading3"/></w:pPr>' +
  "</w:pPrChange>";

const TEXT = (value: string): string => `<w:r><w:t>${value}</w:t></w:r>`;

describe("a paragraph's section break and its formatting revision", () => {
  it("keeps sectPr as a direct pPr child before pPrChange and reaches a fixed point", async () => {
    const sourceBody =
      `<w:p><w:pPr>${sectPr(1)}${REVISION}</w:pPr>${TEXT("First")}</w:p>` +
      `<w:p>${TEXT("Second")}</w:p>` +
      sectPr(2);
    const source = documentXml(packageOf(sourceBody));
    const once = await generateDocument(parseDocumentSync(packageOf(sourceBody)), {
      type: "uint8array",
    });
    const onceXml = documentXml(once);

    expect((onceXml.match(/<w:sectPr\b/g) ?? []).length).toBe(
      (source.match(/<w:sectPr\b/g) ?? []).length,
    );
    expect(onceXml).toMatch(/<w:sectPr\b[\s\S]*?<\/w:sectPr><w:pPrChange\b/);
    const revisionXml = onceXml.match(/<w:pPrChange\b[\s\S]*?<\/w:pPrChange>/)?.[0] ?? "";
    expect(revisionXml).not.toContain("<w:sectPr");
    expect(onceXml).toContain('<w:pStyle w:val="Heading3"/>');

    const twice = await generateDocument(parseDocumentSync(once), { type: "uint8array" });
    expect(documentXml(twice)).toBe(onceXml);
  });

  it("preserves a sectPr that exists inside pPrChange and does not make it a section", async () => {
    const sourceBody =
      `<w:p><w:pPr><w:pPrChange w:id="1" w:author="A" w:date="2024-01-01T00:00:00Z"><w:pPr>` +
      `<w:pStyle w:val="Heading3"/>${sectPr(7)}</w:pPr></w:pPrChange></w:pPr>${TEXT("First")}</w:p>` +
      `<w:p>${TEXT("Second")}</w:p>` +
      sectPr(2);
    const parsed = parseDocumentSync(packageOf(sourceBody));
    expect(parsed.sections).toHaveLength(1);

    const onceBytes = await generateDocument(parsed, { type: "uint8array" });
    const once = documentXml(onceBytes);
    expect(once).toMatch(/<w:pPrChange\b[\s\S]*?<w:pStyle w:val="Heading3"\/>[\s\S]*?<w:sectPr\b/);
    expect((once.match(/<w:sectPr\b/g) ?? []).length).toBe(2);
    const twice = await generateDocument(parseDocumentSync(onceBytes), {
      type: "uint8array",
    });
    expect(documentXml(twice)).toBe(once);
  });
});
