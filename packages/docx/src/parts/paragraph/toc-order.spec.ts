import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import { parseParagraph } from "../../body";
import type { DocxReadContext } from "../../context";

const W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const readCtx = {} as unknown as DocxReadContext;
const fieldHead =
  '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
  '<w:r><w:instrText xml:space="preserve"> TOC \\h </w:instrText></w:r>' +
  '<w:r><w:fldChar w:fldCharType="separate"/></w:r>';
const link = (text: string) =>
  `<w:hyperlink w:anchor="heading"><w:r><w:t>${text}</w:t></w:r></w:hyperlink>`;
const tab = "<w:r><w:tab/></w:r>";

function kindsWith(head: string, inner: string): string[] {
  const doc = parseXml(`<w:p ${W_NS}>${head}${inner}</w:p>`);
  const opts = parseParagraph(doc.elements![0]!, readCtx) as {
    children?: Array<Record<string, unknown>>;
  };
  return (opts.children ?? []).map((c) => {
    const key = Object.keys(c)[0] ?? "?";
    const children = (c as { children?: Array<Record<string, unknown>> }).children;
    return key === "children" && children?.[0] && "tab" in children[0] ? "tab" : key;
  });
}

function kinds(inner: string): string[] {
  return kindsWith(fieldHead, inner);
}

const pageFieldHead =
  '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
  '<w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>' +
  '<w:r><w:fldChar w:fldCharType="separate"/></w:r>';

describe("TOC entry order inside the field result", () => {
  it("keeps the tab before the page link when the title ends in a roman numeral", () => {
    expect(kinds(`${link("Chapter ")}${link("IV")}${tab}${link("12")}`)).toEqual([
      "hyperlink",
      "hyperlink",
      "tab",
      "hyperlink",
    ]);
  });
  it("keeps the tab before the page link when the title ends in a number", () => {
    expect(kinds(`${link("Bab 2")}${tab}${link("7")}`)).toEqual(["hyperlink", "tab", "hyperlink"]);
  });
  it("keeps a tab that precedes an unlinked page number", () => {
    expect(kinds(`${link("Chapter")}${tab}<w:r><w:t>12</w:t></w:r>`)).toEqual([
      "hyperlink",
      "tab",
      "text",
    ]);
  });
  it("does not lift a tab out of a PAGE field result", () => {
    expect(
      kindsWith(pageFieldHead, `<w:r><w:t>1</w:t></w:r>${tab}<w:r><w:t>2</w:t></w:r>`),
    ).toEqual(["text", "tab", "text"]);
  });
  it("does not lift a tab out of a REF field result", () => {
    expect(
      kindsWith(
        pageFieldHead.replace(" PAGE ", " REF _Ref1 \\h "),
        `<w:r><w:t>1</w:t></w:r>${tab}<w:r><w:t>2</w:t></w:r>`,
      ),
    ).toEqual(["text", "tab", "text"]);
  });
  it("keeps the tab run's styled properties in a TOC entry", () => {
    const doc = parseXml(
      `<w:p ${W_NS}>${fieldHead}${link("Chapter")}<w:r><w:rPr><w:rStyle w:val="TOC1"/><w:b/></w:rPr><w:tab/></w:r>${link("12")}</w:p>`,
    );
    const opts = parseParagraph(doc.elements![0]!, readCtx) as {
      children?: Array<Record<string, unknown>>;
    };
    const tab = (opts.children ?? []).find((c) => "children" in c) as
      | { children?: Array<Record<string, unknown>>; bold?: boolean; style?: string }
      | undefined;
    expect(tab?.children?.[0]).toEqual({ tab: true });
    expect(tab?.bold === true || tab?.style !== undefined).toBe(true);
  });
  it("keeps multiple runs inside a hyperlink in order", () => {
    expect(
      kinds(
        `<w:hyperlink w:anchor="h"><w:r><w:t>Multi</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>run</w:t></w:r></w:hyperlink>${tab}${link("9")}`,
      ),
    ).toEqual(["hyperlink", "tab", "hyperlink"]);
  });
});
