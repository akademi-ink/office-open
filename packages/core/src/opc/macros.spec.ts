import { describe, expect, it } from "vitest";

import {
  isMacroEnabledContentType,
  isMacroEnabledPackage,
  isMacroPartPath,
  mainDocumentContentType,
  withoutMacroParts,
} from "./macros";

const DOCM_MAIN = "application/vnd.ms-word.document.macroEnabled.main+xml";
const DOTM_MAIN = "application/vnd.ms-word.template.macroEnabled.main+xml";
const DOCX_MAIN =
  "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";

function table(main: string) {
  return {
    defaults: [{ extension: "bin", contentType: "application/vnd.ms-office.vbaProject" }],
    overrides: [{ partName: "/word/document.xml", contentType: main }],
  };
}

describe("macro-enabled content types", () => {
  it("recognizes docm and dotm main types", () => {
    expect(isMacroEnabledContentType(DOCM_MAIN)).toBe(true);
    expect(isMacroEnabledContentType(DOTM_MAIN)).toBe(true);
    expect(isMacroEnabledContentType(DOCX_MAIN)).toBe(false);
    expect(isMacroEnabledContentType(undefined)).toBe(false);
  });

  it("reads the main document content type case-insensitively", () => {
    expect(
      mainDocumentContentType({
        defaults: [],
        overrides: [{ partName: "/Word/Document.xml", contentType: DOCM_MAIN }],
      }),
    ).toBe(DOCM_MAIN);
    expect(mainDocumentContentType(undefined)).toBeUndefined();
  });

  it("classifies a parsed package from its content-type table", () => {
    expect(isMacroEnabledPackage(table(DOCM_MAIN))).toBe(true);
    expect(isMacroEnabledPackage(table(DOTM_MAIN))).toBe(true);
    expect(isMacroEnabledPackage(table(DOCX_MAIN))).toBe(false);
    expect(isMacroEnabledPackage(undefined)).toBe(false);
  });
});

describe("macro part paths", () => {
  it("matches VBA project parts at any package depth", () => {
    expect(isMacroPartPath("word/vbaProject.bin")).toBe(true);
    expect(isMacroPartPath("word/vbaData.xml")).toBe(true);
    expect(isMacroPartPath("word/vbaProjectSignature.bin")).toBe(true);
    expect(isMacroPartPath("ppt/vbaProject.bin")).toBe(true);
    expect(isMacroPartPath("word/document.xml")).toBe(false);
    expect(isMacroPartPath("word/media/vbaProject.bin.png")).toBe(false);
  });

  it("drops VBA parts without touching other passthrough parts", () => {
    const parts = [
      { path: "word/vbaProject.bin", data: new Uint8Array([1]) },
      { path: "customXml/item1.xml", data: new Uint8Array([2]) },
    ];
    expect(withoutMacroParts(parts).map((part) => part.path)).toEqual(["customXml/item1.xml"]);
  });
});
