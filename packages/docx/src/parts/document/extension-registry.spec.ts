import { unzipSync, zipSync } from "fflate";
import { describe, expect, it } from "vite-plus/test";

import { compileDocument } from "../../compiler";
import { COMMENTS_NS } from "../comments";
import {
  WORD_EXTENSIONS,
  extensionIgnorablePrefixes,
  isWordExtensionNamespace,
} from "./extension-registry";

describe("Word extension registry", () => {
  it("covers Word 2010 through 2024, oldest first", () => {
    expect(WORD_EXTENSIONS.map((extension) => extension.prefix)).toEqual([
      "w14",
      "wp14",
      "w15",
      "w16se",
      "w16cid",
      "w16",
      "w16cex",
      "w16sdtdh",
      "w16du",
      "w16sdtfl",
    ]);
    // Every registered namespace is a Word extension namespace by URI.
    for (const extension of WORD_EXTENSIONS) {
      expect(extension.feature.length).toBeGreaterThan(0);
      expect(extension.ignorable).toBe(true);
    }
  });

  it("derives mc:Ignorable from the prefixes a root declares, in order", () => {
    expect(extensionIgnorablePrefixes(["mc", "r", "w", "w14", "w15"])).toBe("w14 w15");
    // Prefixes that are not registered extensions never enter the list.
    expect(extensionIgnorablePrefixes(["w", "mc", "r", "v"])).toBe("");
    expect(extensionIgnorablePrefixes(["w16cid", "w14"])).toBe("w16cid w14");
  });

  it("recognizes a Word extension namespace newer than the registry", () => {
    expect(isWordExtensionNamespace("http://schemas.microsoft.com/office/word/2026/wordml")).toBe(
      true,
    );
    expect(isWordExtensionNamespace("http://schemas.microsoft.com/office/word/2010/wordml")).toBe(
      true,
    );
    expect(
      isWordExtensionNamespace("http://schemas.openxmlformats.org/wordprocessingml/2006/main"),
    ).toBe(false);
  });

  it("declares w16cid/w16cex as ignorable on the comment-infrastructure roots", () => {
    const ignorable = /mc:Ignorable="([^"]+)"/.exec(COMMENTS_NS)?.[1]?.split(" ") ?? [];
    for (const prefix of ["w14", "w15", "w16cid", "w16cex", "w16se"]) {
      expect(ignorable).toContain(prefix);
      expect(COMMENTS_NS).toContain(`xmlns:${prefix}=`);
    }
  });

  it("emits document.xml with every ignorable prefix declared on the root", () => {
    const files = compileDocument({
      sections: [{ children: [{ paragraph: { children: ["x"] } }] }],
    });
    const xml = new TextDecoder().decode(files["word/document.xml"] as Uint8Array);
    const ignorable = /mc:Ignorable="([^"]+)"/.exec(xml)?.[1]?.split(" ") ?? [];
    expect(ignorable.length).toBeGreaterThan(0);
    for (const prefix of ignorable) {
      expect(xml).toContain(`xmlns:${prefix}=`);
    }
    // w16se can be emitted by w16se:sym, so it must stay ignorable.
    expect(ignorable).toContain("w16se");
  });

  it("declares all new extension namespaces on the emitted w16 comment parts", () => {
    const files = unzipSync(
      zipSync(
        compileDocument({
          sections: [{ children: [{ paragraph: { children: ["x"] } }] }],
          comments: [{ id: 0, author: "A", children: [{ children: ["c"] }] }],
          commentsIds: [{ paraId: "1A2B3C4D", durableId: "42" }],
          commentsExtensible: [{ durableId: "42", dateUtc: "2024-01-01T00:00:00Z" }],
        }),
      ),
    );
    for (const part of ["word/commentsIds.xml", "word/commentsExtensible.xml"]) {
      const xml = new TextDecoder().decode(files[part] as Uint8Array);
      const ignorable = /mc:Ignorable="([^"]+)"/.exec(xml)?.[1]?.split(" ") ?? [];
      expect(ignorable).toContain("w16cid");
      expect(ignorable).toContain("w16cex");
      for (const prefix of ignorable) {
        expect(xml).toContain(`xmlns:${prefix}=`);
      }
    }
  });
});
