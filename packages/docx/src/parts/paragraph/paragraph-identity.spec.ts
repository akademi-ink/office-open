import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import { parseParagraph } from "../../body";
import type { DocxReadContext } from "../../context";
import { paragraphIdentityAttrs } from "./stringify";

// Paragraph identity attributes (rsid family + w14:paraId/textId) never touch
// the read context, so an empty mock suffices.
const readCtx = {} as unknown as DocxReadContext;

const NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"';

describe("paragraph identity attributes parse", () => {
  it("round-trips rsid family + w14:paraId/textId (hex verbatim, leading zeros kept)", () => {
    const xml =
      `<w:p ${NS}` +
      ' w14:paraId="0A1B2C3D" w14:textId="0E0F0A0B"' +
      ' w:rsidR="00992297" w:rsidRDefault="00112233" w:rsidP="11223344"' +
      ' w:rsidRPr="AABBCCDD" w:rsidDel="DEADBEEF"/>';
    const el = parseXml(xml).elements?.[0];
    if (!el) throw new Error("parsed document has no root element");
    const opts = parseParagraph(el, readCtx);
    expect(opts.paraId).toBe("0A1B2C3D");
    expect(opts.textId).toBe("0E0F0A0B");
    expect(opts.additionRsid).toBe("00992297");
    expect(opts.defaultRunRsid).toBe("00112233");
    expect(opts.propertiesRsid).toBe("11223344");
    expect(opts.runPropertiesRsid).toBe("AABBCCDD");
    expect(opts.deletionRsid).toBe("DEADBEEF");
  });

  it("reads paragraph attributes alongside run children", () => {
    const xml =
      `<w:p ${NS} w14:paraId="1A2B3C4D" w:rsidR="00FF">` + "<w:r><w:t>hi</w:t></w:r>" + "</w:p>";
    const el = parseXml(xml).elements?.[0];
    if (!el) throw new Error("parsed document has no root element");
    const opts = parseParagraph(el, readCtx);
    expect(opts.paraId).toBe("1A2B3C4D");
    expect(opts.additionRsid).toBe("00FF");
    // Single text run is promoted to opts.text by the simple-text optimization.
    expect(opts.text).toBe("hi");
  });

  it("reads and re-emits w14:noSpellErr (Word 2010 spelling exemption)", () => {
    const xml = `<w:p ${NS} w14:paraId="0A1B2C3D" w14:noSpellErr="1"/>`;
    const el = parseXml(xml).elements?.[0];
    if (!el) throw new Error("parsed document has no root element");
    const opts = parseParagraph(el, readCtx);
    expect(opts.noSpellError).toBe(true);

    // The writer re-emits the attribute on the w:p opening tag.
    expect(paragraphIdentityAttrs(opts)).toContain(' w14:noSpellErr="1"');
    expect(paragraphIdentityAttrs({ noSpellError: false })).toContain(' w14:noSpellErr="0"');
    // Absent stays absent (default = spell-checked).
    expect(paragraphIdentityAttrs({})).not.toContain("noSpellErr");
  });
});
