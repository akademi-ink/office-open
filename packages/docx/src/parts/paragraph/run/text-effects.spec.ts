import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { describe, expect, it } from "vite-plus/test";

import { generateDocument } from "../../../generate";
import { parseDocumentSync } from "../../../parse";
import { stringifyTextEffects, type TextEffectsOptions } from "./text-effects";

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const W14_NS = "http://schemas.microsoft.com/office/word/2010/wordml";

interface RunChild {
  textEffects?: TextEffectsOptions;
  w14RawXml?: string;
}

/** Package with one body paragraph whose run properties are `rPrInner`. */
function packageWithRunProperties(rPrInner: string): Uint8Array {
  return zipSync({
    "[Content_Types].xml": strToU8(
      '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>',
    ),
    "_rels/.rels": strToU8(
      '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
    ),
    "word/document.xml": strToU8(
      `<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="${W_NS}" xmlns:w14="${W14_NS}"><w:body><w:p><w:r><w:rPr>${rPrInner}</w:rPr><w:t>styled</w:t></w:r></w:p></w:body></w:document>`,
    ),
  });
}

function firstRun(source: Uint8Array): RunChild {
  const parsed = parseDocumentSync(source);
  const paragraph = parsed.sections?.[0]?.children?.[0] as {
    paragraph: { children?: RunChild[] };
  };
  return paragraph.paragraph.children?.[0] as RunChild;
}

// The w14 block Word writes for a run with OpenType features + a glowing,
// shadowed gradient-free text fill and a 3-D bevel.
const WORD_RPR =
  '<w14:ligatures w14:val="standardContextual"/>' +
  '<w14:numForm w14:val="oldStyle"/>' +
  '<w14:numSpacing w14:val="tabular"/>' +
  '<w14:stylisticSets><w14:styleSet w14:id="20" w14:val="1"/><w14:styleSet w14:id="21" w14:val="0"/></w14:stylisticSets>' +
  '<w14:cntxtAlts w14:val="0"/>' +
  '<w14:glow w14:rad="63500"><w14:schemeClr w14:val="accent1"><w14:lumMod w14:val="60000"/><w14:lumOff w14:val="40000"/></w14:schemeClr></w14:glow>' +
  '<w14:shadow w14:blurRad="50800" w14:dist="25400" w14:dir="2700000" w14:algn="tl"><w14:srgbClr w14:val="000000"><w14:alpha w14:val="40000"/></w14:srgbClr></w14:shadow>' +
  '<w14:reflection w14:blurRad="6350" w14:stA="52000" w14:endA="300" w14:dir="5400000" w14:fadeDir="5400000" w14:dist="0" w14:algn="b"/>' +
  '<w14:textFill><w14:solidFill><w14:schemeClr w14:val="accent2"><w14:tint w14:val="60000"/></w14:schemeClr></w14:solidFill></w14:textFill>' +
  '<w14:scene3d><w14:camera w14:prst="orthographicFront"/><w14:lightRig w14:rig="threePt" w14:dir="t"/></w14:scene3d>' +
  '<w14:props3d w14:extrusionH="63500" w14:contourW="12700" w14:prstMaterial="warmMatte">' +
  '<w14:bevelT w14:w="63500" w14:h="25400" w14:prst="circle"/>' +
  '<w14:extrusionClr><w14:srgbClr w14:val="156082"/></w14:extrusionClr>' +
  "</w14:props3d>";

describe("w14 text effects", () => {
  it("parses the elements Word writes by default into the typed model", () => {
    const run = firstRun(packageWithRunProperties(WORD_RPR));
    const effects = run.textEffects!;

    expect(effects.ligatures).toBe("standardContextual");
    expect(effects.numForm).toBe("oldStyle");
    expect(effects.numSpacing).toBe("tabular");
    expect(effects.stylisticSets).toEqual([
      { id: 20, enabled: true },
      { id: 21, enabled: false },
    ]);
    expect(effects.contextualAlternatives).toBe(false);

    expect(effects.glow).toEqual({
      radius: 63500,
      color: { schemeColor: "accent1", transforms: { lumMod: 60, lumOff: 40 } },
    });
    expect(effects.shadow).toEqual({
      blurRadius: 50800,
      distance: 25400,
      direction: 45,
      alignment: "topLeft",
      color: { value: "000000", transforms: { alpha: 40 } },
    });
    expect(effects.reflection).toEqual({
      blurRadius: 6350,
      startAlpha: 52,
      endAlpha: 0.3,
      direction: 90,
      fadeDirection: 90,
      distance: 0,
      alignment: "bottom",
    });
    expect(effects.textFill).toEqual({
      solid: { schemeColor: "accent2", transforms: { tint: 60 } },
    });
    expect(effects.scene3d).toEqual({
      camera: "orthographicFront",
      lightRig: "threePt",
      lightDirection: "top",
    });
    expect(effects.props3d).toEqual({
      bevelT: { w: 63500, h: 25400, prst: "circle" },
      extrusionH: 63500,
      contourW: 12700,
      prstMaterial: "warmMatte",
      extrusionColor: { value: "156082" },
    });
    // Nothing was dropped into the verbatim channel.
    expect(run.w14RawXml).toBeUndefined();
  });

  it("re-emits the typed effects and re-parses them identically", async () => {
    const parsed = parseDocumentSync(packageWithRunProperties(WORD_RPR));
    const files = unzipSync(await generateDocument(parsed, { type: "uint8array" }));
    const xml = strFromU8(files["word/document.xml"] as Uint8Array);

    expect(xml).toContain('<w14:ligatures w14:val="standardContextual"/>');
    expect(xml).toContain('<w14:glow w14:rad="63500">');
    expect(xml).toContain('<w14:schemeClr w14:val="accent1">');
    expect(xml).toContain('<w14:lumMod w14:val="60000"/>');
    expect(xml).toContain('<w14:shadow w14:blurRad="50800" w14:dist="25400" w14:dir="2700000"');
    expect(xml).toContain('w14:algn="tl"');
    expect(xml).toContain("<w14:textFill><w14:solidFill>");
    expect(xml).toContain("<w14:scene3d>");
    expect(xml).toContain('<w14:props3d w14:extrusionH="63500" w14:contourW="12700"');

    // gen2 → parse is stable against the source parse.
    const first = firstRun(packageWithRunProperties(WORD_RPR));
    const second = firstRun(
      zipSync(Object.fromEntries(Object.entries(files).map(([k, v]) => [k, v as Uint8Array]))),
    );
    expect(second.textEffects).toEqual(first.textEffects);
  });

  it("models w14:textOutline as a typed effect instead of verbatim XML", () => {
    const run = firstRun(
      packageWithRunProperties(
        '<w14:ligatures w14:val="standard"/><w14:textOutline w14:w="6350" w14:cap="rnd" w14:cmpd="dbl" w14:algn="ctr"><w14:solidFill><w14:srgbClr w14:val="FF0000"/></w14:solidFill></w14:textOutline>',
      ),
    );
    expect(run.textEffects).toEqual({
      ligatures: "standard",
      textOutline: {
        width: 6350,
        cap: "round",
        compound: "double",
        alignment: "center",
        fill: { solid: { value: "FF0000" } },
      },
    });
    expect(run.w14RawXml).toBeUndefined();
  });

  it("parses a gradient outline (gsLst/lin, prstDash, miter) and round-trips it", async () => {
    const outlineXml =
      '<w14:textOutline w14:w="12700">' +
      "<w14:gradFill><w14:gsLst>" +
      '<w14:gs w14:pos="0"><w14:schemeClr w14:val="accent1"/></w14:gs>' +
      '<w14:gs w14:pos="100000"><w14:srgbClr w14:val="FFFFFF"><w14:alpha w14:val="50000"/></w14:srgbClr></w14:gs>' +
      "</w14:gsLst>" +
      '<w14:lin w14:ang="5400000" w14:scaled="1"/>' +
      "</w14:gradFill>" +
      '<w14:prstDash w14:val="sysDash"/>' +
      '<w14:miter w14:lim="400000"/>' +
      "</w14:textOutline>";

    const run = firstRun(packageWithRunProperties(outlineXml));
    expect(run.textEffects?.textOutline).toMatchObject({
      width: 12700,
      dash: "sysDash",
      join: { miterLimit: 400 },
      fill: {
        gradient: {
          stops: [
            { position: 0, color: { schemeColor: "accent1" } },
            {
              position: 100,
              color: { value: "FFFFFF", transforms: { alpha: 50 } },
            },
          ],
          linear: { angle: 90, scaled: true },
        },
      },
    });

    // gen1 → gen2 keeps the same typed model.
    const parsed = parseDocumentSync(packageWithRunProperties(outlineXml));
    const files = unzipSync(await generateDocument(parsed, { type: "uint8array" }));
    const xml = strFromU8(files["word/document.xml"] as Uint8Array);
    expect(xml).toContain("<w14:gsLst>");
    expect(xml).toContain('<w14:gs w14:pos="100000">');
    expect(xml).toContain('<w14:lin w14:ang="5400000" w14:scaled="1"/>');
    expect(xml).toContain('<w14:prstDash w14:val="sysDash"/>');
    expect(xml).toContain('<w14:miter w14:lim="400000"/>');
    const second = firstRun(
      zipSync(Object.fromEntries(Object.entries(files).map(([k, v]) => [k, v as Uint8Array]))),
    );
    expect(second.textEffects).toEqual(run.textEffects);
  });

  it("parses a path-shaded outline (w14:path + w14:fillToRect)", () => {
    const run = firstRun(
      packageWithRunProperties(
        '<w14:textOutline w14:w="6350">' +
          "<w14:gradFill><w14:gsLst>" +
          '<w14:gs w14:pos="0"><w14:srgbClr w14:val="000000"/></w14:gs>' +
          '<w14:gs w14:pos="100000"><w14:srgbClr w14:val="FFFFFF"/></w14:gs>' +
          "</w14:gsLst>" +
          '<w14:path w14:path="circle"><w14:fillToRect w14:l="50000" w14:t="50000" w14:r="50000" w14:b="50000"/></w14:path>' +
          "</w14:gradFill>" +
          "<w14:round/>" +
          "</w14:textOutline>",
      ),
    );
    expect(run.textEffects?.textOutline).toMatchObject({
      join: "round",
      fill: {
        gradient: {
          path: {
            path: "circle",
            fillToRect: { left: 50, top: 50, right: 50, bottom: 50 },
          },
        },
      },
    });
  });

  it("wraps a typed gradient text fill in w14:textFill and round-trips it", () => {
    const xml = stringifyTextEffects({
      textFill: {
        gradient: {
          stops: [
            { position: 0, color: { value: "FF0000" } },
            { position: 100, color: { value: "0000FF" } },
          ],
          linear: { angle: 0 },
        },
      },
    });
    expect(xml).toBe(
      "<w14:textFill><w14:gradFill><w14:gsLst>" +
        '<w14:gs w14:pos="0"><w14:srgbClr w14:val="FF0000"/></w14:gs>' +
        '<w14:gs w14:pos="100000"><w14:srgbClr w14:val="0000FF"/></w14:gs>' +
        "</w14:gsLst>" +
        '<w14:lin w14:ang="0"/>' +
        "</w14:gradFill></w14:textFill>",
    );
  });

  it("keeps a degenerate gradient verbatim inside its w14:textFill wrapper", () => {
    const run = firstRun(
      packageWithRunProperties(
        '<w14:textFill><w14:gradFill><w14:unknown w14:val="1"/></w14:gradFill></w14:textFill>',
      ),
    );
    expect(run.textEffects?.textFill?.gradientXml).toContain("w14:gradFill");
    expect(run.textEffects?.textFill?.gradientXml).toContain("w14:unknown");
    // The fallback re-emits wrapped in w14:textFill, not as a bare gradFill.
    const xml = stringifyTextEffects(run.textEffects!);
    expect(xml).toContain(
      '<w14:textFill><w14:gradFill><w14:unknown w14:val="1"/></w14:gradFill></w14:textFill>',
    );
  });

  it("emits typed children in schema order", () => {
    const xml = stringifyTextEffects({
      ligatures: "standard",
      numForm: "lining",
      numSpacing: "proportional",
      stylisticSets: [{ id: 7 }],
      contextualAlternatives: true,
      glow: { radius: 12700, color: { value: "FF0000" } },
      shadow: { color: { value: "000000" } },
      reflection: { startAlpha: 50 },
      textFill: { none: true },
      textOutline: { width: 6350 },
      scene3d: { camera: "orthographicFront", lightRig: "threePt", lightDirection: "top" },
      props3d: { extrusionH: 0 },
    });
    const order = [
      "w14:ligatures",
      "w14:numForm",
      "w14:numSpacing",
      "w14:stylisticSets",
      "w14:cntxtAlts",
      "w14:glow",
      "w14:shadow",
      "w14:reflection",
      "w14:textOutline",
      "w14:textFill",
      "w14:scene3d",
      "w14:props3d",
    ];
    let previous = -1;
    for (const tag of order) {
      const index = xml.indexOf(`<${tag}`);
      expect(index).toBeGreaterThan(previous);
      previous = index;
    }
  });
});
