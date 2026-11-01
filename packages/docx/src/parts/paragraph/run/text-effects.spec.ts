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

  it("keeps unmodeled w14 children (w14:textOutline) verbatim", () => {
    const run = firstRun(
      packageWithRunProperties(
        '<w14:ligatures w14:val="standard"/><w14:textOutline w14:w="6350"><w14:solidFill><w14:srgbClr w14:val="FF0000"/></w14:solidFill></w14:textOutline>',
      ),
    );
    expect(run.textEffects).toEqual({ ligatures: "standard" });
    expect(run.w14RawXml).toContain("w14:textOutline");
    expect(run.w14RawXml).toContain('w14:val="FF0000"');
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
