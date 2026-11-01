import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { describe, expect, it } from "vite-plus/test";

import { generateDocument } from "../generate";
import { parseDocumentSync } from "../parse";

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const IMAGE_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
const INK_REL = "http://schemas.microsoft.com/office/2011/relationships/ink";

const PNG = new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10]);

interface RawXmlChild {
  rawXml: string;
  rawMedia?: Array<{ fileName: string; type: string; data: Uint8Array }>;
}

interface MediaPart {
  path: string;
  relationshipType: string;
  body: string;
}

/** Minimal package whose main document body is `bodyXml`, plus optional extra
 *  parts (each pinned to the document part with one relationship). */
function packageOf(bodyXml: string, mediaParts: MediaPart[] = []): Uint8Array {
  const rels = mediaParts
    .map(
      (m, i) => `<Relationship Id="rId${i + 5}" Type="${m.relationshipType}" Target="${m.path}"/>`,
    )
    .join("");
  const entries: Record<string, Uint8Array> = {
    "[Content_Types].xml": strToU8(
      '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>',
    ),
    "_rels/.rels": strToU8(
      '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
    ),
    "word/_rels/document.xml.rels": strToU8(
      `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels}</Relationships>`,
    ),
    "word/document.xml": strToU8(
      `<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="${W_NS}" xmlns:r="${R_NS}" xmlns:mc="${MC_NS}"><w:body>${bodyXml}</w:body></w:document>`,
    ),
  };
  for (const part of mediaParts) {
    if (part.path === "media/pic.png") entries["word/media/pic.png"] = PNG;
    else if (part.body) entries[`word/${part.path}`] = strToU8(part.body);
  }
  return zipSync(entries);
}

/** rId → Target for every entry in a .rels part. */
function relTargetsOf(relsXml: string): Map<string, string> {
  const targets = new Map<string, string>();
  for (const tag of relsXml.match(/<Relationship\b[^>]*\/>/g) ?? []) {
    const id = /\bId="([^"]+)"/.exec(tag)?.[1];
    const target = /\bTarget="([^"]+)"/.exec(tag)?.[1];
    if (id && target) targets.set(id, target);
  }
  return targets;
}

/**
 * Every `r:id` / `r:embed` / `r:link` in an emitted XML part must resolve to a
 * relationship of that part, and an internal target must exist in the package.
 * This is the regression assertion for the rawXml passthrough paths.
 */
function expectNoDanglingReferences(files: Record<string, Uint8Array>): void {
  const paths = new Set(Object.keys(files));
  for (const [path, bytes] of Object.entries(files)) {
    if (!path.endsWith(".xml") || path.includes("_rels/")) continue;
    const xml = strFromU8(bytes);
    const refs = [...xml.matchAll(/\br:(?:id|embed|link)="([^"]+)"/g)].map((m) => m[1]!);
    if (refs.length === 0) continue;
    const dir = path.slice(0, path.lastIndexOf("/"));
    const relsPath = `${dir}/_rels/${path.slice(path.lastIndexOf("/") + 1)}.rels`;
    const rels = files[relsPath];
    if (!rels) throw new Error(`${path} carries relationship refs but ${relsPath} is absent`);
    const targets = relTargetsOf(strFromU8(rels));
    for (const ref of refs) {
      const target = targets.get(ref);
      if (!target) throw new Error(`${path}: ${ref} dangles (no entry in ${relsPath})`);
      if (target.includes("://")) continue;
      const resolved = target.startsWith("/")
        ? target.slice(1)
        : `${dir}/${target}`.replace(/\/[^/]+\/\.\.\//g, "/");
      if (!paths.has(resolved)) throw new Error(`${path}: ${ref} → missing part ${resolved}`);
    }
  }
}

/** One paragraph holding `childXml` — the body of most fixtures below. */
function singleParagraph(childXml: string): string {
  return `<w:p>${childXml}</w:p>`;
}

function firstParagraphChildren(source: Uint8Array): unknown[] {
  const parsed = parseDocumentSync(source);
  const first = parsed.sections?.[0]?.children?.[0] as
    | { paragraph: { children?: unknown[] } }
    | undefined;
  return first?.paragraph.children ?? [];
}

describe("rawXml passthrough — no dangling relationship ids", () => {
  it("keeps an unrecognized drawing payload verbatim with placeholder media", () => {
    const drawing =
      '<w:r><w:drawing><wps:wsp><wps:cNvPr id="1" name="future"/>' +
      '<a:graphic><a:graphicData uri="http://example.com/future"><future:blip r:embed="rId5"/></a:graphicData></a:graphic>' +
      "</wps:wsp></w:drawing></w:r>";
    const source = packageOf(singleParagraph(drawing), [
      { path: "media/pic.png", relationshipType: IMAGE_REL, body: "" },
    ]);
    const children = firstParagraphChildren(source);
    const raw = children[0] as RawXmlChild;

    expect(raw.rawXml).toContain("<w:drawing>");
    expect(raw.rawXml).toContain('r:embed="{pic.png}"');
    expect(raw.rawXml).not.toContain("rId5");
    expect(raw.rawMedia).toMatchObject([{ fileName: "pic.png", type: "png" }]);
    expect(Array.from(raw.rawMedia?.[0]?.data ?? [])).toEqual(Array.from(PNG));
  });

  it("generates a package whose unrecognized-drawing refs resolve", async () => {
    const drawing =
      '<w:r><w:drawing><wps:wsp><wps:cNvPr id="1" name="future"/>' +
      '<a:graphic><a:graphicData uri="http://example.com/future"><future:blip r:embed="rId5"/></a:graphicData></a:graphic>' +
      "</wps:wsp></w:drawing></w:r>";
    const out = await generateDocument(
      parseDocumentSync(
        packageOf(singleParagraph(drawing), [
          { path: "media/pic.png", relationshipType: IMAGE_REL, body: "" },
        ]),
      ),
      { type: "uint8array" },
    );
    const files = unzipSync(out);
    expect(strFromU8(files["word/document.xml"] as Uint8Array)).not.toContain("{pic.png}");
    expect(files["word/media/pic.png"]).toBeInstanceOf(Uint8Array);
    expectNoDanglingReferences(files);
  });

  it("keeps unknown track-change wrapper markup verbatim with placeholder media", () => {
    const trackChange =
      '<w:ins w:id="1" w:author="A" w:date="2024-01-01T00:00:00Z">' +
      "<w:r><w:t>kept</w:t></w:r>" +
      '<w15:futureMarkup w15:x="1" r:id="rId5"/>' +
      "</w:ins>";
    const source = packageOf(singleParagraph(trackChange), [
      { path: "media/pic.png", relationshipType: IMAGE_REL, body: "" },
    ]);
    const children = firstParagraphChildren(source);
    const insertion = (
      children[0] as { insertion: { children: Array<RawXmlChild | { text?: string }> } }
    ).insertion;

    const raw = insertion.children.find(
      (c): c is RawXmlChild => typeof (c as RawXmlChild).rawXml === "string",
    );
    expect(raw?.rawXml).toContain("<w15:futureMarkup");
    expect(raw?.rawXml).toContain('r:id="{pic.png}"');
    expect(raw?.rawXml).not.toContain("rId5");
    expect(raw?.rawMedia).toMatchObject([{ fileName: "pic.png", type: "png" }]);
  });

  it("generates a package whose track-change fallback refs resolve", async () => {
    const trackChange =
      '<w:ins w:id="1" w:author="A" w:date="2024-01-01T00:00:00Z">' +
      "<w:r><w:t>kept</w:t></w:r>" +
      '<w15:futureMarkup w15:x="1" r:id="rId5"/>' +
      "</w:ins>";
    const out = await generateDocument(
      parseDocumentSync(
        packageOf(singleParagraph(trackChange), [
          { path: "media/pic.png", relationshipType: IMAGE_REL, body: "" },
        ]),
      ),
      { type: "uint8array" },
    );
    expectNoDanglingReferences(unzipSync(out));
  });

  it("carries an unknown block-level section child with placeholder media", async () => {
    const source = packageOf('<w:futureBlock r:id="rId5"><w:p/></w:futureBlock>', [
      { path: "media/pic.png", relationshipType: IMAGE_REL, body: "" },
    ]);
    const parsed = parseDocumentSync(source);
    const child = parsed.sections?.[0]?.children?.[0] as RawXmlChild;

    expect(child.rawXml).toContain("<w:futureBlock");
    expect(child.rawXml).toContain('r:id="{pic.png}"');
    expect(child.rawXml).not.toContain("rId5");
    expect(child.rawMedia).toMatchObject([{ fileName: "pic.png", type: "png" }]);

    const files = unzipSync(await generateDocument(parsed, { type: "uint8array" }));
    expect(files["word/media/pic.png"]).toBeInstanceOf(Uint8Array);
    expectNoDanglingReferences(files);
  });

  it("re-registers a header w:contentPart relationship instead of dangling", async () => {
    const source = zipSync({
      "[Content_Types].xml": strToU8(
        '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/><Override PartName="/word/ink/ink1.xml" ContentType="application/xml"/></Types>',
      ),
      "_rels/.rels": strToU8(
        '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
      ),
      "word/_rels/document.xml.rels": strToU8(
        '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/></Relationships>',
      ),
      "word/_rels/header1.xml.rels": strToU8(
        `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId5" Type="${INK_REL}" Target="ink/ink1.xml"/>${'<Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/pic.png"/>'}</Relationships>`,
      ),
      "word/document.xml": strToU8(
        `<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="${W_NS}" xmlns:r="${R_NS}"><w:body><w:p><w:r><w:t>b</w:t></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="rId4"/></w:sectPr></w:body></w:document>`,
      ),
      "word/header1.xml": strToU8(
        `<?xml version="1.0" encoding="UTF-8"?><w:hdr xmlns:w="${W_NS}" xmlns:r="${R_NS}"><w:p><w:r><w:rPr><w:contentPart r:id="rId5"/></w:rPr><w:t>h</w:t></w:r><w:r><w:drawing><wp:inline><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:blipFill><a:blip r:embed="rId6"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p></w:hdr>`,
      ),
      "word/ink/ink1.xml": strToU8('<ink xmlns="http://schemas.microsoft.com/ink/2010/main"/>'),
      "word/media/pic.png": PNG,
    });

    const parsed = parseDocumentSync(source);
    const headerChildren = parsed.sections?.[0]?.headers?.default ?? [];
    const headerParagraph = headerChildren[0] as { paragraph: { children?: unknown[] } };
    expect(headerParagraph.paragraph.children?.[0]).toMatchObject({ contentPartRId: "rId5" });

    const files = unzipSync(await generateDocument(parsed, { type: "uint8array" }));
    const headerXml = strFromU8(files["word/header1.xml"] as Uint8Array);
    // The source ink relationship is re-registered on the rebuilt header rels,
    // so the verbatim r:id resolves instead of dangling.
    expect(headerXml).toContain('<w:contentPart r:id="rId5"/>');
    const headerRels = strFromU8(files["word/_rels/header1.xml.rels"] as Uint8Array);
    expect(relTargetsOf(headerRels).get("rId5")).toBe("ink/ink1.xml");
    expectNoDanglingReferences(files);
  });
});

describe("passthrough story relationships", () => {
  it("keeps a non-numeric source id (rIdChartEx) and its header story rel", async () => {
    // A chartEx drawing is unmodeled (PKG-12 preserves it), so the header keeps
    // the source r:id verbatim. The rebuilt header .rels must re-emit that exact
    // id string — renumbering it to a fresh numeric id dangles the reference and
    // Word shows a silent blank where the chart should be.
    const CHART_EX_REL = "http://schemas.microsoft.com/office/2014/relationships/chartEx";
    const source = zipSync({
      "[Content_Types].xml": strToU8(
        '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/><Override PartName="/word/charts/chartEx1.xml" ContentType="application/vnd.ms-office.chartex+xml"/></Types>',
      ),
      "_rels/.rels": strToU8(
        '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
      ),
      "word/_rels/document.xml.rels": strToU8(
        '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/></Relationships>',
      ),
      "word/_rels/header1.xml.rels": strToU8(
        `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdChartEx" Type="${CHART_EX_REL}" Target="charts/chartEx1.xml"/></Relationships>`,
      ),
      "word/document.xml": strToU8(
        `<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="${W_NS}" xmlns:r="${R_NS}"><w:body><w:p><w:r><w:t>b</w:t></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="rId4"/></w:sectPr></w:body></w:document>`,
      ),
      "word/header1.xml": strToU8(
        `<?xml version="1.0" encoding="UTF-8"?><w:hdr xmlns:w="${W_NS}" xmlns:r="${R_NS}" xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"><w:p><w:r><w:drawing><cx:chart r:id="rIdChartEx"/></w:drawing></w:r></w:p></w:hdr>`,
      ),
      "word/charts/chartEx1.xml": strToU8(
        '<cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"/>',
      ),
    });

    const files = unzipSync(
      await generateDocument(parseDocumentSync(source), { type: "uint8array" }),
    );
    const headerXml = strFromU8(files["word/header1.xml"] as Uint8Array);
    expect(headerXml).toContain('r:id="rIdChartEx"');
    const headerRels = strFromU8(files["word/_rels/header1.xml.rels"] as Uint8Array);
    expect(relTargetsOf(headerRels).get("rIdChartEx")).toBe("charts/chartEx1.xml");
    expect(files["word/charts/chartEx1.xml"]).toBeInstanceOf(Uint8Array);
    expectNoDanglingReferences(files);
  });
});
