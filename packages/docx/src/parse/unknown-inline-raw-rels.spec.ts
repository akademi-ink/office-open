import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { describe, expect, it } from "vite-plus/test";

import { compileDocument } from "../compiler";
import { generateDocument } from "../generate";
import { parseDocumentSync } from "../parse";

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const V_NS = "urn:schemas-microsoft-com:vml";
const MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const IMAGE_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

const PNG = new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10]);

interface RawXmlChild {
  rawXml: string;
  rawMedia?: Array<{ fileName: string; type: string; data: Uint8Array }>;
}

/** A minimal package whose body is `bodyXml`, with one media relationship. */
function packageOf(bodyXml: string): Uint8Array {
  return zipSync({
    "[Content_Types].xml": strToU8(
      '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>',
    ),
    "_rels/.rels": strToU8(
      '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
    ),
    "word/_rels/document.xml.rels": strToU8(
      `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId5" Type="${IMAGE_REL}" Target="media/pic.png"/></Relationships>`,
    ),
    "word/document.xml": strToU8(
      `<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="${W_NS}" xmlns:r="${R_NS}" xmlns:v="${V_NS}" xmlns:mc="${MC_NS}"><w:body>${bodyXml}</w:body></w:document>`,
    ),
    "word/media/pic.png": PNG,
  });
}

// One unknown element per preservation site, each with a resolvable relationship:
// a run-level member (run-parse default), a smartTag container child
// (parseContainerChildren default) and a paragraph-level member
// (parseRunLevelChildren default).
const BODY =
  "<w:p>" +
  "<w:r><w:t>before</w:t></w:r>" +
  '<w:r><w:contentPart r:id="rId5"/></w:r>' +
  '<w:smartTag w:uri="urn:test" w:element="future"><w:futureMedia r:embed="rId5"/></w:smartTag>' +
  '<mc:AlternateContent><mc:Choice Requires="wps"><w:drawing/></mc:Choice><mc:Fallback><v:shape><v:imagedata r:id="rId5"/></v:shape></mc:Fallback></mc:AlternateContent>' +
  "</w:p>";

function paragraphChildren(source: Uint8Array): unknown[] {
  const parsed = parseDocumentSync(source);
  const first = parsed.sections?.[0]?.children?.[0] as
    | { paragraph: { children?: unknown[] } }
    | undefined;
  return first?.paragraph.children ?? [];
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

describe("rawXml preservation with relationship refs", () => {
  it("rewrites r:id/r:embed refs to {fileName} placeholders and collects rawMedia", () => {
    const children = paragraphChildren(packageOf(BODY));

    const runChildren = (children[1] as { children: RawXmlChild[] }).children;
    expect(runChildren).toHaveLength(1);
    expect(runChildren[0]?.rawXml).toContain('r:id="{pic.png}"');
    expect(runChildren[0]?.rawXml).not.toContain("rId5");
    expect(runChildren[0]?.rawMedia).toMatchObject([{ fileName: "pic.png", type: "png" }]);
    expect(Array.from(runChildren[0]?.rawMedia?.[0]?.data ?? [])).toEqual(Array.from(PNG));

    const smartTag = (children[2] as { smartTag: { element: string; children: RawXmlChild[] } })
      .smartTag;
    expect(smartTag.element).toBe("future");
    expect(smartTag.children[0]?.rawXml).toContain('r:embed="{pic.png}"');
    expect(smartTag.children[0]?.rawMedia).toMatchObject([{ fileName: "pic.png", type: "png" }]);

    const paragraphXml = children[3] as RawXmlChild;
    expect(paragraphXml.rawXml).toContain("<mc:AlternateContent>");
    expect(paragraphXml.rawXml).toContain('r:id="{pic.png}"');
    expect(paragraphXml.rawMedia).toMatchObject([{ fileName: "pic.png", type: "png" }]);
  });

  it("generates a package whose rawXml refs resolve to the registered media", async () => {
    const output = await generateDocument(parseDocumentSync(packageOf(BODY)), {
      type: "uint8array",
    });
    const files = unzipSync(output);
    const xml = strFromU8(files["word/document.xml"] as Uint8Array);
    const relsXml = strFromU8(files["word/_rels/document.xml.rels"] as Uint8Array);

    expect(xml).not.toContain("{pic.png}");
    const refs = [...xml.matchAll(/\br:(?:id|embed|link)="([^"]+)"/g)].map((m) => m[1]!);
    expect(refs).toHaveLength(3);
    const targets = relTargetsOf(relsXml);
    for (const ref of refs) {
      expect(targets.get(ref)).toMatch(/^media\//);
    }
    expect(files["word/media/pic.png"]).toBeInstanceOf(Uint8Array);
  });

  it("registers rawMedia on generate even without source passthrough rels", () => {
    const files = compileDocument({
      sections: [
        {
          children: [
            {
              paragraph: {
                children: [
                  {
                    rawXml: '<w:contentPart r:id="{pic.png}"/>',
                    rawMedia: [{ fileName: "pic.png", type: "png", data: PNG }],
                  },
                ],
              },
            },
          ],
        },
      ],
    } as Parameters<typeof compileDocument>[0]);
    const unzipped = unzipSync(zipSync(files));
    const xml = strFromU8(unzipped["word/document.xml"] as Uint8Array);
    const relsXml = strFromU8(unzipped["word/_rels/document.xml.rels"] as Uint8Array);

    expect(xml).not.toContain("{pic.png}");
    expect([...relTargetsOf(relsXml).values()]).toContain("media/pic.png");
    expect(unzipped["word/media/pic.png"]).toBeInstanceOf(Uint8Array);
  });
});
