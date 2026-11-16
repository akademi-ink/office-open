import { zipSync } from "fflate";
import { describe, expect, it } from "vite-plus/test";

import { parseDocumentSync } from "../parse";

/**
 * A `w:pict` is a run-level container for any number of VML shapes. The body
 * reader promotes a paragraph to the block `{ textbox }` branch only when the
 * pict *is* one text box; a pict that also carries connectors (a hand-drawn
 * diagram of boxes and arrows) or a picture must stay a paragraph run, or
 * every sibling shape is dropped with no trace.
 */
const NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
  + 'xmlns:v="urn:schemas-microsoft-com:vml" '
  + 'xmlns:o="urn:schemas-microsoft-com:office:office" '
  + 'xmlns:w10="urn:schemas-microsoft-com:office:word" '
  + 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';

function docx(inner: string): Uint8Array {
  const document = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`
    + `<w:document ${NS}><w:body>${inner}</w:body></w:document>`;
  const contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`
    + `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">`
    + `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>`
    + `<Default Extension="xml" ContentType="application/xml"/>`
    + `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>`
    + `</Types>`;
  const rootRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`
    + `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`
    + `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>`
    + `</Relationships>`;
  const documentRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`
    + `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>`;
  return zipSync({
    "[Content_Types].xml": new TextEncoder().encode(contentTypes),
    "_rels/.rels": new TextEncoder().encode(rootRels),
    "word/document.xml": new TextEncoder().encode(document),
    "word/_rels/document.xml.rels": new TextEncoder().encode(documentRels),
  });
}

const vmlTextbox = (id: string, text: string): string =>
  `<v:shape id="${id}" type="#_x0000_t202" style="width:120pt;height:40pt">`
  + `<v:textbox><w:txbxContent><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:txbxContent></v:textbox>`
  + `</v:shape>`;

const connector = (id: string): string =>
  `<v:shape id="${id}" type="#_x0000_t32" style="width:200pt;height:.5pt"><v:path arrowok="t"/></v:shape>`;

describe("VML pict promotion", () => {
  it("keeps a pict of one text box as the block textbox branch", () => {
    const parsed = parseDocumentSync(docx(`<w:p><w:r><w:pict>${vmlTextbox("TextBox 1", "caption")}</w:pict></w:r></w:p>`));
    const child = parsed.sections[0]!.children[0] as { textbox?: unknown };
    expect(child.textbox).toBeDefined();
  });

  it("keeps every shape of a multi-shape diagram pict inside the paragraph run", () => {
    const parsed = parseDocumentSync(docx(
      `<w:p><w:r><w:pict>${connector("ConnectorOne")}${vmlTextbox("TextBox 60", "Pramesti dkk")}${vmlTextbox("TextBox 61", "Santoso dkk")}</w:pict></w:r></w:p>`,
    ));
    const child = parsed.sections[0]!.children[0] as { textbox?: unknown; paragraph?: unknown };
    expect(child.textbox).toBeUndefined();
    expect(child.paragraph).toBeDefined();
    const serialized = JSON.stringify(child);
    expect(serialized).toContain("ConnectorOne");
    expect(serialized).toContain("Pramesti dkk");
    expect(serialized).toContain("Santoso dkk");
  });

  it("keeps a pict whose only shape is a picture inside the paragraph run", () => {
    const parsed = parseDocumentSync(docx(
      `<w:p><w:r><w:pict><v:shape id="PictureShape" type="#_x0000_t75" style="width:40pt;height:40pt"><v:imagedata r:id="rId7" o:title=""/></v:shape></w:pict></w:r></w:p>`,
    ));
    const child = parsed.sections[0]!.children[0] as { textbox?: unknown; paragraph?: unknown };
    expect(child.textbox).toBeUndefined();
    expect(child.paragraph).toBeDefined();
    expect(JSON.stringify(child)).toContain("PictureShape");
  });
});
