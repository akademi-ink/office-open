import { unzipSync } from "fflate";
import { describe, expect, it } from "vite-plus/test";

import { generateDocument, generateDocumentSync } from "../generate";
import { parseDocumentSync } from "../parse";
import type { ComplexFieldOptions } from "../parts/paragraph/paragraph";

type FieldResultChild = NonNullable<ComplexFieldOptions["resultChildren"]>[number];

/** A document whose only paragraph holds the field with `resultChildren`. */
function fieldDocument(
  resultChildren: FieldResultChild[],
  instruction = ' HYPERLINK \\l "target" ',
): Uint8Array {
  return generateDocumentSync(
    {
      sections: [
        {
          children: [
            {
              paragraph: {
                children: [{ complexField: { instruction, hasSeparate: true, resultChildren } }],
              },
            },
          ],
        },
      ],
    },
    { type: "uint8array" },
  );
}

function complexFieldOf(bytes: Uint8Array): any {
  const children = parseDocumentSync(bytes).sections?.[0]?.children ?? [];
  return (children[0] as any)?.paragraph?.children?.find((child: any) => child.complexField)
    ?.complexField;
}

const hyperlink = (text = "Link"): FieldResultChild => ({
  hyperlink: { anchor: "target", children: [{ text, color: "0563C1" }] },
});

describe("complex-field typed ordered result", () => {
  it("keeps text · tab · hyperlink in that order with their rPr", () => {
    const cf = complexFieldOf(
      fieldDocument([
        { text: "Sebelum", bold: true },
        { children: [{ tab: true }], font: { ascii: "Arial" } },
        hyperlink(),
      ]),
    );
    expect(cf.resultChildren).toHaveLength(3);
    expect(cf.resultChildren[0]).toMatchObject({ text: "Sebelum", bold: true });
    expect(cf.resultChildren[1]).toMatchObject({ tab: true, font: "Arial" });
    expect(cf.resultChildren[2].hyperlink).toMatchObject({ anchor: "target" });
    expect(cf.resultChildren[2].hyperlink.children[0]).toMatchObject({ text: "Link" });
  });

  it("keeps the reverse hyperlink · tab · text order", () => {
    const cf = complexFieldOf(
      fieldDocument([hyperlink(), { children: [{ tab: true }] }, { text: "Sesudah" }]),
    );
    expect(cf.resultChildren[0].hyperlink).toBeDefined();
    expect(cf.resultChildren[1].tab).toBe(true);
    expect(cf.resultChildren[2]).toMatchObject({ text: "Sesudah" });
  });

  it("does not special-case the word TOC in a non-TOC instruction", () => {
    const cf = complexFieldOf(
      fieldDocument(
        [{ text: "Buka" }, { children: [{ tab: true }] }, hyperlink()],
        ' HYPERLINK "https://example.test/TOC" ',
      ),
    );
    expect(cf.instruction).toContain("TOC");
    expect(cf.resultChildren).toHaveLength(3);
    expect(cf.resultChildren[0]).toMatchObject({ text: "Buka" });
    expect(cf.resultChildren[1]).toMatchObject({ tab: true });
    expect(cf.resultChildren[2].hyperlink).toBeDefined();
  });

  it("keeps a nested field as a typed child beside its siblings", async () => {
    const source = {
      sections: [
        {
          children: [
            {
              paragraph: {
                children: [
                  {
                    complexField: {
                      instruction: ' HYPERLINK \\l "target" ',
                      hasSeparate: true,
                      resultChildren: [
                        { text: "hal " },
                        { hyperlink: { anchor: "target", children: [{ text: "Bab " }] } },
                        {
                          complexField: {
                            instruction: " PAGE ",
                            result: "2",
                            hasSeparate: true,
                            resultChildren: [{ text: "2" }],
                          },
                        },
                      ],
                    },
                  },
                ],
              },
            },
          ],
        },
      ],
    };
    const first = await generateDocument(source as never, { type: "uint8array" });
    const cf = complexFieldOf(first);
    expect(cf.resultChildren).toHaveLength(3);
    expect(cf.resultChildren[0]).toMatchObject({ text: "hal " });
    expect(cf.resultChildren[1].hyperlink).toBeDefined();
    expect(cf.resultChildren[2].complexField.instruction).toContain("PAGE");
    const second = await generateDocument(parseDocumentSync(first), { type: "uint8array" });
    const xml = new TextDecoder().decode(unzipSync(second)["word/document.xml"]);
    expect(xml).toContain(" PAGE ");
    expect(xml).toContain("<w:hyperlink");
  });

  it("writes the typed result in order and reaches a parse fixed point", async () => {
    const source = fieldDocument([{ text: "Sebelum" }, { children: [{ tab: true }] }, hyperlink()]);
    const parsed = parseDocumentSync(source);
    const output = await generateDocument(parsed, { type: "uint8array" });
    const xml = new TextDecoder().decode(unzipSync(output)["word/document.xml"]);
    const positions = ["Sebelum", "<w:tab", "<w:hyperlink"].map((needle) => xml.indexOf(needle));
    expect(positions.every((position) => position >= 0)).toBe(true);
    expect(positions).toEqual([...positions].sort((a, b) => a - b));
    const expected = complexFieldOf(source).resultChildren;
    let current = output;
    for (let generation = 0; generation < 2; generation++) {
      current = await generateDocument(parseDocumentSync(current), { type: "uint8array" });
      expect(complexFieldOf(current).resultChildren).toEqual(expected);
    }
  });
});
