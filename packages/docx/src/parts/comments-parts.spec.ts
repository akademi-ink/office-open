import { parse as parseXml } from "@office-open/xml";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { describe, expect, it } from "vite-plus/test";

import { generateDocument } from "../generate";
import { parseDocumentSync } from "../parse";
import { COMMENTS_NS } from "./comments";
import { commentsExtensibleDesc } from "./comments-extensible";
import { commentsIdsDesc } from "./comments-ids";
import { linkCommentMetadata } from "./comments-link";

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const W14_NS = "http://schemas.microsoft.com/office/word/2010/wordml";
const W15_NS = "http://schemas.microsoft.com/office/word/2012/wordml";

const COMMENTS_IDS_REL = "http://schemas.microsoft.com/office/2016/09/relationships/commentsIds";
const COMMENTS_EXT_REL =
  "http://schemas.microsoft.com/office/2018/08/relationships/commentsExtensible";
const COMMENTS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";

const PARA_ID = "1A2B3C4D";
const DURABLE_ID = "1094861568";

const COMMENTS_XML =
  `<?xml version="1.0" encoding="UTF-8"?><w:comments xmlns:w="${W_NS}" xmlns:w14="${W14_NS}">` +
  `<w:comment w:id="1" w:author="Ada" w:initials="A" w:date="2024-01-01T00:00:00Z">` +
  `<w:p w14:paraId="${PARA_ID}"><w:r><w:t>look here</w:t></w:r></w:p></w:comment></w:comments>`;

const COMMENTS_IDS_XML =
  `<?xml version="1.0" encoding="UTF-8"?><w16cid:commentsIds xmlns:w="${W_NS}" ` +
  `xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid">` +
  `<w16cid:commentId w16cid:paraId="${PARA_ID}" w16cid:durableId="${DURABLE_ID}"/></w16cid:commentsIds>`;

const COMMENTS_EXT_XML =
  `<?xml version="1.0" encoding="UTF-8"?><w16cex:commentsExtensible xmlns:w="${W_NS}" ` +
  `xmlns:w16="http://schemas.microsoft.com/office/word/2018/wordml" ` +
  `xmlns:w16cex="http://schemas.microsoft.com/office/word/2018/wordml/cex">` +
  `<w16cex:commentExtensible w16cex:durableId="${DURABLE_ID}" w16cex:dateUtc="2024-01-01T09:30:00Z"/>` +
  `</w16cex:commentsExtensible>`;

const COMMENTS_EX_XML =
  `<?xml version="1.0" encoding="UTF-8"?><w15:commentsEx xmlns:w="${W_NS}" ` +
  `xmlns:w15="${W15_NS}"><w15:commentEx w15:paraId="${PARA_ID}" w15:done="1"/>` +
  `</w15:commentsEx>`;

/** Package with the full Word comment infrastructure and one anchored comment. */
function packageWithCommentParts(): Uint8Array {
  const rels =
    `<Relationship Id="rId3" Type="${COMMENTS_REL}" Target="comments.xml"/>` +
    `<Relationship Id="rId4" Type="${COMMENTS_IDS_REL}" Target="commentsIds.xml"/>` +
    `<Relationship Id="rId5" Type="${COMMENTS_EXT_REL}" Target="commentsExtensible.xml"/>` +
    `<Relationship Id="rId6" Type="http://schemas.microsoft.com/office/2011/relationships/commentsExtended" Target="commentsExtended.xml"/>`;
  return zipSync({
    "[Content_Types].xml": strToU8(
      `<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/><Override PartName="/word/commentsIds.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.commentsIds+xml"/><Override PartName="/word/commentsExtensible.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtensible+xml"/><Override PartName="/word/commentsExtended.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml"/></Types>`,
    ),
    "_rels/.rels": strToU8(
      '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
    ),
    "word/_rels/document.xml.rels": strToU8(
      `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels}</Relationships>`,
    ),
    "word/document.xml": strToU8(
      `<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="${W_NS}" xmlns:w14="${W14_NS}" xmlns:r="${R_NS}"><w:body><w:p><w:commentRangeStart w:id="1"/><w:r><w:t>body</w:t></w:r><w:commentRangeEnd w:id="1"/><w:r><w:commentReference w:id="1"/></w:r></w:p></w:body></w:document>`,
    ),
    "word/comments.xml": strToU8(COMMENTS_XML),
    "word/commentsIds.xml": strToU8(COMMENTS_IDS_XML),
    "word/commentsExtensible.xml": strToU8(COMMENTS_EXT_XML),
    "word/commentsExtended.xml": strToU8(COMMENTS_EX_XML),
  });
}

describe("commentsIds descriptor (w16cid)", () => {
  it("declares w16cid ignorable on a root that emits it", () => {
    expect(COMMENTS_NS).toContain(
      'xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid"',
    );
    expect(COMMENTS_NS).toMatch(/mc:Ignorable="[^"]*\bw16cid\b/);
  });

  it("round-trips commentId entries", () => {
    const entries = [{ paraId: PARA_ID, durableId: DURABLE_ID }];
    const xml = commentsIdsDesc.stringify(entries, undefined as never)!;
    expect(xml).toContain(
      `<w16cid:commentId w16cid:paraId="${PARA_ID}" w16cid:durableId="${DURABLE_ID}"/>`,
    );
    const parsed = commentsIdsDesc.parse(parseXml(xml).elements![0]!, undefined as never);
    expect(parsed).toEqual(entries);
  });
});

describe("commentsExtensible descriptor (w16cex)", () => {
  it("declares w16cex ignorable on a root that emits it", () => {
    expect(COMMENTS_NS).toContain(
      'xmlns:w16cex="http://schemas.microsoft.com/office/word/2018/wordml/cex"',
    );
    expect(COMMENTS_NS).toMatch(/mc:Ignorable="[^"]*\bw16cex\b/);
  });

  it("round-trips durableId/dateUtc/intelligentPlaceholder and a verbatim extLst", () => {
    const entries = [
      {
        durableId: DURABLE_ID,
        dateUtc: "2024-01-01T09:30:00Z",
        intelligentPlaceholder: false,
        extLst: '<w16:extLst><w16:ext uri="{X}"/></w16:extLst>',
      },
    ];
    const xml = commentsExtensibleDesc.stringify(entries, undefined as never)!;
    expect(xml).toContain(`w16cex:durableId="${DURABLE_ID}"`);
    expect(xml).toContain('w16cex:dateUtc="2024-01-01T09:30:00Z"');
    expect(xml).toContain('w16cex:intelligentPlaceholder="0"');
    expect(xml).toContain("<w16:extLst>");
    const parsed = commentsExtensibleDesc.parse(parseXml(xml).elements![0]!, undefined as never);
    expect(parsed).toEqual(entries);
  });
});

describe("comment parts through a package", () => {
  it("parses and preserves the w16 parts", async () => {
    const parsed = parseDocumentSync(packageWithCommentParts());
    expect(parsed.commentsIds).toEqual([{ paraId: PARA_ID, durableId: DURABLE_ID }]);
    expect(parsed.commentsExtensible).toEqual([
      { durableId: DURABLE_ID, dateUtc: "2024-01-01T09:30:00Z" },
    ]);
    expect(parsed.commentsExtended).toEqual([{ paraId: PARA_ID, done: true }]);

    const files = unzipSync(await generateDocument(parsed, { type: "uint8array" }));
    const ids = strFromU8(files["word/commentsIds.xml"] as Uint8Array);
    expect(ids).toContain(`w16cid:paraId="${PARA_ID}"`);
    expect(ids).toContain(`w16cid:durableId="${DURABLE_ID}"`);
    const ext = strFromU8(files["word/commentsExtensible.xml"] as Uint8Array);
    expect(ext).toContain(`w16cex:durableId="${DURABLE_ID}"`);

    // Relationships + content types stay consistent with the emitted parts.
    const rels = strFromU8(files["word/_rels/document.xml.rels"] as Uint8Array);
    expect(rels).toContain(`Type="${COMMENTS_IDS_REL}" Target="commentsIds.xml"`);
    expect(rels).toContain(`Type="${COMMENTS_EXT_REL}" Target="commentsExtensible.xml"`);
    const contentTypes = strFromU8(files["[Content_Types].xml"] as Uint8Array);
    expect(contentTypes).toContain("wordprocessingml.commentsIds+xml");
    expect(contentTypes).toContain("wordprocessingml.commentsExtensible+xml");

    // Second parse sees the same parts: parse → generate → parse is stable.
    const reparsed = parseDocumentSync(
      zipSync(
        Object.fromEntries(Object.entries(files).map(([path, data]) => [path, data as Uint8Array])),
      ),
    );
    expect(reparsed.commentsIds).toEqual(parsed.commentsIds);
    expect(reparsed.commentsExtensible).toEqual(parsed.commentsExtensible);
  });

  it("does not emit the parts when the model carries none", async () => {
    const files = unzipSync(
      await generateDocument(
        { sections: [{ children: [{ paragraph: { text: "x" } }] }] },
        {
          type: "uint8array",
        },
      ),
    );
    expect(files["word/commentsIds.xml"]).toBeUndefined();
    expect(files["word/commentsExtensible.xml"]).toBeUndefined();
  });

  it("links each comment to its metadata through the paraId → durableId chain", () => {
    const parsed = parseDocumentSync(packageWithCommentParts());
    const linked = linkCommentMetadata(
      parsed.comments,
      parsed.commentsExtended,
      parsed.commentsIds,
      parsed.commentsExtensible,
    );
    expect(linked).toEqual([
      {
        id: 1,
        paraId: PARA_ID,
        durableId: DURABLE_ID,
        done: true,
        dateUtc: "2024-01-01T09:30:00Z",
      },
    ]);
  });

  it("resolves metadata without any author-name matching", () => {
    // Two reviewers sharing a display name must not merge: the linkage is the
    // paraId/durableId chain, and an entry with no comment keeps its own id.
    const linked = linkCommentMetadata(
      [
        { id: 7, author: "Alex", children: [{ paraId: "AAAAAAAA" }] },
        { id: 8, author: "Alex", children: [{ paraId: "BBBBBBBB" }] },
      ],
      [
        { paraId: "AAAAAAAA", paraIdParent: "BBBBBBBB" },
        { paraId: "BBBBBBBB", done: true },
      ],
      [
        { paraId: "AAAAAAAA", durableId: "d1" },
        { paraId: "BBBBBBBB", durableId: "d2" },
      ],
      [{ durableId: "d1", dateUtc: "2024-02-02T00:00:00Z" }],
    );
    expect(linked).toEqual([
      {
        id: 7,
        paraId: "AAAAAAAA",
        durableId: "d1",
        paraIdParent: "BBBBBBBB",
        dateUtc: "2024-02-02T00:00:00Z",
      },
      { id: 8, paraId: "BBBBBBBB", durableId: "d2", done: true },
    ]);
  });
});
