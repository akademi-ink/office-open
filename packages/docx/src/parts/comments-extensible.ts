/**
 * Comments-extensible descriptor — produces word/commentsExtensible.xml
 * (Word 2018+ comment metadata keyed by durable id).
 *
 * Each w16cex:commentExtensible carries the durable id allocated in
 * word/commentsIds.xml (w16cid) plus the comment's UTC creation time, so a
 * consumer can follow a comment across edits without matching author strings
 * or the transient w:comment/@w:id.
 *
 * Reference: MS-OE376 / wml-cex-2018.xsd, CT_CommentsExtensible /
 * CT_CommentExtensible
 *
 * @module
 */

import { parseOnOff } from "@office-open/core";
import type { CustomDescriptor } from "@office-open/core/descriptor";
import { attr, escapeXml, findChild } from "@office-open/xml";
import type { Element } from "@office-open/xml";

import { stringifyElement } from "../util/stringify-element";
import { COMMENTS_NS } from "./comments";

/** One extensible comment entry (w16cex:commentExtensible). */
export interface CommentExtensibleOptions {
  /**
   * Durable id from word/commentsIds.xml (w16cex:durableId, required) — the
   * link back to the comment; never recomputed here.
   */
  durableId: string;
  /** UTC creation time (w16cex:dateUtc, ISO 8601 string). */
  dateUtc?: string;
  /** Placeholder-only comment flag (w16cex:intelligentPlaceholder). */
  intelligentPlaceholder?: boolean;
  /**
   * Verbatim `w16cex:extLst` XML (the per-comment extension escape hatch),
   * carried byte-faithfully. Round-trip only — do not hand-author.
   */
  extLst?: string;
}

export const commentsExtensibleDesc: CustomDescriptor<CommentExtensibleOptions[]> = {
  kind: "custom",

  stringify(opts) {
    const parts: string[] = [`<w16cex:commentsExtensible ${COMMENTS_NS}>`];
    for (const entry of opts) {
      const attrs = [`w16cex:durableId="${escapeXml(entry.durableId)}"`];
      if (entry.dateUtc !== undefined) attrs.push(`w16cex:dateUtc="${escapeXml(entry.dateUtc)}"`);
      if (entry.intelligentPlaceholder !== undefined)
        attrs.push(`w16cex:intelligentPlaceholder="${entry.intelligentPlaceholder ? 1 : 0}"`);
      parts.push(
        entry.extLst
          ? `<w16cex:commentExtensible ${attrs.join(" ")}>${entry.extLst}</w16cex:commentExtensible>`
          : `<w16cex:commentExtensible ${attrs.join(" ")}/>`,
      );
    }
    parts.push("</w16cex:commentsExtensible>");
    return parts.join("");
  },

  parse(el) {
    const entries: CommentExtensibleOptions[] = [];
    for (const child of el.elements ?? []) {
      if (child.name !== "w16cex:commentExtensible") continue;
      const durableId = attr(child, "w16cex:durableId");
      if (durableId === undefined) continue;
      const entry: Partial<CommentExtensibleOptions> = { durableId };
      const dateUtc = attr(child, "w16cex:dateUtc");
      if (dateUtc !== undefined) entry.dateUtc = dateUtc;
      const placeholder = attr(child, "w16cex:intelligentPlaceholder");
      if (placeholder !== undefined)
        entry.intelligentPlaceholder = parseOnOff(placeholder) ?? false;
      const extLst = findChild(child, "w16cex:extLst") ?? findChild(child, "w16:extLst");
      if (extLst) entry.extLst = stringifyElement(extLst as Element);
      entries.push(entry as CommentExtensibleOptions);
    }
    return entries;
  },
};
