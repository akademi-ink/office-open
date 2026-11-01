/**
 * Comments-ids descriptor — produces word/commentsIds.xml
 * (Word 2016+ durable comment identifiers).
 *
 * Each w16cid:commentId pairs the w14:paraId of a comment's first paragraph
 * with a durable id that survives edits and reordering. Word keys the durable
 * channel by the paraId, not by w:comment/@w:id, and the durable id is what
 * w16cex:commentsExtensible/commentExtensible entries reference.
 *
 * Reference: MS-OE376 / wml-cid-2016.xsd, CT_CommentsIds / CT_CommentId
 *
 * @module
 */

import type { LongHexNumber } from "@office-open/core";
import type { CustomDescriptor } from "@office-open/core/descriptor";
import { attr, escapeXml } from "@office-open/xml";

import { COMMENTS_NS } from "./comments";

/** One durable comment id (w16cid:commentId). */
export interface CommentIdOptions {
  /** w14:paraId of the comment's first paragraph (w16cid:paraId, required). */
  paraId: LongHexNumber;
  /**
   * Durable identifier Word uses to track the comment across edits
   * (w16cid:durableId, required). Word writes an opaque numeric token, so it
   * round-trips verbatim as a string.
   */
  durableId: string;
}

export const commentsIdsDesc: CustomDescriptor<CommentIdOptions[]> = {
  kind: "custom",

  stringify(opts) {
    const parts: string[] = [`<w16cid:commentsIds ${COMMENTS_NS}>`];
    for (const entry of opts) {
      parts.push(
        `<w16cid:commentId w16cid:paraId="${escapeXml(entry.paraId)}" w16cid:durableId="${escapeXml(entry.durableId)}"/>`,
      );
    }
    parts.push("</w16cid:commentsIds>");
    return parts.join("");
  },

  parse(el) {
    const entries: CommentIdOptions[] = [];
    for (const child of el.elements ?? []) {
      if (child.name !== "w16cid:commentId") continue;
      const paraId = attr(child, "w16cid:paraId");
      const durableId = attr(child, "w16cid:durableId");
      if (paraId === undefined || durableId === undefined) continue;
      entries.push({ paraId, durableId });
    }
    return entries;
  },
};
