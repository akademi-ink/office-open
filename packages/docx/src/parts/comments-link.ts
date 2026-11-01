/**
 * Comment metadata linkage across the Word comment channels.
 *
 * Word does not key a comment's metadata by `w:comment/@w:id` (that id is
 * reassigned as comments are added and removed) and it does not key it by
 * author either. The linkage is:
 *
 * ```text
 * w:comment/@w:id
 *   → w14:paraId of the comment's first paragraph
 *   → w16cid:commentId/@paraId        → @durableId      (Word 2016+)
 *   → w16cex:commentExtensible/@durableId → @dateUtc     (Word 2018+)
 *   → w15:commentEx/@paraId           → @paraIdParent, @done (Word 2013+)
 * ```
 *
 * {@link linkCommentMetadata} resolves that chain once so callers never pair a
 * comment with its thread state, resolved flag, or creation time by matching
 * author strings — which silently merges two reviewers who share a display
 * name and breaks the moment a comment is edited or reply-ordered.
 *
 * `word/people.xml` is the one part with no durable link in the schema: a
 * `w15:person` carries only `@w15:author` / `@w15:contact`, so it remains a
 * display-name directory, not an identity key.
 *
 * @module
 */

import type { CommentExtendedOptions } from "./comments-extended";
import type { CommentExtensibleOptions } from "./comments-extensible";
import type { CommentIdOptions } from "./comments-ids";
import type { CommentOptions } from "./paragraph/run/comment-run";

/** One comment's resolved metadata, keyed by the durable channel. */
export interface LinkedCommentMetadata {
  /** `w:comment/@w:id`. */
  id: number;
  /** `w14:paraId` of the comment's first paragraph — the metadata key. */
  paraId?: string;
  /** Durable id (`w16cid`), stable across edits. */
  durableId?: string;
  /** Reply parent's first-paragraph paraId (`w15:paraIdParent`). */
  paraIdParent?: string;
  /** Resolved state (`w15:done`). */
  done?: boolean;
  /** Next-gen comment variant (`w15:commentExNG`). */
  nextGen?: boolean;
  /** Comment creation time (`w16cex:dateUtc`). */
  dateUtc?: string;
  /** Placeholder-only comment (`w16cex:intelligentPlaceholder`). */
  intelligentPlaceholder?: boolean;
}

/** First paragraph's `w14:paraId`, the key every metadata channel uses. */
function firstParagraphParaId(comment: CommentOptions): string | undefined {
  const first = comment.children[0];
  if (first === null || typeof first !== "object") return undefined;
  const paraId = (first as { paraId?: unknown }).paraId;
  return typeof paraId === "string" ? paraId : undefined;
}

/**
 * Resolve each comment's extended metadata through the paraId → durableId
 * chain, in `comments` order. Comments without a first-paragraph `w14:paraId`
 * (or with metadata the source never wrote) resolve to their id alone.
 */
export function linkCommentMetadata(
  comments: readonly CommentOptions[] | undefined,
  commentsExtended?: readonly CommentExtendedOptions[],
  commentsIds?: readonly CommentIdOptions[],
  commentsExtensible?: readonly CommentExtensibleOptions[],
): LinkedCommentMetadata[] {
  const extendedByParaId = new Map<string, CommentExtendedOptions>();
  for (const entry of commentsExtended ?? []) extendedByParaId.set(entry.paraId, entry);

  const durableIdByParaId = new Map<string, string>();
  for (const entry of commentsIds ?? []) durableIdByParaId.set(entry.paraId, entry.durableId);

  const extensibleByDurableId = new Map<string, CommentExtensibleOptions>();
  for (const entry of commentsExtensible ?? []) {
    extensibleByDurableId.set(entry.durableId, entry);
  }

  return (comments ?? []).map((comment) => {
    const linked: LinkedCommentMetadata = { id: comment.id };
    const paraId = firstParagraphParaId(comment);
    if (paraId === undefined) return linked;
    linked.paraId = paraId;

    const extended = extendedByParaId.get(paraId);
    if (extended) {
      if (extended.paraIdParent !== undefined) linked.paraIdParent = extended.paraIdParent;
      if (extended.done !== undefined) linked.done = extended.done;
      if (extended.nextGen) linked.nextGen = true;
    }

    const durableId = durableIdByParaId.get(paraId);
    if (durableId !== undefined) {
      linked.durableId = durableId;
      const extensible = extensibleByDurableId.get(durableId);
      if (extensible) {
        if (extensible.dateUtc !== undefined) linked.dateUtc = extensible.dateUtc;
        if (extensible.intelligentPlaceholder !== undefined) {
          linked.intelligentPlaceholder = extensible.intelligentPlaceholder;
        }
      }
    }
    return linked;
  });
}
