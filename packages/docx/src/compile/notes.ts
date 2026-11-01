/**
 * Notes compile phase: comments, footnotes, and endnotes parts.
 *
 * All three share the same wiring shape — stringify the part, resolve its
 * {fileName} media and {oleObjectN.bin} embedding placeholders, resolve
 * {chart:}/{smartart:} placeholders, register the matching relationships on
 * the part's own rels, and replace numbering placeholders. Collected here so
 * the three stay symmetric.
 *
 * @module
 */

import {
  type XmlifyedFile,
  optionalRelsPart,
  replaceNumberingPlaceholders,
} from "@office-open/core";

import type { PartCtxFactory } from "../compiler";
import type { DocxWriteContext } from "../context";
import { commentsDesc, endnotesDesc, footnotesDesc } from "../parts";
import {
  XML_DECL,
  claimPartPassthroughRelationships,
  registerPartMedia,
  reservePartSourceRids,
  resolvePartCharts,
  resolvePartMedia,
} from "./shared";

/**
 * Comments carried by the document: those the caller listed explicitly
 * (`options.comments`) plus entries registered by `{ comment }` sugar children
 * during body stringification. Drives both word/comments.xml generation and the
 * [Content_Types] comments Override, which must stay in sync (OPC consistency).
 */
function mergedCommentChildren(ctx: DocxWriteContext) {
  return [...(ctx._options.comments ?? []), ...ctx.comments.entries];
}

/** The three notes parts, present only when the document carries content for
 * them (an empty part with a dangling relationship is the OPC violation Word
 * rejects on open). */
export function compileNotesParts(
  ctx: DocxWriteContext,
  mkCtx: PartCtxFactory,
): {
  Comments?: XmlifyedFile;
  CommentsRelationships?: XmlifyedFile;
  FootNotes?: XmlifyedFile;
  FootNotesRelationships?: XmlifyedFile;
  Endnotes?: XmlifyedFile;
  EndnotesRelationships?: XmlifyedFile;
} {
  const result: {
    Comments?: XmlifyedFile;
    CommentsRelationships?: XmlifyedFile;
    FootNotes?: XmlifyedFile;
    FootNotesRelationships?: XmlifyedFile;
    Endnotes?: XmlifyedFile;
    EndnotesRelationships?: XmlifyedFile;
  } = {};

  // Comments: stringify, then resolve media/embeddings and register the
  // relationships before serializing the part's rels.
  const comments = mergedCommentChildren(ctx);
  if (comments.length > 0) {
    const commentsSource = "word/comments.xml";
    reservePartSourceRids(ctx.comments.relationships, ctx, commentsSource);
    const commentCtx = mkCtx({ relationships: ctx.comments.relationships });
    const commentXmlData = XML_DECL + (commentsDesc.stringify(comments, commentCtx) ?? "");
    // Sampled after stringify, like the document and footnote counts below.
    const commentRelCount = ctx.comments.relationships.nextRelationshipId;
    const resolved = resolvePartMedia(commentXmlData, ctx, commentRelCount, commentsSource);
    registerPartMedia(ctx.comments.relationships, ctx, resolved, commentsSource);
    const commentXml = resolvePartCharts(
      resolved.xml,
      ctx,
      ctx.comments.relationships,
      resolved.embeddingOffset + resolved.embeddingRefs.length,
      commentsSource,
    );
    claimPartPassthroughRelationships(ctx.comments.relationships, ctx, commentsSource);
    result.Comments = {
      data: replaceNumberingPlaceholders(commentXml, ctx.numbering.concreteNumbering),
      path: "word/comments.xml",
    };
    result.CommentsRelationships = optionalRelsPart(
      ctx.comments.relationships,
      XML_DECL,
      "word/_rels/comments.xml.rels",
    );
  }

  // Footnotes: stringify and register media/embedding relationships eagerly so
  // the relationshipCount used to gate footnotes.xml.rels reflects the final
  // state (see FootNotesRelationships).
  const footnotesSource = "word/footnotes.xml";
  reservePartSourceRids(ctx.footNotes.relationships, ctx, footnotesSource);
  const footnoteCtx = mkCtx({
    relationships: ctx.footNotes.relationships,
  });
  const footnoteXmlData =
    XML_DECL +
    (footnotesDesc.stringify(
      {
        notes: ctx.footNotes.notes,
        separator: ctx.footNotes.separator,
        continuationSeparator: ctx.footNotes.continuationSeparator,
        continuationNotice: ctx.footNotes.continuationNotice,
      },
      footnoteCtx,
    ) ?? "");
  const footnoteRelCount = ctx.footNotes.relationships.nextRelationshipId;
  const footnoteResolved = resolvePartMedia(
    footnoteXmlData,
    ctx,
    footnoteRelCount,
    footnotesSource,
  );
  registerPartMedia(ctx.footNotes.relationships, ctx, footnoteResolved, footnotesSource);
  const footnoteXml = resolvePartCharts(
    footnoteResolved.xml,
    ctx,
    ctx.footNotes.relationships,
    footnoteResolved.embeddingOffset + footnoteResolved.embeddingRefs.length,
    footnotesSource,
  );
  claimPartPassthroughRelationships(ctx.footNotes.relationships, ctx, footnotesSource);
  if (ctx.hasFootnotes) {
    result.FootNotes = {
      data: replaceNumberingPlaceholders(footnoteXml, ctx.numbering.concreteNumbering),
      path: "word/footnotes.xml",
    };
    if (ctx.footNotes.relationships.relationshipCount > 0) {
      result.FootNotesRelationships = {
        data: XML_DECL + ctx.footNotes.relationships.serialize(),
        path: "word/_rels/footnotes.xml.rels",
      };
    }
  }

  // Endnotes: same eager-registration shape as footnotes.
  if (ctx.hasEndnotes) {
    const endnotesSource = "word/endnotes.xml";
    reservePartSourceRids(ctx.endnotes.relationships, ctx, endnotesSource);
    const endnoteCtx = mkCtx({
      relationships: ctx.endnotes.relationships,
    });
    const endnoteXmlData =
      XML_DECL +
      (endnotesDesc.stringify(
        {
          notes: ctx.endnotes.notes,
          separator: ctx.endnotes.separator,
          continuationSeparator: ctx.endnotes.continuationSeparator,
          continuationNotice: ctx.endnotes.continuationNotice,
        },
        endnoteCtx,
      ) ?? "");
    const endnoteRelCount = ctx.endnotes.relationships.nextRelationshipId;
    const resolved = resolvePartMedia(endnoteXmlData, ctx, endnoteRelCount, endnotesSource);
    registerPartMedia(ctx.endnotes.relationships, ctx, resolved, endnotesSource);
    const endnoteXml = resolvePartCharts(
      resolved.xml,
      ctx,
      ctx.endnotes.relationships,
      resolved.embeddingOffset + resolved.embeddingRefs.length,
      endnotesSource,
    );
    claimPartPassthroughRelationships(ctx.endnotes.relationships, ctx, endnotesSource);
    result.Endnotes = {
      data: replaceNumberingPlaceholders(endnoteXml, ctx.numbering.concreteNumbering),
      path: "word/endnotes.xml",
    };
    if (ctx.endnotes.relationships.relationshipCount > 0) {
      result.EndnotesRelationships = {
        data: XML_DECL + ctx.endnotes.relationships.serialize(),
        path: "word/_rels/endnotes.xml.rels",
      };
    }
  }

  return result;
}
