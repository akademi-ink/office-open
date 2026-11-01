/**
 * Relationship wiring shared by the docx compile phases: the XML declaration,
 * the OLE/PACKAGE embedding relationship resolution, and the per-part media +
 * embedding placeholder bridge.
 *
 * @module
 */

import {
  RELATIONSHIP_TYPES,
  type RelationshipType,
  type Relationships,
  findAndReplaceImagePlaceholders,
  replaceAllPlaceholders,
} from "@office-open/core";
import { OOXML_XML_DECLARATION } from "@office-open/xml";
import type { EmbeddingCollection } from "@shared/embeddings/embeddings";

import type { DocxWriteContext } from "../context";

/** XML declaration prepended to every OOXML part. */
export const XML_DECL = OOXML_XML_DECLARATION;

/** Relationship type for OLE embedding parts (word|ppt/embeddings/*). */
const OLE_OBJECT_RELATIONSHIP = RELATIONSHIP_TYPES.oleObject;

/** Relationship type for native-format embedding parts (embedded xlsx/docx). */
export const PACKAGE_RELATIONSHIP = RELATIONSHIP_TYPES.package;

/** Re-emit the relationship type the source used for an embedding part — a
 *  native OPC package (xlsx/docx) stays a package rel, an OLE compound binary
 *  stays an oleObject rel. */
export const embeddingRelationship = (
  embeddings: EmbeddingCollection,
  fileName: string,
): RelationshipType =>
  embeddings.array.find((e) => e.fileName === fileName)?.relationshipType === "package"
    ? PACKAGE_RELATIONSHIP
    : OLE_OBJECT_RELATIONSHIP;

/** Resolved media/embedding placeholders for one part, with the offsets its
 *  relationship registrations must use (ids are per-part numbering). */
export interface PartMediaResolution {
  xml: string;
  relCount: number;
  embeddingOffset: number;
  mediaRefs: { fileName: string }[];
  embeddingRefs: { fileName: string }[];
}

/**
 * Resolve a part's {fileName} media placeholders against the package media
 * store, then its embedding placeholders chained past the media
 * relationships (same ordering as the document part — headers/footers and
 * notes can carry w:object runs of their own).
 */
export function resolvePartMedia(
  xml: string,
  ctx: DocxWriteContext,
  relCount: number,
  source?: string,
): PartMediaResolution {
  const media = findAndReplaceImagePlaceholders(
    xml,
    ctx.media.array,
    relCount,
    "rId",
    partSourceRidMap(ctx, source, "media"),
  );
  const embeddingOffset = relCount + media.referenced.length;
  const embeddings = findAndReplaceImagePlaceholders(
    media.xml,
    ctx.embeddings.array,
    embeddingOffset,
    "rId",
    partSourceRidMap(ctx, source, "embeddings"),
  );
  return {
    xml: embeddings.xml,
    relCount,
    embeddingOffset,
    mediaRefs: media.referenced,
    embeddingRefs: embeddings.referenced,
  };
}

/** Register the resolved media + embedding relationships on the part's rels.
 *
 *  When `source` names the part's package path, a relationship the source
 *  carried for the same target is registered at its source rId — the part's
 *  XML references that id (see {@link resolvePartMedia}), and reusing it keeps
 *  parse→generate a fixed point: a fresh offset-derived id would sit above the
 *  reserved source-id watermark and climb by one on every second generation. */
export function registerPartMedia(
  rels: Relationships,
  ctx: DocxWriteContext,
  resolved: PartMediaResolution,
  source?: string,
): void {
  for (const [i, ref] of resolved.mediaRefs.entries()) {
    const target = `media/${ref.fileName}`;
    rels.addRelationship(
      sourceRidFor(ctx, source, RELATIONSHIP_TYPES.image, target) ?? resolved.relCount + i,
      RELATIONSHIP_TYPES.image,
      target,
    );
  }
  for (const [i, ref] of resolved.embeddingRefs.entries()) {
    const target = `embeddings/${ref.fileName}`;
    const type = embeddingRelationship(ctx.embeddings, ref.fileName);
    rels.addRelationship(
      sourceRidFor(ctx, source, type, target) ?? resolved.embeddingOffset + i,
      type,
      target,
    );
  }
}

/**
 * Look up the source rId for a rebuilt part's relationship by kind+target.
 * Returns the numeric id when the source carried the same rel, undefined
 * otherwise. The part's XML references the source rId verbatim on round-trip,
 * so the structured compiler must emit the rel at that exact id — otherwise
 * `r:id="rId4"` would point at a different rel and Word refuses to open.
 *
 * Also the allocation base: an auto-allocated id would sit above the
 * source-id reservation and drift upward on every parse→generate cycle
 * (gen2 footer1 rId2 → gen3 rId3), so model-owned rels reuse the source id
 * instead of taking `nextRelationshipId`.
 */
export function sourceRidFor(
  ctx: DocxWriteContext,
  source: string | undefined,
  relationshipType: string,
  target: string,
): number | undefined {
  if (source === undefined) return undefined;
  for (const rel of ctx._options.passthroughRelationships ?? []) {
    if (rel.source !== source) continue;
    if (rel.relationshipType !== relationshipType || rel.target !== target) continue;
    const m = /^rId(\d+)$/.exec(rel.rId);
    if (m) return Number(m[1]);
  }
  return undefined;
}

/**
 * Build a {fileName → source rId} map for the `media/` or `embeddings/` rels
 * of one rebuilt part. {@link findAndReplaceImagePlaceholders} consults it to
 * substitute the source rId in the part XML. Placeholder keys use the
 * media-collection fileName; a dedup-renamed file misses the map and falls
 * through to the offset id on both the XML and the .rels side.
 */
export function partSourceRidMap(
  ctx: DocxWriteContext,
  source: string | undefined,
  dir: "media" | "embeddings",
): ReadonlyMap<string, string> | undefined {
  if (source === undefined) return undefined;
  const prefix = `${dir}/`;
  const map = new Map<string, string>();
  for (const rel of ctx._options.passthroughRelationships ?? []) {
    if (rel.source !== source || !rel.target.startsWith(prefix)) continue;
    map.set(rel.target.slice(prefix.length), rel.rId);
  }
  return map;
}

/**
 * Pre-claim a rebuilt sub-part's source ids before the model allocates its own
 * relationships (same contract as the document part's constructor reservation):
 * auto-allocated ids then start above the source id space, so a later
 * {@link claimPartPassthroughRelationships} re-emits every source relationship
 * at its exact source id instead of renumbering — renumbering would dangle the
 * verbatim `r:id` the model kept for it.
 */
export function reservePartSourceRids(
  rels: Relationships,
  ctx: DocxWriteContext,
  partName: string,
): void {
  rels.reserveSourceRids(partName, ctx._options.passthroughRelationships ?? []);
}

/**
 * Re-register the source relationships of a rebuilt sub-part (header, footer,
 * comments, footnotes, endnotes) that survive as passthrough parts.
 *
 * Those parts are re-emitted from the model, so their `.rels` are rebuilt too.
 * A source relationship whose target still travels verbatim (a Word ink
 * content part behind `w:contentPart`, an OLE package, a glossary) has no
 * entry in the rebuilt `.rels`, while the model kept the source `r:id` in
 * `contentPartRId`/`rawXml` — the generated package then carries a dangling
 * reference Word cannot resolve. Claim the captured source rels at their
 * source id; relationships the model already registered under the same
 * kind+target are skipped by {@link Relationships.claimSourceRel}.
 *
 * The document part runs the equivalent pass in
 * {@link DocxWriteContext.addPassthroughDocumentRelationships}.
 */
export function claimPartPassthroughRelationships(
  rels: Relationships,
  ctx: DocxWriteContext,
  partName: string,
): void {
  for (const rel of ctx._options.passthroughRelationships ?? []) {
    if (rel.source !== partName) continue;
    rels.claimSourceRel(rel);
  }
}

/** Matches any chart/SmartArt placeholder prefix in a part's XML. */
const CHART_SMARTART_PLACEHOLDER = /\{(chart:|smartart(?:-lo|-qs|-cs)?:)/;

/**
 * Resolve a part's {chart:key} / {smartart*:key} placeholders and register the
 * referenced parts on the part's own rels — headers/footers and notes can
 * carry chart and SmartArt drawings of their own. The document part resolves
 * the full collections (compile/document.ts); here only keys this part
 * actually references get relationships, at ids continuing past the part's
 * media/embedding block. Targets keep the package-wide collection order
 * (charts/chartN.xml, diagrams/{data,layout,quickStyle,colors}N.xml), so
 * several parts referencing one chart all point at the same file.
 */
export function resolvePartCharts(
  xml: string,
  ctx: DocxWriteContext,
  rels: Relationships,
  relCount: number,
  source?: string,
): string {
  if (!CHART_SMARTART_PLACEHOLDER.test(xml)) return xml;
  const entries: Array<{ prefix?: string; key: string; value: string }> = [];
  /** Source rId when the part carried the rel, else the offset-derived id. */
  const allocate = (type: RelationshipType, target: string, fallback: number): number =>
    sourceRidFor(ctx, source, type, target) ?? fallback;
  const referencedCharts = ctx.charts.array.filter((c) => xml.includes(`{chart:${c.key}}`));
  referencedCharts.forEach((chart, i) => {
    const target = `charts/chart${ctx.charts.array.indexOf(chart) + 1}.xml`;
    const id = allocate(RELATIONSHIP_TYPES.chart, target, relCount + i);
    entries.push({ prefix: "chart:", key: chart.key, value: `rId${id}` });
    rels.addRelationship(id, RELATIONSHIP_TYPES.chart, target);
  });
  const referencedSmartArts = ctx.smartArts.array.filter((s) =>
    ["", "-lo", "-qs", "-cs"].some((suffix) => xml.includes(`{smartart${suffix}:${s.key}}`)),
  );
  const chartCount = referencedCharts.length;
  const smartArtCount = referencedSmartArts.length;
  const base = relCount + chartCount;
  const loOffset = base + smartArtCount;
  const qsOffset = loOffset + smartArtCount;
  const csOffset = qsOffset + smartArtCount;
  const drawingOffset = csOffset + smartArtCount;
  referencedSmartArts.forEach((smartArt, i) => {
    const fileIndex = ctx.smartArts.array.indexOf(smartArt) + 1;
    const relsByPrefix: Array<[string, number, RelationshipType, string]> = [
      ["smartart:", base, RELATIONSHIP_TYPES.diagramData, "data"],
      ["smartart-lo:", loOffset, RELATIONSHIP_TYPES.diagramLayout, "layout"],
      ["smartart-qs:", qsOffset, RELATIONSHIP_TYPES.diagramQuickStyle, "quickStyle"],
      ["smartart-cs:", csOffset, RELATIONSHIP_TYPES.diagramColors, "colors"],
    ];
    for (const [prefix, offset, type, file] of relsByPrefix) {
      const target = `diagrams/${file}${fileIndex}.xml`;
      const id = allocate(type, target, offset + i);
      entries.push({ prefix, key: smartArt.key, value: `rId${id}` });
      rels.addRelationship(id, type, target);
    }
    // The drawing part is an Office render cache, present only when the source
    // carried it — Word never emits it for a fresh SmartArt.
    if (smartArt.raw?.drawing !== undefined) {
      const target = `diagrams/drawing${fileIndex}.xml`;
      rels.addRelationship(
        allocate(RELATIONSHIP_TYPES.diagramDrawingMs, target, drawingOffset + i),
        RELATIONSHIP_TYPES.diagramDrawingMs,
        target,
      );
    }
  });
  return replaceAllPlaceholders(xml, entries);
}
