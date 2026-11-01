/**
 * chartEx — Office 2014+ "modern chart" package surface (PKG-12).
 *
 * Word's modern charts live in a part type of their own: a `cx:chartSpace`
 * document under `word/charts/chartEx{N}.xml` with content type
 * `application/vnd.ms-office.chartex+xml`, reached from `word/document.xml`
 * through a `chartEx` relationship rather than the classic `chart` one. The
 * payload is not DrawingML chart schema — it is the `cx` namespace family
 * (`…/office/drawing/2014/chartex` plus the dated 2015/2016 revisions) — so the
 * classic `c:chartSpace` model cannot represent it.
 *
 * The part still round-trips byte-exact as an opaque part and the referencing
 * drawing as inline passthrough XML. Two paths sit on top of that preserved
 * payload:
 *
 * - {@link parseChartExChartSpace} models the common families a canvas can
 *   draw (bar/column, line, area, pie/doughnut, scatter) into a typed chart;
 *   {@link chartExToChartSpace} projects that onto the classic chart model the
 *   existing chart painter consumes.
 * - Anything outside those families (waterfall, treemap, funnel, box & whisker,
 *   …), a dashboard part, or a payload with no cached values returns
 *   `undefined` from the parser and stays on the placeholder path below.
 *
 * This module also owns the vocabulary that decision needs — the part path,
 * content type, relationship type, and namespace detection — so the package
 * validator recognizes chartEx parts and the docx layer can detect the drawing
 * without re-deriving the URIs.
 *
 * @module
 */

import { attrNum, findFirst, parse } from "@office-open/xml";
import type { Element } from "@office-open/xml";

export * from "./parse";

/** The chartEx namespace Word writes for the chart payload (`cx:` prefix). */
export const CHART_EX_NAMESPACE = "http://schemas.microsoft.com/office/drawing/2014/chartex";

/** OPC content type of a chartEx part (`word/charts/chartEx1.xml`). */
export const CHART_EX_CONTENT_TYPE = "application/vnd.ms-office.chartex+xml";

/** Relationship type that reaches a chartEx part from its host part. */
export const CHART_EX_RELATIONSHIP_TYPE =
  "http://schemas.microsoft.com/office/2014/relationships/chartEx";

/** Directory chartEx parts live in, per format (docx is the one we emit). */
export const CHART_EX_DIRECTORY = "word/charts";

/**
 * ZIP path of the `index`-th chartEx part (`word/charts/chartEx1.xml`).
 *
 * @param index 1-based part index, as Word allocates it
 * @param directory directory override for other package formats
 */
export function chartExPartPath(index: number, directory = CHART_EX_DIRECTORY): string {
  return `${directory}/chartEx${index}.xml`;
}

/**
 * Is `path` a chartEx part? The basename is case-insensitive because producers
 * are inconsistent (`chartEx1.xml` in Word, `chartex1.xml` elsewhere).
 *
 * @param path ZIP path or part URI (a leading `/` is ignored)
 */
export function isChartExPartPath(path: string): boolean {
  return /(?:^|\/)charts\/chartex\d+\.xml$/i.test(path.replace(/^\/+/, ""));
}

/**
 * Is `type` the chartEx relationship type? Compares the URI only (case and
 * media-type parameters are irrelevant for relationships).
 *
 * @param type relationship `Type` URI
 */
export function isChartExRelationshipType(type: string): boolean {
  return type.trim().toLowerCase() === CHART_EX_RELATIONSHIP_TYPE.toLowerCase();
}

/**
 * Is `uri` a chartEx `a:graphicData/@uri`? Accepts the 2014 namespace and the
 * dated 2015/2016 revisions Word also emits.
 *
 * @param uri `a:graphicData` URI attribute
 */
export function isChartExGraphicDataUri(uri: string): boolean {
  return /^https?:\/\/schemas\.microsoft\.com\/office\/drawing\/\d{4}(?:\/\d+\/\d+)?\/chartex\/?$/i.test(
    uri.trim(),
  );
}

/** A chartEx part discovered in an unzipped package. */
export interface ChartExPart {
  /** ZIP path of the part. */
  path: string;
  /** OPC content type declared for the part. */
  contentType: string;
}

/**
 * Collect the chartEx parts of an unzipped package.
 *
 * @param paths ZIP paths of the package
 */
export function chartExParts(paths: Iterable<string>): ChartExPart[] {
  const found: ChartExPart[] = [];
  for (const path of paths) {
    if (isChartExPartPath(path)) found.push({ path, contentType: CHART_EX_CONTENT_TYPE });
  }
  return found;
}

/** One `{ rawXml: … }` passthrough child's `rawMedia` entry. */
export interface ChartExRawMediaEntry {
  fileName?: unknown;
  data?: unknown;
}

/**
 * Find the chartEx part a preserved drawing run carries in its `rawMedia`.
 *
 * The docx parser resolves a drawing's relationship refs (`cx:chart/@r:id`)
 * into `{fileName}` placeholders and attaches the referenced parts as
 * `rawMedia`; a chartEx payload therefore arrives with the drawing without
 * re-reading the package. The first entry whose name is a chartEx part wins.
 *
 * @param rawMedia the passthrough child's `rawMedia` array (untyped input)
 */
export function chartExPartOfRawMedia(
  rawMedia: unknown,
): { fileName: string; data: Uint8Array | string } | undefined {
  if (!Array.isArray(rawMedia)) return undefined;
  for (const entry of rawMedia) {
    if (typeof entry !== "object" || entry === null) continue;
    const record = entry as Record<string, unknown>;
    const fileName = typeof record.fileName === "string" ? record.fileName : undefined;
    // rawMedia names are basenames (`chartEx1.xml`); accept a full part path
    // too, for callers that keep the package-relative name.
    if (!fileName || !(isChartExPartPath(fileName) || /^chartex\d+\.xml$/i.test(fileName)))
      continue;
    const data = record.data;
    if (typeof data === "string" && data.length > 0) return { fileName, data };
    if (data instanceof Uint8Array) return { fileName, data };
    if (data instanceof ArrayBuffer) return { fileName, data: new Uint8Array(data) };
  }
  return undefined;
}

// ── Renderer placeholder ──

/**
 * Text a renderer paints inside the frame of a preserved chartEx drawing. The
 * wording is part of the contract: it tells the reader the chart exists and is
 * deliberately not rendered, instead of showing a blank box.
 */
export const CHART_EX_PLACEHOLDER_LABEL = "modern chart (not rendered)";

/** Width/height pair in the drawing's own unit (EMU for `wp:extent`). */
export interface ChartExExtent {
  /** Drawing width in EMU (`wp:extent/@cx`). */
  width: number;
  /** Drawing height in EMU (`wp:extent/@cy`). */
  height: number;
}

/** Paintable placeholder geometry for a preserved chartEx drawing. */
export interface ChartExPlaceholderGeometry {
  /** Frame box in the same unit as the source extent. */
  frame: { x: number; y: number; width: number; height: number };
  /** Label anchor: horizontal center of the frame. */
  labelX: number;
  /** Label anchor: vertical center of the frame. */
  labelY: number;
  /** Text to paint at the label anchor. */
  label: string;
}

/** Placeholder options. */
export interface ChartExPlaceholderOptions {
  /** Inset from the drawing extent to the painted frame, same unit as the extent. */
  inset?: number;
  /**
   * Smallest frame dimension. A degenerate extent (0×0, or absent) still
   * yields a visible box instead of a zero-area blank.
   */
  minSize?: number;
}

/** True when a preserved drawing's XML references a chartEx payload. */
export function isChartExDrawingXml(xml: string): boolean {
  return /<cx:chart\b/.test(xml) || /graphicData[^>]*\buri="[^"]*\/chartex"/i.test(xml);
}

/**
 * Read the `wp:extent` of a preserved chartEx drawing. Returns undefined when
 * `xml` is not a chartEx drawing or carries no usable extent.
 *
 * @param xml verbatim `w:drawing` XML, as preserved on the drawing run
 */
export function chartExExtentFromDrawingXml(xml: string): ChartExExtent | undefined {
  if (!isChartExDrawingXml(xml)) return undefined;
  let extent: Element | undefined;
  try {
    extent = findFirst(parse(xml).elements?.[0], "wp:extent");
  } catch {
    return undefined;
  }
  const width = extent ? attrNum(extent, "cx") : undefined;
  const height = extent ? attrNum(extent, "cy") : undefined;
  if (width === undefined || height === undefined || width <= 0 || height <= 0) return undefined;
  return { width, height };
}

/**
 * Placeholder geometry for a preserved chartEx drawing.
 *
 * The frame is the drawing extent inset by `inset` (default 0) and never
 * smaller than `minSize` (default 1) on either axis, so every call produces a
 * non-blank box — including a degenerate or missing extent. The label anchors
 * at the frame center; the caller owns font, stroke, and clipping.
 *
 * @param extent source drawing extent (EMU); undefined for a malformed source
 * @param options inset/minimum size in the extent's unit
 */
export function chartExPlaceholderGeometry(
  extent: ChartExExtent | undefined,
  options: ChartExPlaceholderOptions = {},
): ChartExPlaceholderGeometry {
  const requestedInset = options.inset;
  const inset =
    requestedInset !== undefined && Number.isFinite(requestedInset) && requestedInset > 0
      ? requestedInset
      : 0;
  const requestedMin = options.minSize;
  const minSize =
    requestedMin !== undefined && Number.isFinite(requestedMin) && requestedMin > 0
      ? requestedMin
      : 1;
  const sourceWidth = extent && Number.isFinite(extent.width) ? Math.max(0, extent.width) : 0;
  const sourceHeight = extent && Number.isFinite(extent.height) ? Math.max(0, extent.height) : 0;
  const width = Math.max(minSize, sourceWidth - inset * 2);
  const height = Math.max(minSize, sourceHeight - inset * 2);
  return {
    frame: { x: inset, y: inset, width, height },
    labelX: inset + width / 2,
    labelY: inset + height / 2,
    label: CHART_EX_PLACEHOLDER_LABEL,
  };
}

/**
 * Convenience for a renderer holding a preserved drawing's verbatim XML:
 * detect chartEx, read its extent, and build the placeholder. Undefined when
 * the XML is not a chartEx drawing.
 */
export function chartExPlaceholderOfDrawingXml(
  xml: string,
  options?: ChartExPlaceholderOptions,
): ChartExPlaceholderGeometry | undefined {
  if (!isChartExDrawingXml(xml)) return undefined;
  return chartExPlaceholderGeometry(chartExExtentFromDrawingXml(xml), options);
}
