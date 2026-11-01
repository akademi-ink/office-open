/**
 * chartex — Office 2014+ "modern chart" package surface (PKG-12).
 *
 * Word's modern charts live in a part type of their own: a `cx:chartSpace`
 * document under `word/charts/chartEx{N}.xml` with content type
 * `application/vnd.ms-office.chartex+xml`, reached from `word/document.xml`
 * through a `chartEx` relationship rather than the classic `chart` one. The
 * payload is not DrawingML chart schema — it is the `cx` namespace family
 * (`…/office/drawing/2014/chartex` plus the dated 2015/2016 revisions) — so the
 * classic `c:chartSpace` model cannot represent it.
 *
 * Decision (recorded in `scripts/standards/README.md`): **preserve, do not
 * model**. A chartEx part round-trips byte-exact as an opaque part and the
 * referencing drawing as inline passthrough XML; we do not parse `cx:chartSpace`
 * into an editable chart model, and no visual chart is reconstructed from it.
 * This module owns the vocabulary that decision needs — the part path, content
 * type, relationship type, and namespace detection — so the package validator
 * recognizes chartEx parts and the docx layer can detect the drawing without
 * re-deriving the URIs.
 *
 * @module
 */

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
