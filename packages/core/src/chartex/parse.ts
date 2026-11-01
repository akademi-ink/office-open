/**
 * chartEx model — the typed chart behind a `cx:chartSpace` payload.
 *
 * The schema ([MS-ODRAWXML] §5.22) separates the data from the plot: a
 * `cx:chartData` block list holds the cached string/numeric dimensions, and
 * every `cx:series` in `cx:plotAreaRegion` points at one block through
 * `cx:dataId`. A series' family is not an element name but the `layoutId`
 * attribute, so a `clusteredColumn` series and a `waterfall` series share the
 * same XML shape.
 *
 * Only the families the classic chart painter can draw are modelled here —
 * bar/column, line, area, pie/doughnut, and scatter. A payload built from
 * anything else (waterfall, treemap, funnel, boxWhisker, …), a dashboard
 * instead of a chart, or a chart with no cached values stays unmodelled: the
 * parser returns undefined and the caller keeps the opaque placeholder.
 *
 * Strings arrive as `cx:pt` element content and `ptCount`/`formatCode` as
 * attributes on `cx:lvl` (the opposite of the classic `c:numCache` shape), so
 * the walk below reads element text, never a `c:v` child.
 *
 * @module
 */

import { attr, attrBool, attrNum, parse, textOf } from "@office-open/xml";
import type { Element } from "@office-open/xml";

import type { ChartSpaceOptions, ChartSeriesData, ScatterSeriesData } from "../chart";

/** Chart families the classic painter renders from a chartEx payload. */
export type ChartExChartType = "bar" | "column" | "line" | "area" | "pie" | "doughnut" | "scatter";

/** Series grouping (`layoutId` tokens carrying stacked/percentStacked). */
export type ChartExGrouping = "clustered" | "stacked" | "percentStacked" | "standard";

/** One plotted series, resolved onto its data block's caches. */
export interface ChartExSeries {
  /** Series name (`cx:tx` text); absent when the source carries only a formula. */
  name?: string;
  /** Category labels aligned with `values`. */
  categories: readonly string[];
  /** Numeric values (`val` for cartesian/pie, `y` for scatter). */
  values: readonly number[];
  /** Scatter x values (`cx:numDim type="x"`); absent on non-scatter charts. */
  xValues?: readonly number[];
  /**
   * Series color, hex RRGGBB for `a:srgbClr` and the scheme token
   * (`accent1`, …) for `a:schemeClr` — the renderer resolves tokens against
   * its palette. Absent when the series carries no solid fill.
   */
  color?: string;
  /** Document-order index among the *visible* series. */
  index: number;
  /** Source `cx:data/@id` the series draws from. */
  dataId?: number;
}

/** One axis of the plot area. */
export interface ChartExAxis {
  /** `cx:axis/@id` — an internal handle the series `cx:axisId`s refer to. */
  id?: number;
  kind: "category" | "value";
  /** Axis title text (`cx:title`), absent when the axis has none. */
  title?: string;
  min?: number;
  max?: number;
  majorUnit?: number;
  /** `cx:numFmt/@formatCode`. */
  numberFormat?: string;
  /** `cx:majorGridlines` present. */
  gridlines?: boolean;
  /** `cx:axis/@hidden` — the axis is not painted. */
  hidden?: boolean;
}

/** Legend placement (`cx:legend/@pos`, ST_SidePos). */
export interface ChartExLegend {
  position: "left" | "top" | "right" | "bottom";
  overlay?: boolean;
}

/** A chartEx payload the renderer can draw. */
export interface ChartExChart {
  /** Family resolved from the series `layoutId`. */
  type: ChartExChartType;
  /** The `layoutId` token the family was resolved from (diagnostics). */
  layoutId: string;
  /** Chart title text (`cx:title`), absent when there is none. */
  title?: string;
  /** Category labels shared by the series (first series' cache). */
  categories: readonly string[];
  series: readonly ChartExSeries[];
  axes: readonly ChartExAxis[];
  /** Absent when the chart carries no `cx:legend` (the source hides it). */
  legend?: ChartExLegend;
  grouping: ChartExGrouping;
}

// ── layout family resolution ──

/** Tokens whose series the classic painter cannot draw — kept on the
 *  placeholder path. The spec's ST_SeriesLayout is the new-chart set; the
 *  classic-family tokens are accepted defensively because producers outside
 *  Excel's writer also populate chartEx parts. */
const UNSUPPORTED_LAYOUTS: ReadonlySet<string> = new Set([
  "boxwhisker",
  "funnel",
  "histogram",
  "paretoline",
  "regionmap",
  "sunburst",
  "treemap",
  "waterfall",
]);

/** Resolve a `cx:series/@layoutId` token to a paintable family. */
export function chartExTypeOfLayout(layoutId: string | undefined): ChartExChartType | undefined {
  const token = layoutId?.trim().toLowerCase() ?? "";
  if (!token || UNSUPPORTED_LAYOUTS.has(token)) return undefined;
  if (token.includes("doughnut") || token.includes("donut")) return "doughnut";
  if (token.includes("pie")) return "pie";
  if (token.includes("scatter")) return "scatter";
  // Families outside the requested scope: a bubble/radar/stock/surface token
  // must not fall through to a bar/column/line/area guess.
  if (
    token.includes("bubble") ||
    token.includes("radar") ||
    token.includes("stock") ||
    token.includes("surface")
  ) {
    return undefined;
  }
  if (token.includes("area")) return "area";
  if (token.includes("line")) return "line";
  if (token.includes("bar")) return "bar";
  if (token.includes("column")) return "column";
  return undefined;
}

function chartExGroupingOf(layoutId: string): ChartExGrouping {
  const token = layoutId.trim().toLowerCase();
  if (token.includes("percentstacked")) return "percentStacked";
  if (token.includes("stacked")) return "stacked";
  if (token.includes("bar") || token.includes("column")) return "clustered";
  return "standard";
}

// ── XML helpers ──

/** Local part of a qualified XML name (`cx:series` → `series`). */
function localNameOf(name: string | undefined): string {
  if (!name) return "";
  const colon = name.indexOf(":");
  return colon < 0 ? name : name.slice(colon + 1);
}

function findLocal(parent: Element | undefined, name: string): Element | undefined {
  return parent?.elements?.find((el) => localNameOf(el.name) === name);
}

function childrenLocal(parent: Element | undefined, name: string): Element[] {
  return parent?.elements?.filter((el) => localNameOf(el.name) === name) ?? [];
}

/** Element content as a trimmed string (simple-content nodes: `cx:v`, `cx:f`). */
function contentOf(el: Element | undefined): string | undefined {
  if (!el) return undefined;
  const text = textOf(el).trim();
  return text.length > 0 ? text : undefined;
}

/** All `a:t` runs below `el`, in document order. */
function textRunsOf(el: Element | undefined): string[] {
  const runs: string[] = [];
  const walk = (node: Element): void => {
    if (localNameOf(node.name) === "t") {
      const text = textOf(node);
      if (text) runs.push(text);
      return;
    }
    for (const child of node.elements ?? []) walk(child);
  };
  if (el) walk(el);
  return runs;
}

/** Title/axis-title text: `cx:tx` → `cx:txData` cached value, else the
 *  `cx:rich` DrawingML text body. */
function titleTextOf(container: Element | undefined): string | undefined {
  const tx = findLocal(container, "tx");
  if (!tx) return undefined;
  const cached = contentOf(findLocal(findLocal(tx, "txData"), "v"));
  if (cached) return cached;
  const rich = findLocal(tx, "rich");
  if (!rich) return undefined;
  const paragraphs = childrenLocal(rich, "p");
  const lines = paragraphs.map((p) => textRunsOf(p).join("").trim());
  const text = lines.filter((line) => line.length > 0).join(" ");
  return text.length > 0 ? text : undefined;
}

/** `cx:legend/@pos` (ST_SidePos l/t/r/b) → semantic side. */
function legendPositionOf(pos: string | undefined): ChartExLegend["position"] {
  switch (pos?.trim().toLowerCase()) {
    case "l":
    case "left":
      return "left";
    case "t":
    case "top":
      return "top";
    case "b":
    case "bottom":
      return "bottom";
    default:
      // The schema default is "r".
      return "right";
  }
}

// ── cached data blocks ──

/** One `cx:data` block, resolved onto plain arrays. */
interface DataBlock {
  id: number;
  /** Leaf (`lvl[0]`) string dimension — categories. */
  categories?: string[];
  /** `numDim` caches by dimension type. */
  val?: number[];
  x?: number[];
  y?: number[];
}

/** `cx:lvl` point count bound — the archive guard already bounds XML nodes;
 *  cap the materialized array anyway so a bogus `ptCount` cannot allocate. */
const MAX_POINTS = 10_000;

/** One `cx:lvl` element → its sparse points materialized by `idx`. */
function pointsOf(
  level: Element | undefined,
  numeric: boolean,
): Array<string | number | undefined> {
  if (!level) return [];
  const declared = attrNum(level, "ptCount");
  const points = childrenLocal(level, "pt");
  let size = declared !== undefined ? Math.min(Math.max(declared, 0), MAX_POINTS) : 0;
  for (const point of points) {
    const idx = attrNum(point, "idx");
    if (idx !== undefined && idx >= 0) size = Math.max(size, Math.min(idx + 1, MAX_POINTS));
  }
  const out = Array.from<string | number | undefined>({ length: size });
  for (const point of points) {
    const idx = attrNum(point, "idx");
    if (idx === undefined || idx < 0 || idx >= size) continue;
    const raw = textOf(point).trim();
    if (numeric) {
      const value = Number(raw);
      if (Number.isFinite(value)) out[idx] = value;
    } else {
      out[idx] = raw;
    }
  }
  return out;
}

function stringPointsOf(level: Element): string[] {
  return pointsOf(level, false).map((value) => (value === undefined ? "" : String(value)));
}

function numericPointsOf(level: Element): number[] {
  return pointsOf(level, true).map((value) => (typeof value === "number" ? value : 0));
}

function dataBlocksOf(space: Element | undefined): Map<number, DataBlock> {
  const blocks = new Map<number, DataBlock>();
  for (const data of childrenLocal(findLocal(space, "chartData"), "data")) {
    const id = attrNum(data, "id");
    if (id === undefined) continue;
    const block: DataBlock = { id };
    for (const dim of data.elements ?? []) {
      const name = localNameOf(dim.name);
      const type = (attr(dim, "type") ?? "").trim().toLowerCase();
      if (name === "strDim") {
        // Levels run leaf-first; the leaf is the label the renderer shows.
        const level = childrenLocal(dim, "lvl")[0];
        if (type === "cat" && level) block.categories = stringPointsOf(level);
      } else if (name === "numDim") {
        const level = childrenLocal(dim, "lvl")[0];
        if (!level) continue;
        const values = numericPointsOf(level);
        if (type === "val") block.val = values;
        else if (type === "x") block.x = values;
        else if (type === "y") block.y = values;
      }
    }
    blocks.set(id, block);
  }
  return blocks;
}

/** Series fill from `cx:series` → `cx:spPr` → `a:solidFill` (also tolerated
 *  under `cx:layoutPr`, where some producers park the formatting). A scheme
 *  token stays verbatim — the renderer resolves it against the active theme,
 *  exactly like a classic `c:schemeClr`. */
function seriesColorOf(series: Element): string | undefined {
  const spPr = findLocal(series, "spPr") ?? findLocal(findLocal(series, "layoutPr"), "spPr");
  const fill = findLocal(spPr, "solidFill");
  if (!fill) return undefined;
  const color = findLocal(fill, "srgbClr") ?? findLocal(fill, "schemeClr");
  const value = attr(color, "val")?.trim();
  if (!value) return undefined;
  return /^[0-9a-f]{6}$/i.test(value) ? value.toUpperCase() : value.toLowerCase();
}

// ── parse ──

/**
 * Parse a chartEx part (`cx:chartSpace`) into the typed model.
 *
 * Returns undefined when the payload is not a chartSpace/chart, when no
 * series resolves to a paintable family, when visible series disagree on the
 * family (a combo the classic painter cannot draw), or when no cached numeric
 * value exists — in every one of those cases the caller keeps the opaque
 * placeholder instead of painting a misleading frame.
 *
 * @param input chartEx part XML (string or UTF-8 bytes)
 */
export function parseChartExChartSpace(
  input: string | Uint8Array | undefined | null,
): ChartExChart | undefined {
  if (input === undefined || input === null) return undefined;
  let xml: string;
  if (typeof input === "string") {
    xml = input;
  } else {
    try {
      xml = new TextDecoder("utf-8").decode(input);
    } catch {
      return undefined;
    }
  }
  if (!xml.includes("<")) return undefined;

  let root: Element | undefined;
  try {
    root = parse(xml).elements?.[0];
  } catch {
    return undefined;
  }
  if (!root) return undefined;

  const rootName = localNameOf(root.name).toLowerCase();
  const space = rootName === "chartspace" ? root : undefined;
  const chartEl = rootName === "chart" ? root : findLocal(space, "chart");
  if (!chartEl) return undefined; // a dashboard part is not a chart

  const blocks = dataBlocksOf(space);
  const plotArea = findLocal(chartEl, "plotArea");
  const rawSeries = childrenLocal(findLocal(plotArea, "plotAreaRegion"), "series");

  let type: ChartExChartType | undefined;
  let layoutId = "";
  let grouping: ChartExGrouping = "standard";
  const series: ChartExSeries[] = [];

  for (let index = 0; index < rawSeries.length; index++) {
    const el = rawSeries[index]!;
    if (attrBool(el, "hidden") === true) continue;
    const token = attr(el, "layoutId") ?? "";
    const seriesType = chartExTypeOfLayout(token);
    // A pareto's computed line shares its column series' data block; it has
    // no values of its own and is skipped, not treated as a chart family.
    if (!seriesType) continue;
    if (type !== undefined && seriesType !== type) return undefined;
    if (type === undefined) {
      type = seriesType;
      layoutId = token;
      grouping = chartExGroupingOf(token);
    }

    const dataId = attrNum(findLocal(el, "dataId"), "val");
    const block =
      (dataId !== undefined ? blocks.get(dataId) : undefined) ??
      blocks.get(index) ??
      [...blocks.values()][series.length];
    if (!block) continue;

    const scatter = type === "scatter";
    const values = scatter ? (block.y ?? block.val ?? []) : (block.val ?? block.y ?? []);
    if (values.length === 0) continue;
    const categories = block.categories ?? [];
    const xValues = scatter ? (block.x ?? []) : undefined;
    const name = titleTextOf(el);
    const color = seriesColorOf(el);
    series.push({
      ...(name ? { name } : {}),
      categories,
      values: [...values],
      ...(xValues ? { xValues } : {}),
      ...(color ? { color } : {}),
      index: series.length,
      ...(dataId !== undefined ? { dataId } : {}),
    });
  }

  if (!type || series.length === 0) return undefined;
  if (!series.some((entry) => entry.values.length > 0)) return undefined;
  const categories = series.find((entry) => entry.categories.length > 0)?.categories ?? [];

  const axes: ChartExAxis[] = [];
  for (const el of childrenLocal(plotArea, "axis")) {
    const categoryScaling = findLocal(el, "catScaling");
    const valueScaling = findLocal(el, "valScaling");
    const kind = categoryScaling ? "category" : valueScaling ? "value" : undefined;
    if (!kind) continue;
    const scaling = categoryScaling ?? valueScaling;
    const axis: ChartExAxis = { kind };
    const id = attrNum(el, "id");
    if (id !== undefined) axis.id = id;
    const title = titleTextOf(findLocal(el, "title"));
    if (title) axis.title = title;
    if (kind === "value") {
      const min = attrNum(scaling, "min");
      const max = attrNum(scaling, "max");
      const majorUnit = attrNum(scaling, "majorUnit");
      if (min !== undefined) axis.min = min;
      if (max !== undefined) axis.max = max;
      if (majorUnit !== undefined && majorUnit > 0) axis.majorUnit = majorUnit;
    }
    const formatCode = attr(findLocal(el, "numFmt"), "formatCode")?.trim();
    if (formatCode) axis.numberFormat = formatCode;
    if (findLocal(el, "majorGridlines")) axis.gridlines = true;
    if (attrBool(el, "hidden") === true) axis.hidden = true;
    axes.push(axis);
  }

  const legendEl = findLocal(chartEl, "legend");
  const title = titleTextOf(findLocal(chartEl, "title"));

  return {
    type,
    layoutId,
    ...(title ? { title } : {}),
    categories,
    series,
    axes,
    ...(legendEl
      ? {
          legend: {
            position: legendPositionOf(attr(legendEl, "pos")),
            ...(attrBool(legendEl, "overlay") === true ? { overlay: true } : {}),
          },
        }
      : {}),
    grouping,
  };
}

// ── renderer bridge ──

/** Fill record the classic painter reads (`shapeProperties.fill`). */
function solidFillOf(color: string): { type: "solid"; color: { value: string } } {
  return { type: "solid", color: { value: color } };
}

/**
 * Project a parsed chartEx chart onto the classic `ChartSpaceOptions` shape.
 *
 * The classic painter consumes that model structurally (type, categories,
 * series, axes, title, legend), so the bridge is a pure data mapping — no
 * chartEx knowledge leaks into the painter. Pie/doughnut series drop their
 * series fill on purpose: the painter colors each slice from the point index,
 * and a series fill would paint one flat disc.
 *
 * @param chart parsed chartEx model
 */
export function chartExToChartSpace(chart: ChartExChart): ChartSpaceOptions {
  const pie = chart.type === "pie" || chart.type === "doughnut";
  const scatter = chart.type === "scatter";

  const valueSeries: ChartSeriesData[] = [];
  const scatterSeries: ScatterSeriesData[] = [];
  for (const entry of chart.series) {
    const shapeProperties = entry.color && !pie ? { fill: solidFillOf(entry.color) } : undefined;
    const common = {
      ...(entry.name ? { name: entry.name } : {}),
      ...(shapeProperties ? { shapeProperties } : {}),
    };
    if (scatter) {
      scatterSeries.push({
        ...common,
        xValues: [...(entry.xValues ?? [])],
        yValues: [...entry.values],
      });
    } else {
      valueSeries.push({ ...common, values: [...entry.values] });
    }
  }

  const axes = chart.axes.map((axis, index) => {
    const projected: NonNullable<ChartSpaceOptions["axes"]>[number] = {
      kind: axis.kind,
      ...(axis.id !== undefined ? { id: axis.id } : {}),
      ...(axis.title ? { title: axis.title } : {}),
      ...(axis.numberFormat ? { numberFormat: axis.numberFormat } : {}),
      ...(axis.min !== undefined || axis.max !== undefined
        ? {
            scaling: {
              ...(axis.min !== undefined ? { min: axis.min } : {}),
              ...(axis.max !== undefined ? { max: axis.max } : {}),
            },
          }
        : {}),
      ...(axis.majorUnit !== undefined ? { majorUnit: axis.majorUnit } : {}),
      ...(axis.gridlines ? { majorGridlines: true } : {}),
      ...(axis.hidden ? { delete: true } : {}),
    };
    // A chartEx axis carries no position; the scatter painter reads it from
    // the first value axis (horizontal) and the second (vertical).
    if (scatter && axis.kind === "value") {
      projected.position =
        chart.axes.slice(0, index).filter((entry) => entry.kind === "value").length === 0
          ? "bottom"
          : "left";
    }
    return projected;
  });

  return {
    type: chart.type,
    ...(chart.title ? { title: chart.title } : {}),
    categories: [...chart.categories],
    series: scatter ? scatterSeries : valueSeries,
    ...(chart.grouping !== "standard" ? { grouping: chart.grouping } : {}),
    showLegend: chart.legend !== undefined,
    ...(chart.legend ? { legendPosition: chart.legend.position } : {}),
    ...(axes.length > 0 ? { axes } : {}),
  };
}
