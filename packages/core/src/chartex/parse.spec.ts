import { describe, expect, it } from "vite-plus/test";

import { chartExPartOfRawMedia } from "./index";
import { chartExToChartSpace, chartExTypeOfLayout, parseChartExChartSpace } from "./parse";

const CX = "http://schemas.microsoft.com/office/drawing/2014/chartex";
const A = "http://schemas.openxmlformats.org/drawingml/2006/main";

/** A chartEx part with one data block and the supplied series/plot markup. */
function chartSpace(options: {
  series: string;
  data?: string;
  title?: string;
  legend?: string;
  axes?: string;
  prefix?: string;
}): string {
  const p = options.prefix ?? "cx";
  const data =
    options.data ??
    `<${p}:data id="0">` +
      `<${p}:strDim type="cat"><${p}:lvl ptCount="3">` +
      `<${p}:pt idx="0">East</${p}:pt><${p}:pt idx="1">West</${p}:pt>` +
      `<${p}:pt idx="2">North</${p}:pt></${p}:lvl></${p}:strDim>` +
      `<${p}:numDim type="val"><${p}:lvl ptCount="3" formatCode="General">` +
      `<${p}:pt idx="0">10</${p}:pt><${p}:pt idx="1">20</${p}:pt>` +
      `<${p}:pt idx="2">30</${p}:pt></${p}:lvl></${p}:numDim>` +
      `</${p}:data>`;
  const title = options.title
    ? `<${p}:title><${p}:tx><${p}:rich><a:bodyPr/><a:p><a:r><a:t>${options.title}</a:t></a:r></a:p></${p}:rich></${p}:tx></${p}:title>`
    : "";
  const axes =
    options.axes ??
    `<${p}:axis id="0"><${p}:catScaling gapWidth="0.5"/></${p}:axis>` +
      `<${p}:axis id="1"><${p}:valScaling min="0" max="40" majorUnit="10"/><${p}:majorGridlines/>` +
      `<${p}:numFmt formatCode="0"/></${p}:axis>`;
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<${p}:chartSpace xmlns:${p}="${CX}" xmlns:a="${A}">` +
    `<${p}:chartData>${data}</${p}:chartData>` +
    `<${p}:chart>${title}<${p}:plotArea><${p}:plotAreaRegion>${options.series}</${p}:plotAreaRegion>${axes}</${p}:plotArea>` +
    `${options.legend ?? ""}</${p}:chart></${p}:chartSpace>`
  );
}

const series = (layoutId: string, name = "2025", extra = ""): string =>
  `<cx:series layoutId="${layoutId}">` +
  `<cx:tx><cx:txData><cx:v>${name}</cx:v></cx:txData></cx:tx>` +
  `<cx:spPr><a:solidFill><a:srgbClr val="4472C4"/></a:solidFill></cx:spPr>` +
  `<cx:dataId val="0"/>${extra}</cx:series>`;

describe("parseChartExChartSpace", () => {
  it("models a clustered column chart: title, categories, series, axes, legend", () => {
    const chart = parseChartExChartSpace(
      chartSpace({
        series: series("clusteredColumn"),
        title: "Revenue",
        legend: '<cx:legend pos="b"/>',
      }),
    );
    expect(chart).toBeDefined();
    expect(chart!.type).toBe("column");
    expect(chart!.layoutId).toBe("clusteredColumn");
    expect(chart!.title).toBe("Revenue");
    expect(chart!.categories).toEqual(["East", "West", "North"]);
    expect(chart!.series).toEqual([
      {
        name: "2025",
        categories: ["East", "West", "North"],
        values: [10, 20, 30],
        color: "4472C4",
        index: 0,
        dataId: 0,
      },
    ]);
    expect(chart!.axes).toEqual([
      { kind: "category", id: 0 },
      { kind: "value", id: 1, min: 0, max: 40, majorUnit: 10, numberFormat: "0", gridlines: true },
    ]);
    expect(chart!.legend).toEqual({ position: "bottom" });
    expect(chart!.grouping).toBe("clustered");
  });

  it("resolves the family from the layoutId token across the supported set", () => {
    for (const [layoutId, type] of [
      ["clusteredBar", "bar"],
      ["bar", "bar"],
      ["stackedColumn", "column"],
      ["line", "line"],
      ["lineMarker", "line"],
      ["area", "area"],
      ["stackedArea", "area"],
      ["pie", "pie"],
      ["doughnut", "doughnut"],
      ["scatter", "scatter"],
    ] as const) {
      expect(chartExTypeOfLayout(layoutId)).toBe(type);
    }
  });

  it("reads a doughnut with its series fill preserved in the model", () => {
    const chart = parseChartExChartSpace(chartSpace({ series: series("doughnut", "Share") }));
    expect(chart!.type).toBe("doughnut");
    expect(chart!.series[0]!.color).toBe("4472C4");
  });

  it("reads scatter x/y dimensions and keeps them on the series", () => {
    const data =
      `<cx:data id="0">` +
      `<cx:numDim type="x"><cx:lvl ptCount="3">` +
      `<cx:pt idx="0">1</cx:pt><cx:pt idx="1">2</cx:pt><cx:pt idx="2">3</cx:pt>` +
      `</cx:lvl></cx:numDim>` +
      `<cx:numDim type="y"><cx:lvl ptCount="3">` +
      `<cx:pt idx="0">5</cx:pt><cx:pt idx="1">7</cx:pt><cx:pt idx="2">9</cx:pt>` +
      `</cx:lvl></cx:numDim></cx:data>`;
    const chart = parseChartExChartSpace(
      chartSpace({
        series: series("scatter", "Points"),
        data,
        axes: `<cx:axis id="0"><cx:valScaling/></cx:axis><cx:axis id="1"><cx:valScaling/></cx:axis>`,
      }),
    );
    expect(chart!.type).toBe("scatter");
    expect(chart!.series[0]!.xValues).toEqual([1, 2, 3]);
    expect(chart!.series[0]!.values).toEqual([5, 7, 9]);
  });

  it("keeps scheme colors as tokens and honours pt idx gaps", () => {
    const data =
      `<cx:data id="0">` +
      `<cx:strDim type="cat"><cx:lvl ptCount="3"><cx:pt idx="0">A</cx:pt><cx:pt idx="2">C</cx:pt></cx:lvl></cx:strDim>` +
      `<cx:numDim type="val"><cx:lvl ptCount="3"><cx:pt idx="0">1.5</cx:pt><cx:pt idx="2">3.5</cx:pt></cx:lvl></cx:numDim>` +
      `</cx:data>`;
    const scheme =
      `<cx:series layoutId="line"><cx:spPr><a:solidFill><a:schemeClr val="accent2"/></a:solidFill></cx:spPr>` +
      `<cx:dataId val="0"/></cx:series>`;
    const chart = parseChartExChartSpace(chartSpace({ series: scheme, data, axes: "" }));
    expect(chart!.series[0]!.color).toBe("accent2");
    expect(chart!.categories).toEqual(["A", "", "C"]);
    expect(chart!.series[0]!.values).toEqual([1.5, 0, 3.5]);
  });

  it("skips hidden series and a pareto's computed line", () => {
    const raw =
      series("clusteredColumn", "A") +
      `<cx:series layoutId="paretoLine" hidden="1"><cx:dataId val="0"/></cx:series>` +
      series("clusteredColumn", "B");
    const chart = parseChartExChartSpace(chartSpace({ series: raw }));
    expect(chart!.type).toBe("column");
    expect(chart!.series.map((entry) => entry.name)).toEqual(["A", "B"]);
    expect(chart!.series.map((entry) => entry.index)).toEqual([0, 1]);
  });

  it("accepts any namespace prefix on the payload", () => {
    const chart = parseChartExChartSpace(chartSpace({ series: series("column"), prefix: "cxx" }));
    expect(chart!.type).toBe("column");
    expect(chart!.categories).toEqual(["East", "West", "North"]);
  });

  it("leaves unsupported families, dashboards, and empty payloads to the placeholder", () => {
    for (const layoutId of [
      "waterfall",
      "treemap",
      "funnel",
      "boxWhisker",
      "regionMap",
      "sunburst",
      "histogram",
    ]) {
      expect(parseChartExChartSpace(chartSpace({ series: series(layoutId) }))).toBeUndefined();
    }
    // A dashboard part is not a chart.
    expect(
      parseChartExChartSpace(`<cx:chartSpace xmlns:cx="${CX}"><cx:dashboard/></cx:chartSpace>`),
    ).toBeUndefined();
    // A payload without cached values cannot be painted.
    expect(
      parseChartExChartSpace(
        `<cx:chartSpace xmlns:cx="${CX}"><cx:chartData/><cx:chart><cx:plotArea>` +
          `<cx:plotAreaRegion><cx:series layoutId="column"/></cx:plotAreaRegion></cx:plotArea></cx:chart></cx:chartSpace>`,
      ),
    ).toBeUndefined();
    // Malformed XML never throws.
    expect(parseChartExChartSpace("<cx:chartSpace")).toBeUndefined();
    expect(parseChartExChartSpace(undefined)).toBeUndefined();
  });

  it("rejects a combo whose visible series disagree on the family", () => {
    const raw = series("clusteredColumn") + series("line", "Second");
    expect(parseChartExChartSpace(chartSpace({ series: raw }))).toBeUndefined();
  });

  it("decodes a UTF-8 byte payload", () => {
    const xml = chartSpace({ series: series("column") });
    const chart = parseChartExChartSpace(new TextEncoder().encode(xml));
    expect(chart!.type).toBe("column");
  });
});

describe("chartExToChartSpace", () => {
  it("projects a column chart onto the classic painter model", () => {
    const chart = parseChartExChartSpace(
      chartSpace({
        series: series("clusteredColumn"),
        title: "Revenue",
        legend: '<cx:legend pos="r"/>',
      }),
    )!;
    const model = chartExToChartSpace(chart);
    expect(model.type).toBe("column");
    expect(model.title).toBe("Revenue");
    expect(model.categories).toEqual(["East", "West", "North"]);
    expect(model.showLegend).toBe(true);
    expect(model.legendPosition).toBe("right");
    expect(model.grouping).toBe("clustered");
    expect(model.series[0]).toMatchObject({
      name: "2025",
      values: [10, 20, 30],
      shapeProperties: { fill: { type: "solid", color: { value: "4472C4" } } },
    });
    expect(model.axes?.map((axis) => axis.kind)).toEqual(["category", "value"]);
  });

  it("drops the series fill for pie/doughnut so slices keep the point palette", () => {
    const chart = parseChartExChartSpace(chartSpace({ series: series("pie") }))!;
    const model = chartExToChartSpace(chart);
    expect(model.type).toBe("pie");
    expect(model.series[0]!.shapeProperties).toBeUndefined();
  });

  it("emits scatter series with xValue/yValue pairs and positioned axes", () => {
    const data =
      `<cx:data id="0"><cx:numDim type="x"><cx:lvl ptCount="2">` +
      `<cx:pt idx="0">1</cx:pt><cx:pt idx="1">2</cx:pt></cx:lvl></cx:numDim>` +
      `<cx:numDim type="y"><cx:lvl ptCount="2">` +
      `<cx:pt idx="0">3</cx:pt><cx:pt idx="1">4</cx:pt></cx:lvl></cx:numDim></cx:data>`;
    const chart = parseChartExChartSpace(
      chartSpace({
        series: series("scatter"),
        data,
        axes: `<cx:axis id="0"><cx:valScaling/></cx:axis><cx:axis id="1"><cx:valScaling/></cx:axis>`,
      }),
    )!;
    const model = chartExToChartSpace(chart);
    expect(model.type).toBe("scatter");
    expect(model.series[0]).toMatchObject({ xValues: [1, 2], yValues: [3, 4] });
    expect(model.axes).toMatchObject([
      { kind: "value", position: "bottom" },
      { kind: "value", position: "left" },
    ]);
  });
});

describe("chartExPartOfRawMedia", () => {
  it("finds the chartEx part among a drawing's raw media", () => {
    const payload = new TextEncoder().encode(chartSpace({ series: series("column") }));
    const found = chartExPartOfRawMedia([
      { fileName: "image1.png", data: new Uint8Array([1, 2, 3]) },
      { fileName: "chartEx1.xml", data: payload, type: "png" },
    ]);
    expect(found?.fileName).toBe("chartEx1.xml");
    expect(found?.data).toBe(payload);
  });

  it("returns undefined when the drawing carries no chartEx part", () => {
    expect(chartExPartOfRawMedia(undefined)).toBeUndefined();
    expect(chartExPartOfRawMedia([])).toBeUndefined();
    expect(
      chartExPartOfRawMedia([{ fileName: "image1.png", data: new Uint8Array() }]),
    ).toBeUndefined();
  });

  it("accepts string payloads", () => {
    expect(chartExPartOfRawMedia([{ fileName: "chartEx2.xml", data: "<cx:chartSpace/>" }])).toEqual(
      {
        fileName: "chartEx2.xml",
        data: "<cx:chartSpace/>",
      },
    );
  });
});
