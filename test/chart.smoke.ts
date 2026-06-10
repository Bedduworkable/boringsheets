// Typecheck-only smoke test for the chart renderer's public API.
// Proves the exported types are usable and drawChart accepts them.
// This file is not meant to be executed; it exists for `tsc --noEmit`.

import {
  drawChart,
  CHART_PALETTE,
  type ChartSpec,
  type ChartType,
  type ChartSeries,
} from "../src/charts/render";

// Minimal no-op stub implementing only the methods drawChart actually calls.
const stub = {
  save() {},
  restore() {},
  beginPath() {},
  closePath() {},
  moveTo(_x: number, _y: number) {},
  lineTo(_x: number, _y: number) {},
  arc(_x: number, _y: number, _r: number, _a0: number, _a1: number) {},
  rect(_x: number, _y: number, _w: number, _h: number) {},
  fillRect(_x: number, _y: number, _w: number, _h: number) {},
  stroke() {},
  fill() {},
  translate(_x: number, _y: number) {},
  rotate(_a: number) {},
  fillText(_t: string, _x: number, _y: number) {},
  measureText(_t: string) {
    return { width: 10 } as TextMetrics;
  },
  font: "",
  fillStyle: "" as string | CanvasGradient | CanvasPattern,
  strokeStyle: "" as string | CanvasGradient | CanvasPattern,
  lineWidth: 1,
  lineJoin: "round" as CanvasLineJoin,
  globalAlpha: 1,
  textAlign: "left" as CanvasTextAlign,
  textBaseline: "alphabetic" as CanvasTextBaseline,
} as unknown as CanvasRenderingContext2D;

const sampleSeries: ChartSeries[] = [
  { name: "Revenue", values: [10, 20, 15, -5] },
  { name: "Cost", values: [8, 12, 9, 4], color: "#123456" },
];

const specs: ChartSpec[] = [
  {
    type: "column",
    title: "Quarterly",
    categories: ["Q1", "Q2", "Q3", "Q4"],
    series: sampleSeries,
  },
  {
    type: "bar",
    categories: ["A", "B", "C", "D"],
    series: [{ name: "S", values: [3, 3, 3, 3] }], // all-equal
  },
  {
    type: "line",
    categories: ["Jan", "Feb", "Mar", "Apr"],
    series: sampleSeries,
  },
  {
    type: "area",
    categories: ["Jan", "Feb", "Mar", "Apr"],
    series: [{ name: "Flow", values: [1, 4, 2, 8] }],
  },
  {
    type: "pie",
    title: "Share",
    categories: ["Red", "Green", "Blue"],
    series: [{ name: "Mix", values: [30, 50, 20] }],
  },
  {
    type: "scatter",
    categories: ["p1", "p2", "p3"],
    series: [
      { name: "X", values: [1, 2, 3] },
      { name: "Y", values: [4, 1, 7] },
    ],
  },
  {
    type: "column",
    categories: [],
    series: [], // empty data must not throw / type-check fine
  },
];

const everyType: ChartType[] = [
  "column",
  "bar",
  "line",
  "area",
  "pie",
  "scatter",
];

export function runSmoke(): void {
  for (const spec of specs) {
    drawChart(stub, spec, 400, 300);
  }
  // Touch the palette and the type list so they are referenced.
  void CHART_PALETTE.length;
  void everyType.length;
}
