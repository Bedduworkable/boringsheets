// Pure Canvas2D chart renderer. No DOM access except the passed context.

export type ChartType = "column" | "bar" | "line" | "area" | "pie" | "scatter";

export interface ChartSeries {
  name: string;
  values: number[];
  color?: string;
}

export interface ChartSpec {
  type: ChartType;
  title?: string;
  categories: string[]; // x-axis / slice labels, one per data point
  series: ChartSeries[]; // one or more data series
}

/** A default categorical palette used when a series has no explicit color. */
export const CHART_PALETTE: string[] = [
  "#4e79a7",
  "#f28e2b",
  "#e15759",
  "#76b7b2",
  "#59a14f",
  "#edc948",
  "#b07aa1",
  "#ff9da7",
  "#9c755f",
  "#bab0ac",
];

const AXIS_COLOR = "#888888";
const GRID_COLOR = "#e0e0e0";
const TEXT_COLOR = "#333333";
const TITLE_FONT = "bold 14px sans-serif";
const LABEL_FONT = "11px sans-serif";

interface Rect {
  x: number;
  y: number;
  w: number;
  h: number;
}

function seriesColor(series: ChartSeries, index: number): string {
  return series.color ?? CHART_PALETTE[index % CHART_PALETTE.length];
}

/** Compute a "nice" rounded number, optionally rounding to a clean step. */
function niceNum(range: number, round: boolean): number {
  if (range <= 0 || !isFinite(range)) {
    return 1;
  }
  const exponent = Math.floor(Math.log10(range));
  const fraction = range / Math.pow(10, exponent);
  let niceFraction: number;
  if (round) {
    if (fraction < 1.5) niceFraction = 1;
    else if (fraction < 3) niceFraction = 2;
    else if (fraction < 7) niceFraction = 5;
    else niceFraction = 10;
  } else {
    if (fraction <= 1) niceFraction = 1;
    else if (fraction <= 2) niceFraction = 2;
    else if (fraction <= 5) niceFraction = 5;
    else niceFraction = 10;
  }
  return niceFraction * Math.pow(10, exponent);
}

interface Ticks {
  min: number;
  max: number;
  step: number;
  values: number[];
}

/** Produce ~4-5 nicely rounded ticks spanning [dataMin, dataMax]. */
function computeTicks(dataMin: number, dataMax: number): Ticks {
  let lo = dataMin;
  let hi = dataMax;
  if (!isFinite(lo) || !isFinite(hi)) {
    lo = 0;
    hi = 1;
  }
  if (lo === hi) {
    // All-equal values: pad around the value so the axis has extent.
    const pad = lo === 0 ? 1 : Math.abs(lo) * 0.5;
    lo -= pad;
    hi += pad;
  }
  // Ensure zero is included for bar/column-style charts visually; callers that
  // need a zero baseline can pass a range that already includes 0.
  const desiredTicks = 4;
  const range = niceNum(hi - lo, false);
  const step = niceNum(range / (desiredTicks - 1), true);
  const niceMin = Math.floor(lo / step) * step;
  const niceMax = Math.ceil(hi / step) * step;
  const values: number[] = [];
  // Use a small epsilon to avoid float drift dropping the final tick.
  for (let v = niceMin; v <= niceMax + step * 1e-9; v += step) {
    values.push(Math.abs(v) < step * 1e-9 ? 0 : v);
  }
  return { min: niceMin, max: niceMax, step, values };
}

function formatTick(v: number, step: number): string {
  if (v === 0) return "0";
  // Decide decimal places from the step magnitude.
  const decimals = step < 1 ? Math.min(4, Math.ceil(-Math.log10(step))) : 0;
  const abs = Math.abs(v);
  if (abs >= 1e6 || (abs > 0 && abs < 1e-3)) {
    return v.toExponential(1);
  }
  return v.toFixed(decimals);
}

/** Flatten all finite values across series. */
function allValues(spec: ChartSpec): number[] {
  const out: number[] = [];
  for (const s of spec.series) {
    for (const v of s.values) {
      if (isFinite(v)) out.push(v);
    }
  }
  return out;
}

function isEmpty(spec: ChartSpec): boolean {
  if (!spec.series || spec.series.length === 0) return true;
  return allValues(spec).length === 0;
}

function drawTitle(
  ctx: CanvasRenderingContext2D,
  title: string | undefined,
  width: number
): number {
  if (!title) return 0;
  ctx.save();
  ctx.font = TITLE_FONT;
  ctx.fillStyle = TEXT_COLOR;
  ctx.textAlign = "center";
  ctx.textBaseline = "top";
  ctx.fillText(title, width / 2, 6);
  ctx.restore();
  return 26; // reserved vertical space
}

function drawEmpty(
  ctx: CanvasRenderingContext2D,
  width: number,
  height: number
): void {
  ctx.save();
  ctx.font = LABEL_FONT;
  ctx.fillStyle = AXIS_COLOR;
  ctx.textAlign = "center";
  ctx.textBaseline = "middle";
  ctx.fillText("No data", width / 2, height / 2);
  ctx.restore();
}

/** Draw the legend; returns the height it consumed at the bottom. */
function drawLegend(
  ctx: CanvasRenderingContext2D,
  labels: string[],
  colors: string[],
  width: number,
  height: number
): number {
  if (labels.length === 0) return 0;
  ctx.save();
  ctx.font = LABEL_FONT;
  ctx.textAlign = "left";
  ctx.textBaseline = "middle";
  const swatch = 10;
  const gap = 6;
  const itemGap = 14;
  const y = height - 12;
  // Measure total width to center the legend row.
  let total = 0;
  for (const label of labels) {
    total += swatch + gap + ctx.measureText(label).width + itemGap;
  }
  total -= itemGap;
  let x = Math.max(4, (width - total) / 2);
  for (let i = 0; i < labels.length; i++) {
    ctx.fillStyle = colors[i % colors.length];
    ctx.fillRect(x, y - swatch / 2, swatch, swatch);
    x += swatch + gap;
    ctx.fillStyle = TEXT_COLOR;
    ctx.fillText(labels[i], x, y);
    x += ctx.measureText(labels[i]).width + itemGap;
  }
  ctx.restore();
  return 24;
}

interface AxisLayout {
  plot: Rect;
  ticks: Ticks;
  valueToY: (v: number) => number; // for column/line/area/scatter (vertical value axis)
}

/** Draw axes, gridlines, y tick labels for a vertical-value chart. */
function drawVerticalAxes(
  ctx: CanvasRenderingContext2D,
  plot: Rect,
  ticks: Ticks
): (v: number) => number {
  const valueToY = (v: number): number =>
    plot.y + plot.h - ((v - ticks.min) / (ticks.max - ticks.min)) * plot.h;

  ctx.save();
  ctx.font = LABEL_FONT;
  ctx.textAlign = "right";
  ctx.textBaseline = "middle";
  for (const t of ticks.values) {
    const y = valueToY(t);
    ctx.strokeStyle = t === 0 ? AXIS_COLOR : GRID_COLOR;
    ctx.lineWidth = 1;
    ctx.beginPath();
    ctx.moveTo(plot.x, Math.round(y) + 0.5);
    ctx.lineTo(plot.x + plot.w, Math.round(y) + 0.5);
    ctx.stroke();
    ctx.fillStyle = TEXT_COLOR;
    ctx.fillText(formatTick(t, ticks.step), plot.x - 4, y);
  }
  // Axis lines.
  ctx.strokeStyle = AXIS_COLOR;
  ctx.lineWidth = 1;
  ctx.beginPath();
  ctx.moveTo(plot.x + 0.5, plot.y);
  ctx.lineTo(plot.x + 0.5, plot.y + plot.h);
  ctx.stroke();
  ctx.restore();
  return valueToY;
}

/** Draw category labels along the x-axis, skipping/rotating when dense. */
function drawCategoryLabels(
  ctx: CanvasRenderingContext2D,
  categories: string[],
  plot: Rect,
  centers: number[]
): void {
  if (categories.length === 0) return;
  ctx.save();
  ctx.font = LABEL_FONT;
  ctx.fillStyle = TEXT_COLOR;
  const baseY = plot.y + plot.h + 4;
  const slot = plot.w / categories.length;
  // Estimate whether labels fit horizontally.
  let maxLabel = 0;
  for (const c of categories) {
    maxLabel = Math.max(maxLabel, ctx.measureText(c).width);
  }
  const rotate = maxLabel > slot - 4;
  // Skip factor so we never draw more labels than will fit.
  const skip = rotate ? 1 : Math.max(1, Math.ceil((maxLabel + 6) / slot));
  for (let i = 0; i < categories.length; i++) {
    if (i % skip !== 0) continue;
    const cx = centers[i];
    if (rotate) {
      ctx.save();
      ctx.translate(cx, baseY);
      ctx.rotate(-Math.PI / 4);
      ctx.textAlign = "right";
      ctx.textBaseline = "middle";
      ctx.fillText(categories[i], 0, 0);
      ctx.restore();
    } else {
      ctx.textAlign = "center";
      ctx.textBaseline = "top";
      ctx.fillText(categories[i], cx, baseY);
    }
  }
  ctx.restore();
}

function setupVerticalLayout(
  ctx: CanvasRenderingContext2D,
  spec: ChartSpec,
  width: number,
  topReserved: number,
  bottomReserved: number,
  zeroBaseline: boolean
): AxisLayout {
  const vals = allValues(spec);
  let dataMin = Math.min(...vals);
  let dataMax = Math.max(...vals);
  if (zeroBaseline) {
    dataMin = Math.min(dataMin, 0);
    dataMax = Math.max(dataMax, 0);
  }
  const ticks = computeTicks(dataMin, dataMax);
  // Reserve room for y tick labels on the left.
  ctx.save();
  ctx.font = LABEL_FONT;
  let yLabelW = 0;
  for (const t of ticks.values) {
    yLabelW = Math.max(yLabelW, ctx.measureText(formatTick(t, ticks.step)).width);
  }
  ctx.restore();
  const leftPad = yLabelW + 10;
  const rightPad = 12;
  const plot: Rect = {
    x: leftPad,
    y: topReserved + 6,
    w: Math.max(10, width - leftPad - rightPad),
    h: 0,
  };
  return { plot, ticks, valueToY: (v) => v };
}

function drawColumnOrBar(
  ctx: CanvasRenderingContext2D,
  spec: ChartSpec,
  width: number,
  height: number,
  horizontal: boolean,
  topReserved: number,
  bottomReserved: number
): void {
  const layout = setupVerticalLayout(
    ctx,
    spec,
    width,
    topReserved,
    bottomReserved,
    true
  );
  const ticks = layout.ticks;
  const n = spec.categories.length;
  const seriesCount = spec.series.length;

  if (!horizontal) {
    const plot: Rect = {
      x: layout.plot.x,
      y: layout.plot.y,
      w: layout.plot.w,
      h: Math.max(10, height - layout.plot.y - bottomReserved),
    };
    const valueToY = drawVerticalAxes(ctx, plot, ticks);
    const zeroY = valueToY(0);
    const slot = n > 0 ? plot.w / n : plot.w;
    const groupPad = slot * 0.15;
    const groupW = slot - groupPad * 2;
    const barW = seriesCount > 0 ? groupW / seriesCount : groupW;
    const centers: number[] = [];
    for (let i = 0; i < n; i++) {
      centers.push(plot.x + slot * i + slot / 2);
    }
    for (let s = 0; s < seriesCount; s++) {
      ctx.fillStyle = seriesColor(spec.series[s], s);
      const series = spec.series[s];
      for (let i = 0; i < n; i++) {
        const v = series.values[i];
        if (v === undefined || !isFinite(v)) continue;
        const x = plot.x + slot * i + groupPad + barW * s;
        const y = valueToY(v);
        const top = Math.min(y, zeroY);
        const h = Math.abs(y - zeroY);
        ctx.fillRect(x, top, Math.max(1, barW - 1), h);
      }
    }
    drawCategoryLabels(ctx, spec.categories, plot, centers);
  } else {
    // Horizontal bars: value axis is horizontal, categories stacked vertically.
    ctx.save();
    ctx.font = LABEL_FONT;
    let catLabelW = 0;
    for (const c of spec.categories) {
      catLabelW = Math.max(catLabelW, ctx.measureText(c).width);
    }
    ctx.restore();
    const leftPad = Math.min(width * 0.4, catLabelW + 10);
    const plot: Rect = {
      x: leftPad,
      y: topReserved + 6,
      w: Math.max(10, width - leftPad - 12),
      h: Math.max(10, height - (topReserved + 6) - bottomReserved),
    };
    const valueToX = (v: number): number =>
      plot.x + ((v - ticks.min) / (ticks.max - ticks.min)) * plot.w;
    // Vertical gridlines + value labels.
    ctx.save();
    ctx.font = LABEL_FONT;
    ctx.textAlign = "center";
    ctx.textBaseline = "top";
    for (const t of ticks.values) {
      const x = valueToX(t);
      ctx.strokeStyle = t === 0 ? AXIS_COLOR : GRID_COLOR;
      ctx.lineWidth = 1;
      ctx.beginPath();
      ctx.moveTo(Math.round(x) + 0.5, plot.y);
      ctx.lineTo(Math.round(x) + 0.5, plot.y + plot.h);
      ctx.stroke();
      ctx.fillStyle = TEXT_COLOR;
      ctx.fillText(formatTick(t, ticks.step), x, plot.y + plot.h + 4);
    }
    ctx.strokeStyle = AXIS_COLOR;
    ctx.beginPath();
    ctx.moveTo(plot.x, plot.y + plot.h + 0.5);
    ctx.lineTo(plot.x + plot.w, plot.y + plot.h + 0.5);
    ctx.stroke();
    ctx.restore();

    const zeroX = valueToX(0);
    const slot = n > 0 ? plot.h / n : plot.h;
    const groupPad = slot * 0.15;
    const groupH = slot - groupPad * 2;
    const barH = seriesCount > 0 ? groupH / seriesCount : groupH;
    for (let s = 0; s < seriesCount; s++) {
      ctx.fillStyle = seriesColor(spec.series[s], s);
      const series = spec.series[s];
      for (let i = 0; i < n; i++) {
        const v = series.values[i];
        if (v === undefined || !isFinite(v)) continue;
        const y = plot.y + slot * i + groupPad + barH * s;
        const x = valueToX(v);
        const left = Math.min(x, zeroX);
        const w = Math.abs(x - zeroX);
        ctx.fillRect(left, y, w, Math.max(1, barH - 1));
      }
    }
    // Category labels on the left.
    ctx.save();
    ctx.font = LABEL_FONT;
    ctx.fillStyle = TEXT_COLOR;
    ctx.textAlign = "right";
    ctx.textBaseline = "middle";
    for (let i = 0; i < n; i++) {
      const cy = plot.y + slot * i + slot / 2;
      ctx.fillText(spec.categories[i], plot.x - 6, cy);
    }
    ctx.restore();
  }
}

function drawLineOrArea(
  ctx: CanvasRenderingContext2D,
  spec: ChartSpec,
  width: number,
  height: number,
  area: boolean,
  topReserved: number,
  bottomReserved: number
): void {
  const layout = setupVerticalLayout(
    ctx,
    spec,
    width,
    topReserved,
    bottomReserved,
    false
  );
  const ticks = layout.ticks;
  const n = spec.categories.length;
  const plot: Rect = {
    x: layout.plot.x,
    y: layout.plot.y,
    w: layout.plot.w,
    h: Math.max(10, height - layout.plot.y - bottomReserved),
  };
  const valueToY = drawVerticalAxes(ctx, plot, ticks);
  const zeroY = valueToY(0);
  // x position for each category center.
  const xAt = (i: number): number =>
    n <= 1 ? plot.x + plot.w / 2 : plot.x + (plot.w / (n - 1)) * i;
  const centers: number[] = [];
  for (let i = 0; i < n; i++) centers.push(xAt(i));

  for (let s = 0; s < spec.series.length; s++) {
    const series = spec.series[s];
    const color = seriesColor(series, s);
    // Build the polyline points (skip non-finite).
    const pts: Array<{ x: number; y: number }> = [];
    for (let i = 0; i < n; i++) {
      const v = series.values[i];
      if (v === undefined || !isFinite(v)) continue;
      pts.push({ x: xAt(i), y: valueToY(v) });
    }
    if (pts.length === 0) continue;
    if (area) {
      ctx.save();
      ctx.globalAlpha = 0.25;
      ctx.fillStyle = color;
      ctx.beginPath();
      ctx.moveTo(pts[0].x, zeroY);
      for (const p of pts) ctx.lineTo(p.x, p.y);
      ctx.lineTo(pts[pts.length - 1].x, zeroY);
      ctx.closePath();
      ctx.fill();
      ctx.restore();
    }
    ctx.save();
    ctx.strokeStyle = color;
    ctx.lineWidth = 2;
    ctx.lineJoin = "round";
    ctx.beginPath();
    pts.forEach((p, idx) => {
      if (idx === 0) ctx.moveTo(p.x, p.y);
      else ctx.lineTo(p.x, p.y);
    });
    ctx.stroke();
    // Markers.
    ctx.fillStyle = color;
    for (const p of pts) {
      ctx.beginPath();
      ctx.arc(p.x, p.y, 2.5, 0, Math.PI * 2);
      ctx.fill();
    }
    ctx.restore();
  }
  drawCategoryLabels(ctx, spec.categories, plot, centers);
}

function drawScatter(
  ctx: CanvasRenderingContext2D,
  spec: ChartSpec,
  width: number,
  height: number,
  topReserved: number,
  bottomReserved: number
): void {
  // Determine x/y data. If exactly 2 series, treat series[0]=x, series[1]=y.
  const pairMode = spec.series.length === 2;
  let xVals: number[];
  let plotSeries: ChartSeries[];
  if (pairMode) {
    xVals = spec.series[0].values;
    plotSeries = [spec.series[1]];
  } else {
    // x = category index.
    xVals = spec.categories.map((_, i) => i);
    plotSeries = spec.series;
  }

  // Compute x range.
  const finiteX = xVals.filter((v) => isFinite(v));
  let xMin = finiteX.length ? Math.min(...finiteX) : 0;
  let xMax = finiteX.length ? Math.max(...finiteX) : 1;
  const xTicks = computeTicks(xMin, xMax);
  xMin = xTicks.min;
  xMax = xTicks.max;

  // y range from plotted series.
  const yVals: number[] = [];
  for (const s of plotSeries) {
    for (const v of s.values) if (isFinite(v)) yVals.push(v);
  }
  const yTicks = computeTicks(
    yVals.length ? Math.min(...yVals) : 0,
    yVals.length ? Math.max(...yVals) : 1
  );

  ctx.save();
  ctx.font = LABEL_FONT;
  let yLabelW = 0;
  for (const t of yTicks.values) {
    yLabelW = Math.max(yLabelW, ctx.measureText(formatTick(t, yTicks.step)).width);
  }
  ctx.restore();
  const plot: Rect = {
    x: yLabelW + 10,
    y: topReserved + 6,
    w: 0,
    h: 0,
  };
  plot.w = Math.max(10, width - plot.x - 12);
  plot.h = Math.max(10, height - plot.y - bottomReserved);

  const valueToY = drawVerticalAxes(ctx, plot, yTicks);
  const valueToX = (v: number): number =>
    plot.x + ((v - xMin) / (xMax - xMin)) * plot.w;

  // x gridlines + labels.
  ctx.save();
  ctx.font = LABEL_FONT;
  ctx.textAlign = "center";
  ctx.textBaseline = "top";
  for (const t of xTicks.values) {
    const x = valueToX(t);
    ctx.strokeStyle = t === 0 ? AXIS_COLOR : GRID_COLOR;
    ctx.lineWidth = 1;
    ctx.beginPath();
    ctx.moveTo(Math.round(x) + 0.5, plot.y);
    ctx.lineTo(Math.round(x) + 0.5, plot.y + plot.h);
    ctx.stroke();
    ctx.fillStyle = TEXT_COLOR;
    ctx.fillText(formatTick(t, xTicks.step), x, plot.y + plot.h + 4);
  }
  ctx.strokeStyle = AXIS_COLOR;
  ctx.beginPath();
  ctx.moveTo(plot.x, plot.y + plot.h + 0.5);
  ctx.lineTo(plot.x + plot.w, plot.y + plot.h + 0.5);
  ctx.stroke();
  ctx.restore();

  for (let s = 0; s < plotSeries.length; s++) {
    const series = plotSeries[s];
    ctx.fillStyle = seriesColor(series, pairMode ? 0 : s);
    const count = series.values.length;
    for (let i = 0; i < count; i++) {
      const yv = series.values[i];
      const xv = pairMode ? xVals[i] : i;
      if (xv === undefined || !isFinite(xv) || !isFinite(yv)) continue;
      const px = valueToX(xv);
      const py = valueToY(yv);
      ctx.beginPath();
      ctx.arc(px, py, 3, 0, Math.PI * 2);
      ctx.fill();
    }
  }
}

function drawPie(
  ctx: CanvasRenderingContext2D,
  spec: ChartSpec,
  width: number,
  height: number,
  topReserved: number,
  bottomReserved: number
): void {
  const values = spec.series[0].values;
  // Only positive, finite values contribute to a pie.
  const slices: Array<{ value: number; label: string; color: string }> = [];
  let total = 0;
  for (let i = 0; i < values.length; i++) {
    const v = values[i];
    if (!isFinite(v) || v <= 0) continue;
    const color = CHART_PALETTE[i % CHART_PALETTE.length];
    slices.push({ value: v, label: spec.categories[i] ?? String(i), color });
    total += v;
  }
  if (total <= 0 || slices.length === 0) {
    drawEmpty(ctx, width, height);
    return;
  }
  const availH = height - topReserved - bottomReserved;
  const cx = width / 2;
  const cy = topReserved + availH / 2;
  const radius = Math.max(4, Math.min(width, availH) / 2 - 8);

  let angle = -Math.PI / 2;
  for (const slice of slices) {
    const frac = slice.value / total;
    const next = angle + frac * Math.PI * 2;
    ctx.save();
    ctx.fillStyle = slice.color;
    ctx.beginPath();
    ctx.moveTo(cx, cy);
    ctx.arc(cx, cy, radius, angle, next);
    ctx.closePath();
    ctx.fill();
    ctx.restore();
    // Percentage label at slice mid-angle.
    const mid = (angle + next) / 2;
    const lx = cx + Math.cos(mid) * radius * 0.6;
    const ly = cy + Math.sin(mid) * radius * 0.6;
    if (frac > 0.04) {
      ctx.save();
      ctx.font = LABEL_FONT;
      ctx.fillStyle = "#ffffff";
      ctx.textAlign = "center";
      ctx.textBaseline = "middle";
      ctx.fillText(`${Math.round(frac * 100)}%`, lx, ly);
      ctx.restore();
    }
    angle = next;
  }
}

/**
 * Draw the chart filling a width×height box (in CSS pixels) using the given
 * context. Assumes the context is already DPR-scaled and translated so (0,0) is
 * the chart's top-left; draws within [0,0,width,height]. Never creates DOM.
 */
export function drawChart(
  ctx: CanvasRenderingContext2D,
  spec: ChartSpec,
  width: number,
  height: number
): void {
  ctx.save();
  ctx.font = LABEL_FONT;
  ctx.textBaseline = "alphabetic";

  const topReserved = drawTitle(ctx, spec.title, width);

  if (isEmpty(spec)) {
    drawEmpty(ctx, width, height);
    ctx.restore();
    return;
  }

  // Build legend info and reserve bottom space.
  const showLegend = spec.type === "pie" || spec.series.length > 1;
  let legendLabels: string[] = [];
  let legendColors: string[] = [];
  if (spec.type === "pie") {
    legendLabels = spec.categories.slice();
    legendColors = spec.categories.map(
      (_, i) => CHART_PALETTE[i % CHART_PALETTE.length]
    );
  } else {
    legendLabels = spec.series.map((s) => s.name);
    legendColors = spec.series.map((s, i) => seriesColor(s, i));
  }

  // x-axis category labels need bottom room for non-pie charts.
  const axisBottom = spec.type === "pie" ? 0 : 34;
  const legendBottom = showLegend ? 24 : 0;
  const bottomReserved = axisBottom + legendBottom;

  switch (spec.type) {
    case "column":
      drawColumnOrBar(ctx, spec, width, height, false, topReserved, bottomReserved);
      break;
    case "bar":
      drawColumnOrBar(ctx, spec, width, height, true, topReserved, bottomReserved);
      break;
    case "line":
      drawLineOrArea(ctx, spec, width, height, false, topReserved, bottomReserved);
      break;
    case "area":
      drawLineOrArea(ctx, spec, width, height, true, topReserved, bottomReserved);
      break;
    case "scatter":
      drawScatter(ctx, spec, width, height, topReserved, bottomReserved);
      break;
    case "pie":
      drawPie(ctx, spec, width, height, topReserved, bottomReserved);
      break;
    default: {
      // Exhaustiveness guard; unknown type renders nothing meaningful.
      drawEmpty(ctx, width, height);
      break;
    }
  }

  if (showLegend) {
    drawLegend(ctx, legendLabels, legendColors, width, height);
  }

  ctx.restore();
}
