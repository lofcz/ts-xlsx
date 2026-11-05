import {AuthoringError} from '../errors.ts';
import {tokenSet} from '../token-set.ts';
import {type SheetSplice, spliceFormula} from './formula-references.ts';
import {shiftIndex} from './grid-shift.ts';
import {isOneCellAnchor, type ImageAnchor} from './image.ts';

/** Chart families authored as native DrawingML, editable by spreadsheet applications. */
export type ChartType =
  | 'column'
  | 'bar'
  | 'line'
  | 'area'
  | 'pie'
  | 'doughnut'
  | 'scatter'
  | 'bubble'
  | 'radar';
export const isChartType = tokenSet({
  column: true,
  bar: true,
  line: true,
  area: true,
  pie: true,
  doughnut: true,
  scatter: true,
  bubble: true,
  radar: true,
});

/** Values are the chart's cached data; a formula optionally links them to worksheet cells. */
export interface ChartData<T extends string | number> {
  readonly values: readonly (T | null)[];
  readonly formula?: string;
  readonly formatCode?: string;
}

export interface ChartSeries {
  readonly name: string;
  readonly nameFormula?: string;
  readonly values: ChartData<number>;
  readonly categories?: ChartData<string>;
  readonly xValues?: ChartData<number>;
  readonly bubbleSizes?: ChartData<number>;
  /** Six hexadecimal RGB digits, with an optional leading #. */
  readonly color?: string;
  /** Per-series type enables column/line/area combinations. */
  readonly type?: 'column' | 'line' | 'area';
  readonly secondaryAxis?: boolean;
}

export interface ChartAxis {
  readonly title?: string;
  readonly min?: number;
  readonly max?: number;
  readonly majorUnit?: number;
  readonly formatCode?: string;
}

export interface ChartOptions {
  readonly type: ChartType;
  readonly title?: string;
  readonly series: readonly ChartSeries[];
  readonly categories?: ChartData<string>;
  readonly grouping?: 'clustered' | 'stacked' | 'percentStacked';
  readonly legend?: false | 'top' | 'bottom' | 'left' | 'right';
  readonly categoryAxis?: ChartAxis;
  readonly valueAxis?: ChartAxis;
  readonly secondaryValueAxis?: ChartAxis;
  readonly showValues?: boolean;
  readonly varyColors?: boolean;
  readonly holeSize?: number;
}

/** A related XML part, such as a chart's user-shapes drawing. Paths are assigned by the writer. */
export interface ChartRelatedPart {
  readonly relationshipId: string;
  readonly relationshipType: string;
  readonly contentType: string;
  readonly xml: string;
}

/**
 * A complete standard DrawingML chart for producers with a richer chart model.
 * The writer owns package paths, drawing anchors, relationship ids and content types.
 * XML is validated for well-formedness; its schema conformance remains the author's responsibility.
 */
export interface ChartPart {
  readonly xml: string;
  readonly name?: string;
  readonly relatedParts?: readonly ChartRelatedPart[];
}

export type WorksheetChart =
  | {readonly kind: 'chart'; readonly options: ChartOptions; readonly anchor: ImageAnchor}
  | {
      readonly kind: 'xml';
      readonly part: ChartPart;
      readonly anchor: ImageAnchor;
      readonly splices?: readonly SheetSplice[];
    };

export function validateChart(options: ChartOptions): void {
  if (!isChartType(options.type)) throw new AuthoringError('unknown chart type');
  if (options.series.length === 0) throw new AuthoringError('a chart needs at least one series');
  if ((options.type === 'pie' || options.type === 'doughnut') && options.series.length !== 1) {
    throw new AuthoringError('pie and doughnut charts require one series');
  }
  if (
    options.holeSize !== undefined &&
    (!Number.isInteger(options.holeSize) || options.holeSize < 10 || options.holeSize > 90)
  ) {
    throw new AuthoringError('doughnut holeSize must be an integer from 10 to 90');
  }
  for (const series of options.series) {
    if (series.color !== undefined && !/^#?[\da-f]{6}$/i.test(series.color))
      throw new AuthoringError('chart colors must be six hexadecimal RGB digits');
    const length = series.values.values.length;
    if (length === 0) throw new AuthoringError('a chart series needs at least one value');
    for (const data of [series.values, series.xValues, series.bubbleSizes]) {
      if (data?.values.some((value) => value !== null && !Number.isFinite(value)))
        throw new AuthoringError('chart values must be finite numbers or null');
      if (data !== undefined && data.values.length !== length)
        throw new AuthoringError('chart series data lengths must agree');
    }
    const categories = series.categories ?? options.categories;
    if (categories !== undefined && categories.values.length !== length)
      throw new AuthoringError('chart categories and values must have equal lengths');
    if ((options.type === 'scatter' || options.type === 'bubble') && series.xValues === undefined)
      throw new AuthoringError('scatter and bubble charts require xValues');
    if (options.type === 'bubble' && series.bubbleSizes === undefined)
      throw new AuthoringError('bubble charts require bubbleSizes');
    if (series.type !== undefined && !['column', 'line', 'area'].includes(options.type))
      throw new AuthoringError('combination charts support column, line and area series');
    if (
      series.secondaryAxis &&
      !['column', 'bar', 'line', 'area', 'scatter', 'bubble'].includes(options.type)
    )
      throw new AuthoringError('this chart type does not support a secondary axis');
  }
  for (const axis of [options.categoryAxis, options.valueAxis, options.secondaryValueAxis]) {
    if (axis === undefined) continue;
    for (const value of [axis.min, axis.max, axis.majorUnit]) {
      if (value !== undefined && !Number.isFinite(value))
        throw new AuthoringError('chart axis values must be finite');
    }
    if (axis.min !== undefined && axis.max !== undefined && axis.min >= axis.max)
      throw new AuthoringError('chart axis min must be less than max');
    if (axis.majorUnit !== undefined && axis.majorUnit <= 0)
      throw new AuthoringError('chart axis majorUnit must be positive');
  }
}

export function spliceChart(
  chart: WorksheetChart,
  home: string,
  edit: SheetSplice,
): WorksheetChart {
  const movePoint = (point: ImageAnchor['from']): ImageAnchor['from'] => {
    const axis = edit.splice.axis === 'row' ? 'row' : 'col';
    return {...point, [axis]: shiftIndex(point[axis] + 1, edit.splice) - 1};
  };
  let anchor = chart.anchor;
  if (
    home.toLowerCase() === edit.sheet.toLowerCase() &&
    (isOneCellAnchor(anchor) || anchor.editAs !== 'absolute')
  ) {
    const from = movePoint(anchor.from);
    anchor = isOneCellAnchor(anchor)
      ? {...anchor, from}
      : {...anchor, from, to: movePoint(anchor.to)};
  }
  if (chart.kind === 'xml') return {...chart, anchor, splices: [...(chart.splices ?? []), edit]};
  const move = <T extends string | number>(data: ChartData<T>): ChartData<T> =>
    data.formula === undefined ? data : {...data, formula: spliceFormula(data.formula, home, edit)};
  const options: ChartOptions = {
    ...chart.options,
    ...(chart.options.categories === undefined ? {} : {categories: move(chart.options.categories)}),
    series: chart.options.series.map((series) => ({
      ...series,
      values: move(series.values),
      ...(series.categories === undefined ? {} : {categories: move(series.categories)}),
      ...(series.xValues === undefined ? {} : {xValues: move(series.xValues)}),
      ...(series.bubbleSizes === undefined ? {} : {bubbleSizes: move(series.bubbleSizes)}),
      ...(series.nameFormula === undefined
        ? {}
        : {nameFormula: spliceFormula(series.nameFormula, home, edit)}),
    })),
  };
  return {...chart, options, anchor};
}
