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
  | 'radar'
  | 'stock'
  | 'waterfall'
  | 'histogram'
  | 'pareto'
  | 'funnel'
  | 'treemap'
  | 'sunburst'
  | 'boxWhisker';
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
  stock: true,
  waterfall: true,
  histogram: true,
  pareto: true,
  funnel: true,
  treemap: true,
  sunburst: true,
  boxWhisker: true,
});

/** Office 2016 charts use the chartEx vocabulary and relationship type. */
export function isExtendedChart(type: ChartType): boolean {
  return [
    'waterfall',
    'histogram',
    'pareto',
    'funnel',
    'treemap',
    'sunburst',
    'boxWhisker',
  ].includes(type);
}

export interface ChartTrendline {
  readonly type: 'linear' | 'exp' | 'log' | 'poly' | 'power' | 'movingAvg';
  readonly name?: string;
  readonly order?: number;
  readonly period?: number;
  readonly forward?: number;
  readonly backward?: number;
  readonly intercept?: number;
  readonly displayEquation?: boolean;
  readonly displayRSquared?: boolean;
}

export interface ChartErrorBars {
  readonly type: 'fixedVal' | 'percentage' | 'stdDev' | 'stdErr' | 'cust';
  readonly direction?: 'x' | 'y';
  readonly sides?: 'both' | 'plus' | 'minus';
  readonly value?: number;
  readonly plus?: ChartData<number>;
  readonly minus?: ChartData<number>;
  readonly endCap?: boolean;
}

export interface ChartDataLabels {
  readonly showValue?: boolean;
  readonly showCategory?: boolean;
  readonly showSeriesName?: boolean;
  readonly showPercent?: boolean;
  readonly showLegendKey?: boolean;
  readonly showBubbleSize?: boolean;
  readonly position?: 'bestFit' | 'b' | 'ctr' | 'inBase' | 'inEnd' | 'l' | 'outEnd' | 'r' | 't';
  readonly formatCode?: string;
  readonly separator?: string;
  readonly lastPointOnly?: boolean;
}

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
  readonly categories?: ChartData<string | number>;
  /** Outer-to-inner category levels for hierarchy charts. */
  readonly categoryLevels?: readonly ChartData<string>[];
  readonly xValues?: ChartData<number>;
  readonly bubbleSizes?: ChartData<number>;
  /** Six hexadecimal RGB digits, with an optional leading #. */
  readonly color?: string;
  /** Per-series type enables column/line/area combinations. */
  readonly type?: 'column' | 'line' | 'area';
  readonly secondaryAxis?: boolean;
  readonly trendlines?: readonly ChartTrendline[];
  readonly errorBars?: readonly ChartErrorBars[];
  readonly dataLabels?: ChartDataLabels;
  readonly pointColors?: readonly (string | null)[];
  readonly smooth?: boolean;
}

export interface ChartAxis {
  readonly title?: string;
  readonly min?: number;
  readonly max?: number;
  readonly majorUnit?: number;
  readonly formatCode?: string;
  readonly type?: 'category' | 'date' | 'value';
  readonly logBase?: number;
  readonly reversed?: boolean;
  readonly minorUnit?: number;
  readonly baseTimeUnit?: 'days' | 'months' | 'years';
  readonly majorTimeUnit?: 'days' | 'months' | 'years';
  readonly minorTimeUnit?: 'days' | 'months' | 'years';
}

export interface ChartOptions {
  readonly type: ChartType;
  readonly title?: string;
  readonly series: readonly ChartSeries[];
  readonly categories?: ChartData<string | number>;
  readonly grouping?: 'clustered' | 'stacked' | 'percentStacked';
  readonly legend?: false | 'top' | 'bottom' | 'left' | 'right';
  readonly categoryAxis?: ChartAxis;
  readonly valueAxis?: ChartAxis;
  readonly secondaryValueAxis?: ChartAxis;
  readonly showValues?: boolean;
  readonly varyColors?: boolean;
  readonly holeSize?: number;
  readonly dataLabels?: ChartDataLabels;
  /** Zero-based waterfall points rendered as totals. */
  readonly waterfallTotals?: readonly number[];
  readonly waterfallConnectors?: boolean;
  readonly binning?: {
    readonly width?: number;
    readonly count?: number;
    readonly underflow?: number;
    readonly overflow?: number;
  };
  readonly boxWhisker?: {
    readonly quartileMethod?: 'inclusive' | 'exclusive';
    readonly showMeanLine?: boolean;
    readonly showMeanMarker?: boolean;
    readonly showInnerPoints?: boolean;
    readonly showOutlierPoints?: boolean;
  };
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

/** Original drawing identity and markup retained while an imported chart is edited. */
export interface ChartSource {
  readonly relationshipId?: string;
  readonly name?: string;
  readonly editedXml?: string;
  readonly xml: string;
  readonly options: ChartOptions;
  readonly anchor: ImageAnchor;
}

export type WorksheetChart =
  | {
      readonly kind: 'chart';
      readonly options: ChartOptions;
      readonly anchor: ImageAnchor;
      readonly source?: ChartSource;
    }
  | {
      readonly kind: 'xml';
      readonly part: ChartPart;
      readonly anchor: ImageAnchor;
      readonly splices?: readonly SheetSplice[];
    };

export function validateChart(options: ChartOptions): void {
  if (!isChartType(options.type)) throw new AuthoringError('unknown chart type');
  if (options.series.length === 0) throw new AuthoringError('a chart needs at least one series');
  if (isExtendedChart(options.type) && options.type !== 'boxWhisker' && options.series.length !== 1)
    throw new AuthoringError('this extended chart requires one series');
  if (options.type === 'stock' && ![3, 4].includes(options.series.length))
    throw new AuthoringError('stock charts require high/low/close or open/high/low/close series');
  for (const [key, value] of Object.entries(options.binning ?? {})) {
    if (
      !Number.isFinite(value) ||
      ((key === 'width' || key === 'count') && value <= 0) ||
      (key === 'count' && !Number.isInteger(value))
    )
      throw new AuthoringError('invalid histogram binning');
  }
  if (options.binning?.width !== undefined && options.binning.count !== undefined)
    throw new AuthoringError('histogram bin width and count are mutually exclusive');
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
    for (const trend of series.trendlines ?? []) {
      if (
        isExtendedChart(options.type) ||
        ['pie', 'doughnut', 'radar', 'stock'].includes(options.type)
      )
        throw new AuthoringError('this chart type does not support trendlines');
      if (!['linear', 'exp', 'log', 'poly', 'power', 'movingAvg'].includes(trend.type))
        throw new AuthoringError('invalid trendline type');
      for (const value of [
        trend.order,
        trend.period,
        trend.forward,
        trend.backward,
        trend.intercept,
      ])
        if (value !== undefined && !Number.isFinite(value))
          throw new AuthoringError('trendline values must be finite');
      if (
        trend.type === 'poly' &&
        (trend.order === undefined ||
          !Number.isInteger(trend.order) ||
          trend.order < 2 ||
          trend.order > 6)
      )
        throw new AuthoringError('polynomial trendline order must be an integer from 2 to 6');
      if (
        trend.type === 'movingAvg' &&
        (trend.period === undefined ||
          !Number.isInteger(trend.period) ||
          trend.period < 2 ||
          trend.period > 255)
      )
        throw new AuthoringError('moving average period must be an integer from 2 to 255');
      if ((trend.forward ?? 0) < 0 || (trend.backward ?? 0) < 0)
        throw new AuthoringError('trendline forecast lengths must be nonnegative');
    }
    for (const error of series.errorBars ?? []) {
      if (
        ['pie', 'doughnut', 'radar', 'stock'].includes(options.type) ||
        isExtendedChart(options.type)
      )
        throw new AuthoringError('this chart type does not support error bars');
      if (error.type === 'cust') {
        for (const side of error.sides === 'plus'
          ? (['plus'] as const)
          : error.sides === 'minus'
            ? (['minus'] as const)
            : (['plus', 'minus'] as const))
          if (error[side]?.values.length !== length)
            throw new AuthoringError('custom error bars must match the series length');
      } else if (
        error.type !== 'stdErr' &&
        (error.value === undefined || !Number.isFinite(error.value) || error.value < 0)
      )
        throw new AuthoringError('error bar value must be finite and nonnegative');
    }
    for (const level of series.categoryLevels ?? [])
      if (level.values.length !== length)
        throw new AuthoringError('hierarchy categories must match the series length');
    for (const color of series.pointColors ?? [])
      if (color !== null && !/^#?[\da-f]{6}$/i.test(color))
        throw new AuthoringError('chart point colors must be six hexadecimal RGB digits');
    if (length === 0) throw new AuthoringError('a chart series needs at least one value');
    for (const data of [series.values, series.xValues, series.bubbleSizes]) {
      if (data?.values.some((value) => value !== null && !Number.isFinite(value)))
        throw new AuthoringError('chart values must be finite numbers or null');
      if (data !== undefined && data.values.length !== length)
        throw new AuthoringError('chart series data lengths must agree');
    }
    const categories = series.categories ?? options.categories;
    if (
      options.categoryAxis?.type === 'date' &&
      categories?.values.some((value) => value !== null && typeof value !== 'number')
    )
      throw new AuthoringError('date axis categories must be numeric date serials');
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
  for (const index of options.waterfallTotals ?? [])
    if (
      !Number.isInteger(index) ||
      index < 0 ||
      index >= (options.series[0]?.values.values.length ?? 0)
    )
      throw new AuthoringError('waterfall total indices must identify a series point');
  for (const axis of [options.categoryAxis, options.valueAxis, options.secondaryValueAxis]) {
    if (axis === undefined) continue;
    if (
      isExtendedChart(options.type) &&
      (axis.type === 'date' || axis.logBase !== undefined || axis.reversed)
    )
      throw new AuthoringError('chartEx axes do not support date, logarithmic or reversed scales');
    for (const value of [axis.min, axis.max, axis.majorUnit, axis.minorUnit, axis.logBase]) {
      if (value !== undefined && !Number.isFinite(value))
        throw new AuthoringError('chart axis values must be finite');
    }
    if (axis.min !== undefined && axis.max !== undefined && axis.min >= axis.max)
      throw new AuthoringError('chart axis min must be less than max');
    if (axis.majorUnit !== undefined && axis.majorUnit <= 0)
      throw new AuthoringError('chart axis majorUnit must be positive');
    if (axis.minorUnit !== undefined && axis.minorUnit <= 0)
      throw new AuthoringError('chart axis minorUnit must be positive');
    if (
      axis.logBase !== undefined &&
      (axis.logBase < 2 || axis.logBase > 1000 || (axis.min !== undefined && axis.min <= 0))
    )
      throw new AuthoringError('log axis base must be from 2 to 1000 and minimum must be positive');
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
      ...(series.categoryLevels ? {categoryLevels: series.categoryLevels.map(move)} : {}),
      ...(series.errorBars
        ? {
            errorBars: series.errorBars.map((error) => ({
              ...error,
              ...(error.plus ? {plus: move(error.plus)} : {}),
              ...(error.minus ? {minus: move(error.minus)} : {}),
            })),
          }
        : {}),
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
