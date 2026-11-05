import {
  type ChartAxis,
  type ChartData,
  type ChartDataLabels,
  type ChartErrorBars,
  type ChartTrendline,
  type ChartOptions,
  type ChartSeries,
  type ChartType,
  type WorksheetChart,
  validateChart,
} from '../../core/chart.ts';
import type {AnchorPoint, ImageAnchor} from '../../core/image.ts';
import {INTERNAL} from '../../core/internal.ts';
import type {Worksheet} from '../../core/worksheet.ts';
import {tokenSet} from '../../token-set.ts';
import {localName} from '../../xml/xml-scan.ts';
import {xmlChild, xmlChildren, xmlText, xmlTree, type XmlNode} from '../../xml/xml-tree.ts';
import {resolveRelativePart} from '../opc/part-paths.ts';
import {
  type PackageAccessors,
  type PartRelationships,
  readPartRelationships,
} from '../opc/read-opc.ts';
import {parseChartEx} from './chart-ex-read.ts';
import {CHARTEX_NS} from './chart-ex-xml.ts';
import {importedShapes} from './shape-xml.ts';

const val = (node: XmlNode | undefined, ...path: string[]): string | undefined =>
  xmlChild(node, ...path)?.attrs.val;
const number = (node: XmlNode | undefined, ...path: string[]): number | undefined => {
  const raw = val(node, ...path);
  const n = raw === undefined ? Number.NaN : Number(raw);
  return Number.isFinite(n) ? n : undefined;
};

const isTrend = tokenSet({
  linear: true,
  exp: true,
  log: true,
  poly: true,
  power: true,
  movingAvg: true,
});
const isError = tokenSet({
  fixedVal: true,
  percentage: true,
  stdDev: true,
  stdErr: true,
  cust: true,
});
const isPosition = tokenSet({
  bestFit: true,
  b: true,
  ctr: true,
  inBase: true,
  inEnd: true,
  l: true,
  outEnd: true,
  r: true,
  t: true,
});

function readLabels(node: XmlNode | undefined): ChartDataLabels | undefined {
  if (!node) return undefined;
  const point = xmlChild(node, 'dLbl'),
    source = point ?? node;
  const position = val(source, 'dLblPos'),
    formatCode = xmlChild(source, 'numFmt')?.attrs.formatCode;
  const separator = xmlChild(source, 'separator');
  return {
    showValue: val(source, 'showVal') === '1',
    showCategory: val(source, 'showCatName') === '1',
    showSeriesName: val(source, 'showSerName') === '1',
    showPercent: val(source, 'showPercent') === '1',
    showLegendKey: val(source, 'showLegendKey') === '1',
    showBubbleSize: val(source, 'showBubbleSize') === '1',
    ...(position !== undefined && isPosition(position) ? {position} : {}),
    ...(formatCode ? {formatCode} : {}),
    ...(separator ? {separator: xmlText(separator)} : {}),
    ...(point ? {lastPointOnly: true} : {}),
  };
}

function readSeriesDetails(node: XmlNode): Partial<ChartSeries> {
  const trendlines: ChartTrendline[] = [];
  for (const trend of xmlChildren(node, 'trendline')) {
    const type = val(trend, 'trendlineType');
    if (type === undefined || !isTrend(type)) continue;
    const numbers: {
      order?: number;
      period?: number;
      forward?: number;
      backward?: number;
      intercept?: number;
    } = {};
    for (const key of ['order', 'period', 'forward', 'backward', 'intercept'] as const) {
      const value = number(trend, key);
      if (value !== undefined) numbers[key] = value;
    }
    const name = xmlText(xmlChild(trend, 'name'));
    trendlines.push({
      type,
      ...numbers,
      ...(name ? {name} : {}),
      displayEquation: val(trend, 'dispEq') === '1',
      displayRSquared: val(trend, 'dispRSqr') === '1',
    });
  }
  const errorBars: ChartErrorBars[] = [];
  for (const error of xmlChildren(node, 'errBars')) {
    const type = val(error, 'errValType');
    if (type === undefined || !isError(type)) continue;
    const value = number(error, 'val'),
      direction = val(error, 'errDir'),
      sides = val(error, 'errBarType');
    const plus = data<number>(xmlChild(error, 'plus'), true),
      minus = data<number>(xmlChild(error, 'minus'), true);
    errorBars.push({
      type,
      ...(value === undefined ? {} : {value}),
      ...(direction === 'x' || direction === 'y' ? {direction} : {}),
      ...(sides === 'both' || sides === 'plus' || sides === 'minus' ? {sides} : {}),
      ...(plus ? {plus} : {}),
      ...(minus ? {minus} : {}),
      endCap: val(error, 'noEndCap') !== '1',
    });
  }
  const labels = readLabels(xmlChild(node, 'dLbls'));
  const points = xmlChildren(node, 'dPt');
  const colors: (string | null)[] = [];
  for (const point of points) {
    const index = number(point, 'idx');
    if (index !== undefined && Number.isInteger(index) && index >= 0 && index < 100000) {
      while (colors.length <= index) colors.push(null);
      colors[index] = xmlChild(point, 'spPr', 'solidFill', 'srgbClr')?.attrs.val ?? null;
    }
  }
  return {
    ...(trendlines.length ? {trendlines} : {}),
    ...(errorBars.length ? {errorBars} : {}),
    ...(labels ? {dataLabels: labels} : {}),
    ...(colors.length ? {pointColors: colors} : {}),
    ...(val(node, 'smooth') === '1' ? {smooth: true} : {}),
  };
}

function data<T extends string | number>(
  parent: XmlNode | undefined,
  numeric: boolean,
): ChartData<T> | undefined {
  if (!parent) return undefined;
  const ref = xmlChild(parent, numeric ? 'numRef' : 'strRef') ?? xmlChild(parent, 'numRef');
  const cache = ref
    ? (xmlChild(ref, numeric ? 'numCache' : 'strCache') ?? xmlChild(ref, 'numCache'))
    : (xmlChild(parent, numeric ? 'numLit' : 'strLit') ?? xmlChild(parent, 'numLit'));
  if (!cache) return undefined;
  const points = xmlChildren(cache, 'pt');
  const count = number(cache, 'ptCount') ?? points.length;
  if (!Number.isInteger(count) || count < 0 || count > 100000) return undefined;
  const values: (string | number | null)[] = Array.from({length: count}, () => null);
  for (const point of points) {
    const index = Number(point.attrs.idx);
    const text = xmlText(xmlChild(point, 'v'));
    if (!Number.isInteger(index) || index < 0 || index >= count) continue;
    values[index] = numeric ? (Number.isFinite(Number(text)) ? Number(text) : null) : text;
  }
  const formula = xmlText(xmlChild(ref, 'f'));
  const formatCode = xmlText(xmlChild(cache, 'formatCode'));
  return {
    values: values as (T | null)[],
    ...(formula ? {formula} : {}),
    ...(formatCode ? {formatCode} : {}),
  };
}

function axis(node: XmlNode | undefined): ChartAxis | undefined {
  if (!node) return undefined;
  const min = number(node, 'scaling', 'min'),
    max = number(node, 'scaling', 'max'),
    majorUnit = number(node, 'majorUnit');
  const logBase = number(node, 'scaling', 'logBase'),
    minorUnit = number(node, 'minorUnit');
  const unit = (key: string) => {
    const raw = val(node, key);
    return raw === 'days' || raw === 'months' || raw === 'years' ? raw : undefined;
  };
  const baseTimeUnit = unit('baseTimeUnit'),
    majorTimeUnit = unit('majorTimeUnit'),
    minorTimeUnit = unit('minorTimeUnit');
  const title = xmlText(xmlChild(node, 'title', 'tx', 'rich'));
  const formatCode = xmlChild(node, 'numFmt')?.attrs.formatCode;
  return {
    ...(min === undefined ? {} : {min}),
    ...(max === undefined ? {} : {max}),
    ...(majorUnit === undefined ? {} : {majorUnit}),
    ...(title ? {title} : {}),
    ...(formatCode ? {formatCode} : {}),
    ...(logBase === undefined ? {} : {logBase}),
    ...(minorUnit === undefined ? {} : {minorUnit}),
    ...(localName(node.name) === 'dateAx' ? {type: 'date'} : {}),
    ...(val(node, 'scaling', 'orientation') === 'maxMin' ? {reversed: true} : {}),
    ...(baseTimeUnit ? {baseTimeUnit} : {}),
    ...(majorTimeUnit ? {majorTimeUnit} : {}),
    ...(minorTimeUnit ? {minorTimeUnit} : {}),
  };
}

/** Decode supported chart families; unfamiliar chart families remain opaque preserved parts. */
export function parseChart(xml: string): ChartOptions | undefined {
  const root = xmlTree(xml);
  if (Object.values(root.attrs).includes(CHARTEX_NS)) return parseChartEx(root);
  const chart = xmlChild(root, 'chart');
  const plot = xmlChild(chart, 'plotArea');
  if (!plot) return undefined;
  const series: ChartSeries[] = [];
  let type: ChartType | undefined;
  const groups = xmlChildren(plot).filter((node) => localName(node.name).endsWith('Chart'));
  let grouping: ChartOptions['grouping'];
  for (const group of groups) {
    const tag = localName(group.name);
    const families: Readonly<Record<string, ChartType>> = {
      lineChart: 'line',
      areaChart: 'area',
      pieChart: 'pie',
      doughnutChart: 'doughnut',
      scatterChart: 'scatter',
      bubbleChart: 'bubble',
      radarChart: 'radar',
      stockChart: 'stock',
    };
    const kind =
      tag === 'barChart' ? (val(group, 'barDir') === 'bar' ? 'bar' : 'column') : families[tag];
    if (!kind) return undefined;
    type ??= kind;
    const g = val(group, 'grouping');
    if (g === 'stacked' || g === 'percentStacked' || g === 'clustered') grouping = g;
    for (const ser of xmlChildren(group, 'ser')) {
      const values = data<number>(
        xmlChild(ser, kind === 'scatter' || kind === 'bubble' ? 'yVal' : 'val'),
        true,
      );
      if (!values || values.values.length === 0) return undefined;
      const named = data<string>(xmlChild(ser, 'tx'), false);
      const name = named?.values[0] ?? xmlText(xmlChild(ser, 'tx', 'v'));
      const cat = xmlChild(ser, 'cat');
      const categories = data<string | number>(
        cat,
        !!(xmlChild(cat, 'numRef') ?? xmlChild(cat, 'numLit')),
      );
      const xValues = data<number>(xmlChild(ser, 'xVal'), true),
        bubbleSizes = data<number>(xmlChild(ser, 'bubbleSize'), true);
      const color =
        xmlChild(ser, 'spPr', 'solidFill', 'srgbClr')?.attrs.val ??
        xmlChild(ser, 'spPr', 'ln', 'solidFill', 'srgbClr')?.attrs.val;
      series.push({
        name: String(name),
        values,
        ...(named?.formula ? {nameFormula: named.formula} : {}),
        ...(categories ? {categories} : {}),
        ...(xValues ? {xValues} : {}),
        ...(bubbleSizes ? {bubbleSizes} : {}),
        ...(color ? {color} : {}),
        ...(groups.length > 1 && (kind === 'column' || kind === 'line' || kind === 'area')
          ? {type: kind}
          : {}),
        ...(val(group, 'axId') !== val(groups[0], 'axId') ? {secondaryAxis: true} : {}),
        ...readSeriesDetails(ser),
      });
    }
  }
  if (!type || series.length === 0) return undefined;
  const title = xmlText(xmlChild(chart, 'title', 'tx', 'rich'));
  const legendPosition = val(chart, 'legend', 'legendPos');
  const positions: Record<string, 'top' | 'bottom' | 'left' | 'right'> = {
    t: 'top',
    b: 'bottom',
    l: 'left',
    r: 'right',
  };
  const categoryAxis = axis(xmlChild(plot, 'catAx') ?? xmlChild(plot, 'dateAx'));
  const valueAxes = xmlChildren(plot, 'valAx');
  const valueAxis = axis(valueAxes[0]),
    secondaryValueAxis = axis(valueAxes[1]);
  const holeSize = number(groups[0], 'holeSize');
  const labels = readLabels(xmlChild(groups[0], 'dLbls'));
  const options: ChartOptions = {
    type,
    series,
    ...(title ? {title} : {}),
    ...(grouping ? {grouping} : {}),
    legend: xmlChild(chart, 'legend') ? (positions[legendPosition ?? 'b'] ?? 'bottom') : false,
    ...(categoryAxis ? {categoryAxis} : {}),
    ...(valueAxis ? {valueAxis} : {}),
    ...(secondaryValueAxis ? {secondaryValueAxis} : {}),
    ...(holeSize === undefined ? {} : {holeSize}),
    ...(labels ? {dataLabels: labels} : {}),
    showValues: val(groups[0], 'dLbls', 'showVal') === '1',
    varyColors: val(groups[0], 'varyColors') === '1',
  };
  try {
    validateChart(options);
    return options;
  } catch {
    return undefined;
  }
}

function anchorPoint(node: XmlNode | undefined): AnchorPoint {
  const n = (key: string) => Number(xmlText(xmlChild(node, key))) || 0;
  return {col: n('col'), row: n('row'), colOff: n('colOff'), rowOff: n('rowOff')};
}

/** Discover editable charts without taking ownership away from the preserved drawing graph. */
export function readSheetCharts(
  pkg: PackageAccessors,
  rels: PartRelationships,
  sheet: Worksheet,
): void {
  const path = rels.targetPath('drawing');
  if (!path) return;
  const drawing = pkg.partText(path);
  if (!drawing) return;
  sheet[INTERNAL].shapes.push(...importedShapes(drawing));
  const relationships = readPartRelationships(path, pkg.partText);
  const charts: WorksheetChart[] = [];
  for (const node of xmlChildren(xmlTree(drawing))) {
    const chart =
      xmlChild(node, 'graphicFrame', 'graphic', 'graphicData', 'chart') ??
      xmlChild(
        node,
        'AlternateContent',
        'Choice',
        'graphicFrame',
        'graphic',
        'graphicData',
        'chart',
      );
    const id = chart?.attrs['r:id'];
    if (!id) continue;
    const rel = relationships.byId(id);
    if (!rel || rel.external) continue;
    const xml = pkg.partText(resolveRelativePart(path, rel.target));
    if (!xml) continue;
    const options = parseChart(xml);
    if (!options) continue;
    const from = anchorPoint(xmlChild(node, 'from'));
    const to = xmlChild(node, 'to');
    const ext = xmlChild(node, 'ext');
    const anchor: ImageAnchor = to
      ? {
          from,
          to: anchorPoint(to),
          editAs:
            node.attrs.editAs === 'absolute'
              ? 'absolute'
              : node.attrs.editAs === 'oneCell'
                ? 'oneCell'
                : 'twoCell',
        }
      : {from, ext: {cx: Number(ext?.attrs.cx) || 1, cy: Number(ext?.attrs.cy) || 1}};
    charts.push({
      kind: 'chart',
      options,
      anchor,
      source: {
        relationshipId: id,
        name:
          xmlChild(
            xmlChild(node, 'graphicFrame') ??
              xmlChild(node, 'AlternateContent', 'Choice', 'graphicFrame'),
            'nvGraphicFramePr',
            'cNvPr',
          )?.attrs.name ?? id,
        xml,
        options,
        anchor,
      },
    });
  }
  sheet[INTERNAL].restoreCharts(charts);
}
