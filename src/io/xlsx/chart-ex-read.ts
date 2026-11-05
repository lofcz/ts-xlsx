import {
  type ChartAxis,
  type ChartDataLabels,
  type ChartData,
  type ChartOptions,
  type ChartSeries,
  type ChartType,
  isChartType,
  validateChart,
} from '../../core/chart.ts';
import {tokenSet} from '../../token-set.ts';
import {xmlChild, xmlChildren, xmlText, type XmlNode} from '../../xml/xml-tree.ts';

function dimension<T extends string | number>(
  node: XmlNode | undefined,
  numeric: boolean,
): ChartData<T> | undefined {
  const level = xmlChild(node, 'lvl');
  if (!level) return undefined;
  const count = Number(level.attrs.ptCount);
  if (!Number.isInteger(count) || count < 1 || count > 100000) return undefined;
  const values: (string | number | null)[] = Array.from({length: count}, () => null);
  for (const pt of xmlChildren(level, 'pt')) {
    const index = Number(pt.attrs.idx),
      text = xmlText(pt);
    if (Number.isInteger(index) && index >= 0 && index < count)
      values[index] = numeric ? (Number.isFinite(Number(text)) ? Number(text) : null) : text;
  }
  const formula = xmlText(xmlChild(node, 'f'));
  // The caller chooses T and numeric together from the dimension's schema type.
  return {
    values: values as (T | null)[],
    ...(formula ? {formula} : {}),
    ...(level.attrs.formatCode ? {formatCode: level.attrs.formatCode} : {}),
  };
}

const isLabelPosition = tokenSet({
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
  const visibility = xmlChild(node, 'visibility')?.attrs ?? {},
    position = node.attrs.pos;
  const formatCode = xmlChild(node, 'numFmt')?.attrs.formatCode,
    separator = xmlChild(node, 'separator');
  return {
    showValue: visibility.value !== '0',
    showCategory: visibility.categoryName === '1',
    showSeriesName: visibility.seriesName === '1',
    ...(position && isLabelPosition(position) ? {position} : {}),
    ...(formatCode ? {formatCode} : {}),
    ...(separator ? {separator: xmlText(separator)} : {}),
  };
}
function readAxis(node: XmlNode | undefined): ChartAxis | undefined {
  if (!node) return undefined;
  const scale = xmlChild(node, 'valScaling')?.attrs ?? {};
  const values: {min?: number; max?: number; majorUnit?: number; minorUnit?: number} = {};
  for (const key of ['min', 'max', 'majorUnit', 'minorUnit'] as const)
    if (scale[key] !== undefined && Number.isFinite(Number(scale[key])))
      values[key] = Number(scale[key]);
  const title = xmlText(xmlChild(node, 'title', 'tx', 'txData', 'v')),
    formatCode = xmlChild(node, 'numFmt')?.attrs.formatCode;
  return {...values, ...(title ? {title} : {}), ...(formatCode ? {formatCode} : {})};
}

export function parseChartEx(root: XmlNode): ChartOptions | undefined {
  const chart = xmlChild(root, 'chart'),
    plot = xmlChild(chart, 'plotArea'),
    region = xmlChild(plot, 'plotAreaRegion');
  if (!region) return undefined;
  const nodes = xmlChildren(region, 'series'),
    first = nodes[0];
  if (!first) return undefined;
  const layout = first.attrs.layoutId;
  const type: ChartType | undefined =
    layout === 'clusteredColumn'
      ? nodes.some((node) => node.attrs.layoutId === 'paretoLine')
        ? 'pareto'
        : 'histogram'
      : layout !== undefined && isChartType(layout)
        ? layout
        : undefined;
  if (!type) return undefined;
  const chartData = xmlChild(root, 'chartData');
  if (!chartData) return undefined;
  const series: ChartSeries[] = [];
  for (const node of nodes.filter((item) => item.attrs.layoutId !== 'paretoLine')) {
    const id = xmlChild(node, 'dataId')?.attrs.val;
    const source = xmlChildren(chartData, 'data').find((item) => item.attrs.id === id);
    if (!source) return undefined;
    const values = dimension<number>(
      xmlChildren(source, 'numDim').find((item) => item.attrs.type === 'val'),
      true,
    );
    if (!values) return undefined;
    const cats = xmlChildren(source, 'strDim').find((item) => item.attrs.type === 'cat');
    const categories = dimension<string>(cats, false);
    const levels = cats
      ? xmlChildren(cats, 'lvl')
          .map((level) =>
            dimension<string>(
              {
                ...cats,
                children: [
                  ...(xmlChild(cats, 'f')
                    ? [xmlChild(cats, 'f')].filter((item) => item !== undefined)
                    : []),
                  level,
                ],
              },
              false,
            ),
          )
          .filter((item) => item !== undefined)
      : [];
    const labels = readLabels(xmlChild(node, 'dataLabels'));
    const tx = xmlChild(node, 'tx', 'txData');
    const name = xmlText(xmlChild(tx, 'v')),
      nameFormula = xmlText(xmlChild(tx, 'f'));
    const color = xmlChild(node, 'spPr', 'solidFill', 'srgbClr')?.attrs.val;
    const points: (string | null)[] = Array.from({length: values.values.length}, () => null);
    for (const point of xmlChildren(node, 'dataPt')) {
      const index = Number(point.attrs.idx);
      if (Number.isInteger(index) && index >= 0 && index < points.length)
        points[index] = xmlChild(point, 'spPr', 'solidFill', 'srgbClr')?.attrs.val ?? null;
    }
    series.push({
      name,
      values,
      ...(labels ? {dataLabels: labels} : {}),
      ...(nameFormula ? {nameFormula} : {}),
      ...(categories ? {categories} : {}),
      ...(levels.length > 1 ? {categoryLevels: levels} : {}),
      ...(color ? {color} : {}),
      ...(points.some((item) => item !== null) ? {pointColors: points} : {}),
    });
  }
  const title = xmlText(xmlChild(chart, 'title', 'tx', 'txData', 'v'));
  const positions: Readonly<Record<string, 'top' | 'bottom' | 'left' | 'right'>> = {
    t: 'top',
    b: 'bottom',
    l: 'left',
    r: 'right',
  };
  const legend = xmlChild(chart, 'legend');
  const properties = xmlChild(first, 'layoutPr'),
    bins = xmlChild(properties, 'binning');
  const binning: {width?: number; count?: number; underflow?: number; overflow?: number} = {};
  for (const [key, raw] of [
    ['width', xmlText(xmlChild(bins, 'binSize'))],
    ['count', xmlText(xmlChild(bins, 'binCount'))],
    ['underflow', bins?.attrs.underflow],
    ['overflow', bins?.attrs.overflow],
  ] as const)
    if (raw && Number.isFinite(Number(raw))) binning[key] = Number(raw);
  const subtotals = xmlChild(properties, 'subtotals');
  const totals = subtotals
    ? xmlChildren(subtotals, 'idx')
        .map((node) => Number(node.attrs.val))
        .filter(Number.isInteger)
    : [];
  const axes = plot ? xmlChildren(plot, 'axis') : [];
  const categoryAxis = readAxis(axes.find((node) => xmlChild(node, 'catScaling'))),
    valueAxis = readAxis(axes.find((node) => node.attrs.id === '1'));
  const visibility = xmlChild(properties, 'visibility')?.attrs ?? {};
  const options: ChartOptions = {
    type,
    series,
    ...(categoryAxis ? {categoryAxis} : {}),
    ...(valueAxis ? {valueAxis} : {}),
    ...(type === 'boxWhisker'
      ? {
          boxWhisker: {
            quartileMethod:
              xmlChild(properties, 'statistics')?.attrs.quartileMethod === 'inclusive'
                ? 'inclusive'
                : 'exclusive',
            showMeanLine: visibility.meanLine === '1',
            showMeanMarker: visibility.meanMarker !== '0',
            showInnerPoints: visibility.nonoutliers === '1',
            showOutlierPoints: visibility.outliers !== '0',
          },
        }
      : {}),
    ...(title ? {title} : {}),
    legend: legend ? (positions[legend.attrs.pos ?? 'b'] ?? 'bottom') : false,
    ...(totals.length ? {waterfallTotals: totals} : {}),
    ...(type === 'waterfall'
      ? {waterfallConnectors: xmlChild(properties, 'visibility')?.attrs.connectorLines !== '0'}
      : {}),
    ...(Object.keys(binning).length ? {binning} : {}),
  };
  try {
    validateChart(options);
    return options;
  } catch {
    return undefined;
  }
}
