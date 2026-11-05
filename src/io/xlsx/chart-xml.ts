import {
  type ChartAxis,
  type ChartData,
  type ChartOptions,
  type ChartSeries,
  type ChartType,
  type ChartDataLabels,
  type ChartTrendline,
  type ChartErrorBars,
  isExtendedChart,
  validateChart,
} from '../../core/chart.ts';
import {escapeAttr, escapeText, XML_DECLARATION} from '../../xml/xml.ts';
import {chartExXml} from './chart-ex-xml.ts';
import {DRAWINGML_NS} from './namespaces.ts';

export const CHART_NS = 'http://schemas.openxmlformats.org/drawingml/2006/chart';

function title(text: string): string {
  return `<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>${escapeText(text)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>`;
}

export function dataXml(data: ChartData<string | number>, numeric: boolean): string {
  const kind = numeric ? 'num' : 'str';
  const points = data.values
    .map((value, i) =>
      value === null ? '' : `<c:pt idx="${i}"><c:v>${escapeText(String(value))}</c:v></c:pt>`,
    )
    .join('');
  const cache = `${numeric ? `<c:formatCode>${escapeText(data.formatCode ?? 'General')}</c:formatCode>` : ''}<c:ptCount val="${data.values.length}"/>${points}`;
  return data.formula === undefined
    ? `<c:${kind}Lit>${cache}</c:${kind}Lit>`
    : `<c:${kind}Ref><c:f>${escapeText(data.formula.replace(/^=/, ''))}</c:f><c:${kind}Cache>${cache}</c:${kind}Cache></c:${kind}Ref>`;
}

function trendlineXml(trend: ChartTrendline): string {
  const numeric = ['order', 'period', 'forward', 'backward', 'intercept'] as const;
  return `<c:trendline>${trend.name === undefined ? '' : `<c:name>${escapeText(trend.name)}</c:name>`}<c:trendlineType val="${trend.type}"/>${numeric.map((key) => (trend[key] === undefined ? '' : `<c:${key} val="${trend[key]}"/>`)).join('')}<c:dispRSqr val="${trend.displayRSquared ? 1 : 0}"/><c:dispEq val="${trend.displayEquation ? 1 : 0}"/></c:trendline>`;
}

function errorBarsXml(error: ChartErrorBars, type: ChartType): string {
  const direction =
    type === 'scatter' || type === 'bubble' || type === 'bar' || type === 'column'
      ? `<c:errDir val="${error.direction ?? 'y'}"/>`
      : '';
  return `<c:errBars>${direction}<c:errBarType val="${error.sides ?? 'both'}"/><c:errValType val="${error.type}"/><c:noEndCap val="${error.endCap === false ? 1 : 0}"/>${error.plus ? `<c:plus>${dataXml(error.plus, true)}</c:plus>` : ''}${error.minus ? `<c:minus>${dataXml(error.minus, true)}</c:minus>` : ''}${error.value === undefined ? '' : `<c:val val="${error.value}"/>`}</c:errBars>`;
}

function labelsXml(labels: ChartDataLabels | undefined, count: number): string {
  if (!labels) return '';
  const body = `${labels.formatCode === undefined ? '' : `<c:numFmt formatCode="${escapeAttr(labels.formatCode)}" sourceLinked="0"/>`}${labels.position === undefined ? '' : `<c:dLblPos val="${labels.position}"/>`}<c:showLegendKey val="${labels.showLegendKey ? 1 : 0}"/><c:showVal val="${labels.showValue === false ? 0 : 1}"/><c:showCatName val="${labels.showCategory ? 1 : 0}"/><c:showSerName val="${labels.showSeriesName ? 1 : 0}"/><c:showPercent val="${labels.showPercent ? 1 : 0}"/><c:showBubbleSize val="${labels.showBubbleSize ? 1 : 0}"/>${labels.separator === undefined ? '' : `<c:separator>${escapeText(labels.separator)}</c:separator>`}`;
  return `<c:dLbls>${labels.lastPointOnly ? `<c:dLbl><c:idx val="${count - 1}"/>${body}</c:dLbl><c:showVal val="0"/>` : body}</c:dLbls>`;
}

function seriesXml(
  options: ChartOptions,
  series: ChartSeries,
  index: number,
  type: ChartType,
): string {
  const tx =
    series.nameFormula === undefined
      ? `<c:v>${escapeText(series.name)}</c:v>`
      : dataXml({formula: series.nameFormula, values: [series.name]}, false);
  const color = series.color?.replace(/^#/, '').toUpperCase();
  const fill = color === undefined ? '' : `<a:solidFill><a:srgbClr val="${color}"/></a:solidFill>`;
  const style =
    fill === ''
      ? ''
      : `<c:spPr>${type === 'line' || type === 'scatter' || type === 'radar' ? `<a:ln>${fill}</a:ln>` : fill}</c:spPr>`;
  let data: string;
  if (type === 'scatter' || type === 'bubble') {
    const x = series.xValues;
    // validateChart guarantees the x source and bubble sizes before this serializer runs.
    if (x === undefined) throw new RangeError('missing chart xValues');
    data = `<c:xVal>${dataXml(x, true)}</c:xVal><c:yVal>${dataXml(series.values, true)}</c:yVal>`;
    if (type === 'bubble') {
      if (series.bubbleSizes === undefined) throw new RangeError('missing chart bubbleSizes');
      data += `<c:bubbleSize>${dataXml(series.bubbleSizes, true)}</c:bubbleSize>`;
    }
  } else {
    const categories = series.categories ??
      options.categories ?? {values: series.values.values.map((_, i) => String(i + 1))};
    data = `<c:cat>${dataXml(categories, options.categoryAxis?.type === 'date' || categories.values.some((value) => typeof value === 'number'))}</c:cat><c:val>${dataXml(series.values, true)}</c:val>`;
  }
  const marker =
    type === 'line' || type === 'scatter' || type === 'radar'
      ? '<c:marker><c:symbol val="circle"/><c:size val="5"/></c:marker>'
      : '';
  const points =
    series.pointColors
      ?.map((color, i) =>
        color === null
          ? ''
          : `<c:dPt><c:idx val="${i}"/><c:spPr><a:solidFill><a:srgbClr val="${color.replace(/^#/, '').toUpperCase()}"/></a:solidFill></c:spPr></c:dPt>`,
      )
      .join('') ?? '';
  const labels =
    series.dataLabels ?? (options.dataLabels?.lastPointOnly ? options.dataLabels : undefined);
  return `<c:ser><c:idx val="${index}"/><c:order val="${index}"/><c:tx>${tx}</c:tx>${style}${marker}${points}${labelsXml(labels, series.values.values.length)}${(series.trendlines ?? []).map(trendlineXml).join('')}${(series.errorBars ?? []).map((error) => errorBarsXml(error, type)).join('')}${data}${type === 'line' || type === 'scatter' || type === 'stock' ? `<c:smooth val="${series.smooth ? 1 : 0}"/>` : ''}</c:ser>`;
}

function axisXml(
  id: number,
  cross: number,
  value: boolean,
  position: string,
  settings: ChartAxis | undefined,
  secondary: boolean,
): string {
  const scaling = `<c:scaling>${settings?.logBase === undefined ? '' : `<c:logBase val="${settings.logBase}"/>`}<c:orientation val="${settings?.reversed ? 'maxMin' : 'minMax'}"/>${settings?.max === undefined ? '' : `<c:max val="${settings.max}"/>`}${settings?.min === undefined ? '' : `<c:min val="${settings.min}"/>`}</c:scaling>`;
  const body = `<c:axId val="${id}"/>${scaling}<c:delete val="${secondary && !value ? 1 : 0}"/><c:axPos val="${position}"/>${settings?.title === undefined ? '' : title(settings.title)}<c:numFmt formatCode="${escapeAttr(settings?.formatCode ?? 'General')}" sourceLinked="${settings?.formatCode === undefined ? 1 : 0}"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="${cross}"/><c:crosses val="${secondary ? 'max' : 'autoZero'}"/>`;
  const units = `${settings?.majorUnit === undefined ? '' : `<c:majorUnit val="${settings.majorUnit}"/>`}${!value && settings?.majorTimeUnit ? `<c:majorTimeUnit val="${settings.majorTimeUnit}"/>` : ''}${settings?.minorUnit === undefined ? '' : `<c:minorUnit val="${settings.minorUnit}"/>`}${!value && settings?.minorTimeUnit ? `<c:minorTimeUnit val="${settings.minorTimeUnit}"/>` : ''}`;
  if (!value && settings?.type === 'date')
    return `<c:dateAx>${body}<c:auto val="1"/><c:lblOffset val="100"/><c:baseTimeUnit val="${settings.baseTimeUnit ?? 'days'}"/>${units}</c:dateAx>`;
  return value
    ? `<c:valAx>${body}<c:crossBetween val="between"/>${units}</c:valAx>`
    : `<c:catAx>${body}<c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/></c:catAx>`;
}

/** Serialize a typed chart without allocating package paths or relationship ids. */
export function chartXml(options: ChartOptions): string {
  validateChart(options);
  if (isExtendedChart(options.type)) return chartExXml(options);
  const groups = new Map<
    string,
    {type: ChartType; secondary: boolean; series: {series: ChartSeries; index: number}[]}
  >();
  options.series.forEach((series, index) => {
    const type = series.type ?? options.type;
    const secondary = series.secondaryAxis ?? false;
    const key = `${type}:${secondary}`;
    let group = groups.get(key);
    if (group === undefined) {
      group = {type, secondary, series: []};
      groups.set(key, group);
    }
    group.series.push({series, index});
  });
  let plot = '';
  const axes = new Set<boolean>();
  for (const group of groups.values()) {
    const {type, secondary} = group;
    const tag = type === 'column' || type === 'bar' ? 'barChart' : `${type}Chart`;
    const grouping = options.grouping ?? 'clustered';
    let before = '';
    if (type === 'bar' || type === 'column')
      before = `<c:barDir val="${type === 'bar' ? 'bar' : 'col'}"/><c:grouping val="${grouping}"/>`;
    else if (type === 'line' || type === 'area')
      before = `<c:grouping val="${grouping === 'clustered' ? 'standard' : grouping}"/>`;
    else if (type === 'scatter') before = '<c:scatterStyle val="lineMarker"/>';
    else if (type === 'radar') before = '<c:radarStyle val="marker"/>';
    const series = group.series
      .map((item) => seriesXml(options, item.series, item.index, type))
      .join('');
    const labels = options.dataLabels?.lastPointOnly
      ? ''
      : labelsXml(options.dataLabels ?? (options.showValues ? {} : undefined), 0);
    let after =
      type === 'doughnut'
        ? `<c:firstSliceAng val="0"/><c:holeSize val="${options.holeSize ?? 50}"/>`
        : '';
    if ((type === 'column' || type === 'bar') && grouping !== 'clustered')
      after += '<c:overlap val="100"/>';
    if (type === 'stock')
      after += `<c:hiLowLines/>${options.series.length === 4 ? '<c:upDownBars><c:gapWidth val="150"/><c:upBars/><c:downBars/></c:upDownBars>' : ''}`;
    if (type !== 'pie' && type !== 'doughnut') {
      axes.add(secondary);
      after += `<c:axId val="${secondary ? 3 : 1}"/><c:axId val="${secondary ? 4 : 2}"/>`;
    }
    plot += `<c:${tag}>${before}${type === 'stock' ? '' : `<c:varyColors val="${(options.varyColors ?? (type === 'pie' || type === 'doughnut')) ? 1 : 0}"/>`}${series}${labels}${after}</c:${tag}>`;
  }
  for (const secondary of axes) {
    const x = secondary ? 3 : 1;
    const y = x + 1;
    const horizontal = options.type === 'bar';
    plot += axisXml(
      x,
      y,
      options.type === 'scatter' || options.type === 'bubble',
      horizontal ? 'l' : 'b',
      options.categoryAxis,
      secondary,
    );
    plot += axisXml(
      y,
      x,
      true,
      horizontal ? 'b' : secondary ? 'r' : 'l',
      secondary ? options.secondaryValueAxis : options.valueAxis,
      secondary,
    );
  }
  const legendPosition = {top: 't', bottom: 'b', left: 'l', right: 'r'} as const;
  const legend =
    options.legend === false
      ? ''
      : `<c:legend><c:legendPos val="${legendPosition[options.legend ?? 'bottom']}"/><c:layout/><c:overlay val="0"/></c:legend>`;
  return `${XML_DECLARATION}<c:chartSpace xmlns:c="${CHART_NS}" xmlns:a="${DRAWINGML_NS}"><c:chart>${options.title === undefined ? '' : title(options.title)}<c:autoTitleDeleted val="${options.title === undefined ? 1 : 0}"/><c:plotArea><c:layout/>${plot}</c:plotArea>${legend}<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/><c:showDLblsOverMax val="0"/></c:chart></c:chartSpace>`;
}
