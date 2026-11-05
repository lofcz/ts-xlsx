import type {ChartData, ChartOptions} from '../../core/chart.ts';
import {escapeAttr, escapeText, XML_DECLARATION} from '../../xml/xml.ts';
import {DRAWINGML_NS} from './namespaces.ts';

export const CHARTEX_NS = 'http://schemas.microsoft.com/office/drawing/2014/chartex';
export const CHARTEX_REL = 'http://schemas.microsoft.com/office/2014/relationships/chartEx';
export const CHARTEX_CONTENT_TYPE = 'application/vnd.ms-office.chartex+xml';

function level(data: ChartData<string | number>, numeric: boolean): string {
  return `<cx:lvl ptCount="${data.values.length}"${numeric ? ` formatCode="${escapeAttr(data.formatCode ?? 'General')}"` : ''}>${data.values.map((value, index) => (value === null ? '' : `<cx:pt idx="${index}">${escapeText(String(value))}</cx:pt>`)).join('')}</cx:lvl>`;
}

function dimension(data: readonly ChartData<string | number>[], numeric: boolean): string {
  const first = data[0];
  return `<cx:${numeric ? 'num' : 'str'}Dim type="${numeric ? 'val' : 'cat'}">${first?.formula ? `<cx:f>${escapeText(first.formula.replace(/^=/, ''))}</cx:f>` : ''}${data.map((item) => level(item, numeric)).join('')}</cx:${numeric ? 'num' : 'str'}Dim>`;
}

/** MS-ODRAWXML chartEx keeps referenced data separate from the chart's series layout. */
export function chartExXml(options: ChartOptions): string {
  let data = '',
    series = '';
  options.series.forEach((item, index) => {
    const cats =
      item.categoryLevels ??
      (item.categories || options.categories
        ? [item.categories ?? options.categories].filter((value) => value !== undefined)
        : []);
    data += `<cx:data id="${index}">${cats.length ? dimension(cats, false) : ''}${dimension([item.values], true)}</cx:data>`;
    const layout =
      options.type === 'histogram' || options.type === 'pareto' ? 'clusteredColumn' : options.type;
    const color = item.color?.replace(/^#/, '').toUpperCase();
    const tx = `<cx:tx><cx:txData>${item.nameFormula ? `<cx:f>${escapeText(item.nameFormula)}</cx:f>` : ''}<cx:v>${escapeText(item.name)}</cx:v></cx:txData></cx:tx>`;
    const fill = color
      ? `<cx:spPr><a:solidFill><a:srgbClr val="${color}"/></a:solidFill></cx:spPr>`
      : '';
    const points =
      item.pointColors
        ?.map((point, i) =>
          point === null
            ? ''
            : `<cx:dataPt idx="${i}"><cx:spPr><a:solidFill><a:srgbClr val="${point.replace(/^#/, '').toUpperCase()}"/></a:solidFill></cx:spPr></cx:dataPt>`,
        )
        .join('') ?? '';
    const label = item.dataLabels ?? options.dataLabels ?? (options.showValues ? {} : undefined);
    const labels = label
      ? `<cx:dataLabels${label.position ? ` pos="${label.position}"` : ''}>${label.formatCode ? `<cx:numFmt formatCode="${escapeAttr(label.formatCode)}" sourceLinked="0"/>` : ''}<cx:visibility seriesName="${label.showSeriesName ? 1 : 0}" categoryName="${label.showCategory ? 1 : 0}" value="${label.showValue === false ? 0 : 1}"/>${label.separator === undefined ? '' : `<cx:separator>${escapeText(label.separator)}</cx:separator>`}</cx:dataLabels>`
      : '';
    let properties = '';
    if (options.type === 'waterfall')
      properties = `<cx:visibility connectorLines="${options.waterfallConnectors === false ? 0 : 1}"/>${options.waterfallTotals?.length ? `<cx:subtotals>${options.waterfallTotals.map((i) => `<cx:idx val="${i}"/>`).join('')}</cx:subtotals>` : ''}`;
    if (options.type === 'histogram' || options.type === 'pareto') {
      const bins = options.binning;
      properties =
        options.type === 'pareto' && cats.length
          ? '<cx:aggregation/>'
          : `<cx:binning intervalClosed="r"${bins?.underflow === undefined ? '' : ` underflow="${bins.underflow}"`}${bins?.overflow === undefined ? '' : ` overflow="${bins.overflow}"`}>${bins?.width === undefined ? (bins?.count === undefined ? '' : `<cx:binCount>${bins.count}</cx:binCount>`) : `<cx:binSize>${bins.width}</cx:binSize>`}</cx:binning>`;
    }
    if (options.type === 'boxWhisker') {
      const box = options.boxWhisker;
      properties = `<cx:visibility meanLine="${box?.showMeanLine ? 1 : 0}" meanMarker="${box?.showMeanMarker === false ? 0 : 1}" nonoutliers="${box?.showInnerPoints ? 1 : 0}" outliers="${box?.showOutlierPoints === false ? 0 : 1}"/><cx:statistics quartileMethod="${box?.quartileMethod === 'inclusive' ? 'inclusive' : 'exclusive'}"/>`;
    }
    series += `<cx:series layoutId="${layout}" uniqueId="{6D3A0000-0000-4000-8000-${(index + 1).toString(16).padStart(12, '0')}}">${tx}${fill}${points}${labels}<cx:dataId val="${index}"/>${properties ? `<cx:layoutPr>${properties}</cx:layoutPr>` : ''}${options.type === 'pareto' ? '<cx:axisId>1</cx:axisId>' : ''}</cx:series>`;
  });
  if (options.type === 'pareto')
    series += '<cx:series layoutId="paretoLine" ownerIdx="0"><cx:axisId>2</cx:axisId></cx:series>';
  const title = (value: string | undefined, axis = false) =>
    value
      ? `<cx:title${axis ? '' : ' pos="t" align="ctr" overlay="0"'}><cx:tx><cx:txData><cx:v>${escapeText(value)}</cx:v></cx:txData></cx:tx></cx:title>`
      : '';
  const scale = (value: number | undefined, key: string) =>
    value === undefined ? '' : ` ${key}="${value}"`;
  const numberFormat = (axis: ChartOptions['valueAxis']) =>
    axis?.formatCode
      ? `<cx:numFmt formatCode="${escapeAttr(axis.formatCode)}" sourceLinked="0"/>`
      : '';
  let axes = ['treemap', 'sunburst'].includes(options.type)
    ? ''
    : `<cx:axis id="0"><cx:catScaling gapWidth="0.5"/>${title(options.categoryAxis?.title, true)}<cx:tickLabels/>${numberFormat(options.categoryAxis)}</cx:axis><cx:axis id="1"${options.type === 'funnel' ? ' hidden="1"' : ''}><cx:valScaling${scale(options.valueAxis?.min, 'min')}${scale(options.valueAxis?.max, 'max')}${scale(options.valueAxis?.majorUnit, 'majorUnit')}${scale(options.valueAxis?.minorUnit, 'minorUnit')}/>${title(options.valueAxis?.title, true)}<cx:majorGridlines/><cx:tickLabels/>${numberFormat(options.valueAxis)}</cx:axis>`;
  if (options.type === 'pareto')
    axes +=
      '<cx:axis id="2"><cx:valScaling min="0" max="1"/><cx:units unit="percentage"/><cx:tickLabels/></cx:axis>';
  const positions = {top: 't', bottom: 'b', left: 'l', right: 'r'};
  const legend =
    options.legend === false
      ? ''
      : `<cx:legend pos="${positions[options.legend ?? 'bottom']}" align="ctr" overlay="0"/>`;
  return `${XML_DECLARATION}<cx:chartSpace xmlns:cx="${CHARTEX_NS}" xmlns:a="${DRAWINGML_NS}"><cx:chartData>${data}</cx:chartData><cx:chart>${title(options.title)}<cx:plotArea><cx:plotAreaRegion>${series}</cx:plotAreaRegion>${axes}</cx:plotArea>${legend}</cx:chart></cx:chartSpace>`;
}
