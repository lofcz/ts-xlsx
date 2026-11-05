import type {SparklineGroup} from '../../core/sheet-features.ts';
import {AuthoringError} from '../../errors.ts';
import {directChildRange, elementRange} from '../../xml/xml-read.ts';
import {xmlChild, xmlChildren, xmlText, xmlTree} from '../../xml/xml-tree.ts';
import {escapeAttr, escapeText} from '../../xml/xml.ts';

export const SPARKLINE_URI = '{05C60535-1F16-4fd2-B633-F4F36F0B64E0}';
const flags = [
  'markers',
  'high',
  'low',
  'first',
  'last',
  'negative',
  'displayXAxis',
  'displayHidden',
  'rightToLeft',
] as const;
const colorNames = {
  series: 'colorSeries',
  negative: 'colorNegative',
  axis: 'colorAxis',
  markers: 'colorMarkers',
  first: 'colorFirst',
  last: 'colorLast',
  high: 'colorHigh',
  low: 'colorLow',
} as const;

export function sparklineExtXml(groups: readonly SparklineGroup[]): string {
  if (groups.length === 0) return '';
  const body = groups
    .map((group) => {
      let attrs = ` type="${group.type === 'winloss' ? 'stacked' : (group.type ?? 'line')}" displayEmptyCellsAs="${group.displayEmptyCellsAs ?? 'gap'}"`;
      for (const flag of flags)
        if (group[flag] !== undefined) attrs += ` ${flag}="${group[flag] ? 1 : 0}"`;
      for (const key of [
        'lineWeight',
        'manualMin',
        'manualMax',
        'minAxisType',
        'maxAxisType',
      ] as const)
        if (group[key] !== undefined) attrs += ` ${key}="${escapeAttr(String(group[key]))}"`;
      if (group.dateAxis) attrs += ' dateAxis="1"';
      const colors = Object.entries(colorNames)
        .map(([key, tag]) => {
          const value = group.colors?.[key as keyof typeof colorNames];
          if (value === undefined) return '';
          if (!/^[\da-f]{8}$/i.test(value))
            throw new AuthoringError('sparkline colors must be eight hexadecimal ARGB digits');
          return `<x14:${tag} rgb="${value.toUpperCase()}"/>`;
        })
        .join('');
      return `<x14:sparklineGroup${attrs}>${colors}${group.dateAxis ? `<xm:f>${escapeText(group.dateAxis)}</xm:f>` : ''}<x14:sparklines>${group.sparklines.map((item) => `<x14:sparkline><xm:f>${escapeText(item.formula)}</xm:f><xm:sqref>${escapeText(item.ref)}</xm:sqref></x14:sparkline>`).join('')}</x14:sparklines></x14:sparklineGroup>`;
    })
    .join('');
  return `<ext uri="${SPARKLINE_URI}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:sparklineGroups xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">${body}</x14:sparklineGroups></ext>`;
}

export function appendSparklineXml(xml: string, groups: readonly SparklineGroup[]): string {
  const ext = sparklineExtXml(groups);
  if (!ext) return xml;
  const range = directChildRange(xml, 'worksheet', 'extLst');
  if (range) {
    if (xml.slice(range.start, range.contentStart).endsWith('/>'))
      return (
        xml.slice(0, range.contentStart - 2) + '>' + ext + `</${range.name}>` + xml.slice(range.end)
      );
    return xml.slice(0, range.contentEnd) + ext + xml.slice(range.contentEnd);
  }
  const root = elementRange(xml, ['worksheet']);
  if (!root) throw new AuthoringError('sparkline worksheet root is missing');
  return xml.slice(0, root.contentEnd) + `<extLst>${ext}</extLst>` + xml.slice(root.contentEnd);
}

export function parseSparklines(xml: string): SparklineGroup[] {
  const range = directChildRange(xml, 'worksheet', 'extLst');
  if (!range) return [];
  const root = xmlTree(xml.slice(range.start, range.end));
  const ext = xmlChildren(root, 'ext').find(
    (item) => item.attrs.uri?.toUpperCase() === SPARKLINE_URI.toUpperCase(),
  );
  const groups = xmlChild(ext, 'sparklineGroups');
  if (!groups) return [];
  return xmlChildren(groups, 'sparklineGroup')
    .map((group) => {
      const attrs = group.attrs,
        lines = xmlChild(group, 'sparklines');
      const result: {-readonly [K in keyof SparklineGroup]: SparklineGroup[K]} = {
        sparklines: lines
          ? xmlChildren(lines, 'sparkline').map((line) => ({
              ref: xmlText(xmlChild(line, 'sqref')),
              formula: xmlText(xmlChild(line, 'f')),
            }))
          : [],
      };
      result.type =
        attrs.type === 'stacked' ? 'winloss' : attrs.type === 'column' ? 'column' : 'line';
      for (const flag of flags)
        if (attrs[flag] !== undefined) result[flag] = attrs[flag] === '1' || attrs[flag] === 'true';
      for (const key of ['lineWeight', 'manualMin', 'manualMax'] as const) {
        const value = attrs[key];
        if (value !== undefined && Number.isFinite(Number(value))) result[key] = Number(value);
      }
      for (const key of ['minAxisType', 'maxAxisType'] as const) {
        const value = attrs[key];
        if (value === 'individual' || value === 'group' || value === 'custom') result[key] = value;
      }
      const empty = attrs.displayEmptyCellsAs;
      if (empty === 'gap' || empty === 'zero' || empty === 'span')
        result.displayEmptyCellsAs = empty;
      const date = xmlText(xmlChild(group, 'f'));
      if (date) result.dateAxis = date;
      const colors: Partial<Record<keyof typeof colorNames, string>> = {};
      for (const key of Object.keys(colorNames) as (keyof typeof colorNames)[]) {
        const color = xmlChild(group, colorNames[key])?.attrs.rgb;
        if (color) colors[key] = color;
      }
      if (Object.keys(colors).length) result.colors = colors;
      return result;
    })
    .filter((group) => group.sparklines.length > 0);
}
