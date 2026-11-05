import {directChildRange, elementRange, tagRanges} from '../../xml/xml-read.ts';
import {localName} from '../../xml/xml-scan.ts';
import {xmlChildren, xmlSerialize, xmlTree, type XmlNode} from '../../xml/xml-tree.ts';
import {startTag} from '../../xml/xml.ts';
import {CF_EXT_URI, DATA_VALIDATION_EXT_URI, SLICER_LIST_EXT_URI} from './namespaces.ts';
import {SPARKLINE_URI} from './sparklines.ts';
import {transitionalPart} from './strict-parts.ts';

const sheetOrder = [
  'sheetPr',
  'dimension',
  'sheetViews',
  'sheetFormatPr',
  'cols',
  'sheetData',
  'sheetCalcPr',
  'sheetProtection',
  'protectedRanges',
  'scenarios',
  'autoFilter',
  'sortState',
  'dataConsolidate',
  'customSheetViews',
  'mergeCells',
  'phoneticPr',
  'conditionalFormatting',
  'dataValidations',
  'hyperlinks',
  'printOptions',
  'pageMargins',
  'pageSetup',
  'headerFooter',
  'rowBreaks',
  'colBreaks',
  'customProperties',
  'cellWatches',
  'ignoredErrors',
  'smartTags',
  'drawing',
  'legacyDrawing',
  'legacyDrawingHF',
  'picture',
  'oleObjects',
  'controls',
  'webPublishItems',
  'tableParts',
  'extLst',
];
const modeled = new Set([
  'sheetPr',
  'dimension',
  'sheetViews',
  'sheetFormatPr',
  'cols',
  'sheetData',
  'sheetProtection',
  'autoFilter',
  'mergeCells',
  'conditionalFormatting',
  'dataValidations',
  'hyperlinks',
  'printOptions',
  'pageMargins',
  'pageSetup',
  'headerFooter',
  'rowBreaks',
  'colBreaks',
  'drawing',
  'legacyDrawing',
  'legacyDrawingHF',
  'picture',
  'tableParts',
]);
const ownedExtensions = new Set(
  [CF_EXT_URI, DATA_VALIDATION_EXT_URI, SLICER_LIST_EXT_URI, SPARKLINE_URI].map((uri) =>
    uri.toUpperCase(),
  ),
);

/** Only small metadata fragments become trees; worksheet cell grids remain streamed text. */
export function captureWorksheetXml(xml: string): string[] {
  xml = new TextDecoder().decode(
    transitionalPart(new TextEncoder().encode(xml), 'application/xml'),
  );
  const root = elementRange(xml, ['worksheet']);
  if (!root) return [];
  const namespaces = Object.fromEntries(
    Object.entries(root.attrs).filter(([key]) => key === 'xmlns' || key.startsWith('xmlns:')),
  );
  let depth = 0,
    start = 0;
  const fragments: string[] = [
    xmlSerialize({name: root.name, attrs: {...root.attrs}, children: []}),
  ];
  for (const tag of tagRanges(xml)) {
    if (!tag.close) {
      if (depth === 1) start = tag.start;
      if (!tag.selfClosing) depth++;
    } else depth--;
    if (depth !== 1 || (!tag.close && !tag.selfClosing) || tag.start < root.contentStart) continue;
    const name = localName(tag.name);
    if (modeled.has(name)) continue;
    const node = xmlTree(xml.slice(start, tag.end));
    Object.assign(node.attrs, namespaces, node.attrs);
    if (name === 'extLst')
      node.children = xmlChildren(node, 'ext').filter(
        (ext) => !ownedExtensions.has((ext.attrs.uri ?? '').toUpperCase()),
      );
    if (name !== 'extLst' || node.children.length) fragments.push(xmlSerialize(node));
  }
  return fragments;
}

export function restoreWorksheetXml(
  xml: string,
  fragments: readonly string[],
  relationshipIds: ReadonlyMap<string, string>,
): string {
  const rewrite = (node: XmlNode): void => {
    for (const [key, value] of Object.entries(node.attrs))
      if (key.endsWith(':id') || key.endsWith(':embed') || key.endsWith(':link')) {
        const replacement = relationshipIds.get(value);
        if (replacement) node.attrs[key] = replacement;
      }
    for (const child of xmlChildren(node)) rewrite(child);
  };
  for (const fragment of fragments) {
    const node = xmlTree(fragment);
    rewrite(node);
    const name = localName(node.name);
    if (name === 'worksheet') {
      const root = elementRange(xml, ['worksheet']);
      if (root)
        xml =
          xml.slice(0, root.start) +
          startTag(root.name, {...node.attrs, ...root.attrs}, false) +
          xml.slice(root.contentStart);
      continue;
    }
    if (name === 'extLst') {
      const existing = directChildRange(xml, 'worksheet', 'extLst');
      if (existing) {
        const target = xmlTree(xml.slice(existing.start, existing.end));
        target.children.push(...node.children);
        Object.assign(target.attrs, node.attrs, target.attrs);
        xml = xml.slice(0, existing.start) + xmlSerialize(target) + xml.slice(existing.end);
        continue;
      }
    }
    const index = sheetOrder.indexOf(name),
      next = sheetOrder
        .slice(index + 1)
        .map((key) => directChildRange(xml, 'worksheet', key))
        .find((range) => range !== undefined);
    const root = elementRange(xml, ['worksheet']);
    if (!root) continue;
    const at = next?.start ?? root.contentEnd;
    xml = xml.slice(0, at) + xmlSerialize(node) + xml.slice(at);
  }
  return xml;
}
