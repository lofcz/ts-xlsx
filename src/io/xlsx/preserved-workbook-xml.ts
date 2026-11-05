import {localName} from '../../xml/xml-scan.ts';
import {xmlChildren, xmlSerialize, xmlTree} from '../../xml/xml-tree.ts';
import {XML_DECLARATION} from '../../xml/xml.ts';
import {SLICER_CACHES_EXT_URI} from './namespaces.ts';
import {transitionalPart} from './strict-parts.ts';

const order = [
  'fileVersion',
  'fileSharing',
  'workbookPr',
  'workbookProtection',
  'bookViews',
  'sheets',
  'functionGroups',
  'externalReferences',
  'definedNames',
  'calcPr',
  'oleSize',
  'customWorkbookViews',
  'pivotCaches',
  'smartTagPr',
  'smartTagTypes',
  'webPublishing',
  'fileRecoveryPr',
  'webPublishObjects',
  'extLst',
];
const modeled = new Set([
  'workbookProtection',
  'bookViews',
  'sheets',
  'externalReferences',
  'definedNames',
  'pivotCaches',
]);

export function captureWorkbookXml(xml: string): string[] {
  xml = new TextDecoder().decode(
    transitionalPart(new TextEncoder().encode(xml), 'application/xml'),
  );
  const root = xmlTree(xml);
  const namespaces = Object.fromEntries(
    Object.entries(root.attrs).filter(([key]) => key.startsWith('xmlns')),
  );
  const fragments = xmlChildren(root).flatMap((node) => {
    const name = localName(node.name);
    if (modeled.has(name)) return [];
    if (name === 'workbookPr') for (const key of ['date1904', 'codeName']) delete node.attrs[key];
    if (name === 'calcPr') delete node.attrs.fullCalcOnLoad;
    if (name === 'extLst')
      node.children = xmlChildren(node).filter(
        (child) => child.attrs.uri?.toUpperCase() !== SLICER_CACHES_EXT_URI.toUpperCase(),
      );
    if (Object.keys(node.attrs).length === 0 && node.children.length === 0) return [];
    Object.assign(node.attrs, namespaces, node.attrs);
    return [xmlSerialize(node)];
  });
  return [xmlSerialize({...root, children: []}), ...fragments];
}

export function restoreWorkbookXml(
  xml: string,
  fragments: readonly string[],
  relationshipIds: ReadonlyMap<string, string>,
): string {
  if (!fragments.length) return xml;
  const root = xmlTree(xml);
  const rewrite = (node: ReturnType<typeof xmlTree>) => {
    for (const [key, value] of Object.entries(node.attrs))
      if (key.startsWith('r:') && relationshipIds.has(value))
        node.attrs[key] = relationshipIds.get(value) ?? value;
    for (const child of xmlChildren(node)) rewrite(child);
  };
  for (const fragment of fragments) {
    const node = xmlTree(fragment),
      name = localName(node.name);
    rewrite(node);
    if (name === 'workbook') {
      root.attrs = {...node.attrs, ...root.attrs};
      continue;
    }
    const existing = xmlChildren(root).find((child) => localName(child.name) === name);
    if (existing) {
      existing.attrs = {...node.attrs, ...existing.attrs};
      if (name === 'extLst') existing.children.push(...node.children);
    } else {
      const index = order.indexOf(name);
      const following = root.children.findIndex(
        (child) => typeof child !== 'string' && order.indexOf(localName(child.name)) > index,
      );
      root.children.splice(following < 0 ? root.children.length : following, 0, node);
    }
  }
  return XML_DECLARATION + xmlSerialize(root);
}
