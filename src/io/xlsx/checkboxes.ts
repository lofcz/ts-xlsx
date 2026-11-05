import type {Worksheet} from '../../core/worksheet.ts';
import {AuthoringError} from '../../errors.ts';
import {tagRanges} from '../../xml/xml-read.ts';
import {localName} from '../../xml/xml-scan.ts';
import {xmlChild, xmlChildren, xmlTree, xmlText, type XmlNode} from '../../xml/xml-tree.ts';
import {escapeAttr, XML_DECLARATION} from '../../xml/xml.ts';
import {PackageEditor} from './package-editor.ts';

export const CHECKBOX_REL =
  'http://schemas.microsoft.com/office/2022/11/relationships/FeaturePropertyBag';
const ns = 'http://schemas.microsoft.com/office/spreadsheetml/2022/featurepropertybag';
const uri = '{C7286773-470A-42A8-94C5-96B5CB345126}';
const path = 'xl/featurePropertyBag/featurePropertyBag.xml';

export function appendCheckboxParts(
  parts: Record<string, Uint8Array>,
  sheets: readonly Worksheet[],
): void {
  if (!sheets.some((sheet) => sheet.checkboxes.length)) return;
  const pkg = new PackageEditor(parts),
    styles = pkg.tree('xl/styles.xml'),
    cellXfs = xmlChild(styles, 'cellXfs');
  if (!cellXfs) throw new AuthoringError('checkbox styles have no cellXfs');
  const xfs = xmlChildren(cellXfs, 'xf'),
    clones = new Map<number, number>();
  const clone = (index: number) => {
    const existing = clones.get(index);
    if (existing !== undefined) return existing;
    const source = xfs[index];
    if (!source) throw new AuthoringError('checkbox cell style index is missing');
    const xf = structuredClone(source),
      extension = xmlTree(`<ext uri="${uri}" xmlns:xfpb="${ns}"><xfpb:xfComplement i="0"/></ext>`);
    const ext = xmlChild(xf, 'extLst');
    if (ext) ext.children.push(extension);
    else xf.children.push({name: 'extLst', attrs: {}, children: [extension]});
    const next = xmlChildren(cellXfs, 'xf').length;
    cellXfs.children.push(xf);
    clones.set(index, next);
    return next;
  };
  for (const [index, sheet] of sheets.entries()) {
    const cells = new Set(sheet.checkboxes);
    if (cells.size === 0) continue;
    const sheetPath = `xl/worksheets/sheet${index + 1}.xml`;
    let xml = pkg.read(sheetPath);
    const edits: {start: number; end: number; xml: string}[] = [];
    for (const tag of tagRanges(xml)) {
      if (tag.close || localName(tag.name) !== 'c' || !cells.has(tag.attrs.r ?? '')) continue;
      const attrs = {...tag.attrs, s: String(clone(Number(tag.attrs.s ?? 0)))};
      edits.push({
        start: tag.start,
        end: tag.end,
        xml: `<${tag.name}${Object.entries(attrs)
          .map(([key, value]) => ` ${key}="${escapeAttr(value)}"`)
          .join('')}${tag.selfClosing ? '/' : ''}>`,
      });
    }
    for (const edit of edits.reverse())
      xml = xml.slice(0, edit.start) + edit.xml + xml.slice(edit.end);
    pkg.write(sheetPath, xml);
  }
  cellXfs.attrs.count = String(xmlChildren(cellXfs, 'xf').length);
  pkg.writeTree('xl/styles.xml', styles);
  pkg.write(
    path,
    `${XML_DECLARATION}<FeaturePropertyBags xmlns="${ns}"><bag type="Checkbox"/><bag type="XFControls"><bagId k="CellControl">0</bagId></bag><bag type="XFComplement"><bagId k="XFControls">1</bagId></bag><bag type="XFComplements" extRef="XFComplementsMapperExtRef"><a k="MappedFeaturePropertyBags"><bagId>2</bagId></a></bag></FeaturePropertyBags>`,
  );
  pkg.override(path, 'application/vnd.ms-excel.featurepropertybag+xml');
  pkg.relationship('xl/workbook.xml', CHECKBOX_REL, path);
}

function descendants(node: XmlNode, name: string): XmlNode[] {
  return xmlChildren(node).flatMap((child) => [
    ...(localName(child.name) === name ? [child] : []),
    ...descendants(child, name),
  ]);
}

export function checkboxStyleIds(stylesXml: string, bagsXml: string): Set<number> {
  const result = new Set<number>();
  if (!stylesXml || !bagsXml) return result;
  const bags = xmlChildren(xmlTree(bagsXml), 'bag'),
    mapper = bags.find((bag) => bag.attrs.type === 'XFComplements');
  const mapped = mapper ? descendants(mapper, 'bagId').map((item) => Number(xmlText(item))) : [];
  const reaches = (start: number) => {
    const seen = new Set<number>(),
      queue = [start];
    while (queue.length) {
      const index = queue.pop();
      if (index === undefined || seen.has(index)) continue;
      seen.add(index);
      const bag = bags[index];
      if (!bag) continue;
      if (bag.attrs.type === 'Checkbox') return true;
      queue.push(...descendants(bag, 'bagId').map((item) => Number(xmlText(item))));
    }
    return false;
  };
  const xfs = xmlChild(xmlTree(stylesXml), 'cellXfs');
  if (!xfs) return result;
  xmlChildren(xfs, 'xf').forEach((xf, index) => {
    const complement = descendants(xf, 'xfComplement')[0],
      mappedIndex = complement ? mapped[Number(complement.attrs.i)] : undefined;
    if (mappedIndex !== undefined && reaches(mappedIndex)) result.add(index);
  });
  return result;
}
