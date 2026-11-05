import type {Worksheet} from '../../core/worksheet.ts';
import {tagRanges} from '../../xml/xml-read.ts';
import {localName} from '../../xml/xml-scan.ts';
import {escapeAttr, XML_DECLARATION} from '../../xml/xml.ts';
import {RELATIONSHIPS_NS} from '../opc/namespaces.ts';
import type {WorkbookMetadataTable} from './cell-metadata.ts';
import {PackageEditor} from './package-editor.ts';
import type {MediaPlan} from './package-plan.ts';

export const RICH_IMAGE_REL =
  'http://schemas.microsoft.com/office/2022/10/relationships/richValueRel';
export const RICH_IMAGE_PATH = 'xl/richData/richValueRel.xml';

/** Cells and all rich-value records are assigned together before the package metadata is emitted. */
export function cellImageMetadata(
  xml: string,
  sheet: Worksheet,
  metadata: WorkbookMetadataTable,
): string {
  const images = new Map(sheet.cellImages.map((image) => [image.ref, image]));
  if (images.size === 0) return xml;
  const edits: {start: number; end: number; text: string}[] = [];
  for (const tag of tagRanges(xml)) {
    if (tag.close || localName(tag.name) !== 'c') continue;
    const image = images.get(tag.attrs.r ?? '');
    if (!image) continue;
    const vm = metadata.markCellImage(image.imageId, image.alt);
    const attrs = {...tag.attrs, vm: String(vm)};
    edits.push({
      start: tag.start,
      end: tag.end,
      text: `<${tag.name}${Object.entries(attrs)
        .map(([key, value]) => ` ${key}="${escapeAttr(value)}"`)
        .join('')}${tag.selfClosing ? '/' : ''}>`,
    });
  }
  for (const edit of edits.reverse())
    xml = xml.slice(0, edit.start) + edit.text + xml.slice(edit.end);
  return xml;
}

export function appendCellImageParts(
  parts: Record<string, Uint8Array>,
  metadata: WorkbookMetadataTable,
  media: MediaPlan,
): void {
  const images = metadata.images;
  if (images.length === 0) return;
  const pkg = new PackageEditor(parts);
  const refs = images.map((image) => {
    const resolved = media.resolve(image.imageId);
    const id = pkg.relationship(
      RICH_IMAGE_PATH,
      `${RELATIONSHIPS_NS}/image`,
      `xl/media/image${resolved.number}.${resolved.image.extension}`,
    );
    return `<rel r:id="${id}"/>`;
  });
  pkg.write(
    RICH_IMAGE_PATH,
    `${XML_DECLARATION}<richValueRels xmlns="http://schemas.microsoft.com/office/spreadsheetml/2022/richvaluerel" xmlns:r="${RELATIONSHIPS_NS}">${refs.join('')}</richValueRels>`,
  );
  pkg.override(RICH_IMAGE_PATH, 'application/vnd.ms-excel.richvaluerel+xml');
  pkg.relationship('xl/workbook.xml', RICH_IMAGE_REL, RICH_IMAGE_PATH);
  const typesPath = 'xl/richData/rdRichValueTypes.xml';
  const keys = [
    '_DisplayString',
    '_Flags',
    '_Format',
    '_SubLabel',
    '_Attribution',
    '_Icon',
    '_Display',
    '_CanonicalPropertyNames',
    '_ClassificationId',
  ];
  pkg.write(
    typesPath,
    XML_DECLARATION +
      '<rvTypesInfo xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2"><global><keyFlags><key name="_Self"><flag name="ExcludeFromFile" value="1"/><flag name="ExcludeFromCalcComparison" value="1"/></key>' +
      keys
        .map((key) => `<key name="${key}"><flag name="ExcludeFromCalcComparison" value="1"/></key>`)
        .join('') +
      '</keyFlags></global></rvTypesInfo>',
  );
  pkg.override(typesPath, 'application/vnd.ms-excel.rdrichvaluetypes+xml');
  pkg.relationship(
    'xl/workbook.xml',
    'http://schemas.microsoft.com/office/2017/06/relationships/rdRichValueTypes',
    typesPath,
  );
}
