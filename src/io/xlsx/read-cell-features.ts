import {INTERNAL} from '../../core/internal.ts';
import type {Worksheet} from '../../core/worksheet.ts';
import {openElements} from '../../xml/xml-read.ts';
import {xmlChildren, xmlText, xmlTree} from '../../xml/xml-tree.ts';
import {extensionOf, resolveRelativePart} from '../opc/part-paths.ts';
import {type PartRelationships, readPartRelationships} from '../opc/read-opc.ts';
import {parseMetadataPart} from './cell-metadata.ts';
import {checkboxStyleIds} from './checkboxes.ts';
import type {PackageReadContext} from './read-parts.ts';
import {parseSparklines} from './sparklines.ts';

export function readCellFeatures(
  context: PackageReadContext,
  workbookRels: PartRelationships,
  xml: string,
  sheet: Worksheet,
): void {
  sheet[INTERNAL].sparklines.push(...parseSparklines(xml));
  const formats = checkboxStyleIds(
    workbookRels.relatedText('styles') ?? '',
    workbookRels.relatedText('FeaturePropertyBag') ?? '',
  );
  const structureXml = workbookRels.relatedText('rdRichValueStructure'),
    valuesXml = workbookRels.relatedText('rdRichValue');
  const metadata = parseMetadataPart(workbookRels.relatedText('sheetMetadata') ?? '');
  const structures = structureXml ? xmlChildren(xmlTree(structureXml), 's') : [],
    values = valuesXml ? xmlChildren(xmlTree(valuesXml), 'rv') : [];
  const relationPath = workbookRels.targetPath('richValueRel'),
    relationXml = relationPath ? context.pkg.partText(relationPath) : undefined;
  const relations = relationXml ? xmlChildren(xmlTree(relationXml), 'rel') : [],
    imageRels = relationPath
      ? readPartRelationships(relationPath, context.pkg.partText)
      : undefined;
  for (const cell of openElements(xml, 'c')) {
    const ref = cell.attrs.r;
    if (!ref) continue;
    if (formats.has(Number(cell.attrs.s))) sheet[INTERNAL].checkboxes.push(ref);
    const vm = Number(cell.attrs.vm);
    if (!Number.isInteger(vm) || vm <= 0) continue;
    const record = metadata.valueBlocks[vm - 1]?.find(
      (item) => metadata.typeNames[item.type - 1] === 'XLRICHVALUE',
    );
    const richIndex = record ? metadata.richValueBlocks[record.value] : undefined,
      rich = richIndex === undefined ? undefined : values[richIndex];
    if (!rich) continue;
    const structure = structures[Number(rich.attrs.s)];
    if (structure?.attrs.t !== '_localImage') continue;
    const keys = xmlChildren(structure, 'k'),
      data = xmlChildren(rich, 'v').map(xmlText);
    const imageIndex = Number(
      data[keys.findIndex((key) => key.attrs.n === '_rvRel:LocalImageIdentifier')],
    );
    const relation = relations[imageIndex],
      id = relation?.attrs['r:id'];
    const target = id ? imageRels?.byId(id) : undefined;
    if (!target || target.external || !relationPath) continue;
    const imagePath = resolveRelativePart(relationPath, target.target),
      bytes = context.pkg.partBytes(imagePath),
      extension = extensionOf(imagePath);
    if (!bytes || !extension) continue;
    const imageId = context.workbook.addImage({buffer: bytes, extension});
    const alt = data[keys.findIndex((key) => key.attrs.n === 'Text')] ?? '';
    sheet[INTERNAL].cellImages.push({ref, imageId, alt});
  }
}
