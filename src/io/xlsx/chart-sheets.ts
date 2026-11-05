import type {Worksheet} from '../../core/worksheet.ts';
import {AuthoringError} from '../../errors.ts';
import {xmlChild, xmlChildren, xmlSerialize, xmlTree} from '../../xml/xml-tree.ts';
import {XML_DECLARATION} from '../../xml/xml.ts';
import {RELATIONSHIPS_NS} from '../opc/namespaces.ts';
import {relsPathFor, resolveRelativePart} from '../opc/part-paths.ts';

/** Chart sheets share the chart/drawing allocator, then replace the grid-shaped sheet container. */
export function appendChartSheets(
  parts: Record<string, Uint8Array>,
  sheets: readonly Worksheet[],
): void {
  const decoder = new TextDecoder(),
    encoder = new TextEncoder();
  const read = (path: string) => xmlTree(decoder.decode(parts[path]));
  const write = (path: string, xml: string) => {
    parts[path] = encoder.encode(xml);
  };
  for (const [index, sheet] of sheets.entries()) {
    if (sheet.kind !== 'chartsheet') continue;
    if (sheet.rowCount > 0 || sheet.charts.length !== 1)
      throw new AuthoringError('a chart sheet must contain exactly one chart and no cells');
    const oldPath = `xl/worksheets/sheet${index + 1}.xml`,
      path = `xl/chartsheets/sheet${index + 1}.xml`;
    const old = read(oldPath),
      drawing = xmlChild(old, 'drawing');
    if (!drawing) throw new AuthoringError('chart sheet drawing is missing');
    const margins = xmlChild(old, 'pageMargins'),
      setup = xmlChild(old, 'pageSetup');
    const xml = `${XML_DECLARATION}<chartsheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="${RELATIONSHIPS_NS}"><sheetPr/><sheetViews><sheetView workbookViewId="0" zoomToFit="1"/></sheetViews>${margins ? xmlSerialize(margins) : ''}${setup ? xmlSerialize(setup) : ''}${xmlSerialize(drawing)}</chartsheet>`;
    write(path, xml);
    const rels = parts[relsPathFor(oldPath)];
    if (rels) parts[relsPathFor(path)] = rels;
    if (rels) {
      const drawingRel = xmlChildren(xmlTree(decoder.decode(rels)), 'Relationship').find(
        (rel) => rel.attrs.Type === `${RELATIONSHIPS_NS}/drawing`,
      );
      if (drawingRel?.attrs.Target) {
        const drawingPath = resolveRelativePart(oldPath, drawingRel.attrs.Target),
          tree = read(drawingPath);
        for (const anchor of xmlChildren(tree)) {
          if (!xmlChild(anchor, 'graphicFrame') && !xmlChild(anchor, 'AlternateContent')) continue;
          const extent = xmlChild(anchor, 'ext') ?? {
            name: 'xdr:ext',
            attrs: {cx: '9144000', cy: '5715000'},
            children: [],
          };
          anchor.name = 'xdr:absoluteAnchor';
          delete anchor.attrs.editAs;
          anchor.children = [
            {name: 'xdr:pos', attrs: {x: '0', y: '0'}, children: []},
            extent,
            ...xmlChildren(anchor).filter(
              (child) => !['from', 'to', 'ext'].includes(child.name.split(':').at(-1) ?? ''),
            ),
          ];
        }
        write(drawingPath, XML_DECLARATION + xmlSerialize(tree));
      }
    }
    delete parts[oldPath];
    delete parts[relsPathFor(oldPath)];
    const workbookRels = read('xl/_rels/workbook.xml.rels');
    for (const rel of xmlChildren(workbookRels, 'Relationship'))
      if (rel.attrs.Target === `worksheets/sheet${index + 1}.xml`) {
        rel.attrs.Target = `chartsheets/sheet${index + 1}.xml`;
        rel.attrs.Type = `${RELATIONSHIPS_NS}/chartsheet`;
      }
    write('xl/_rels/workbook.xml.rels', XML_DECLARATION + xmlSerialize(workbookRels));
    const types = read('[Content_Types].xml');
    for (const type of xmlChildren(types, 'Override'))
      if (type.attrs.PartName === `/${oldPath}`) {
        type.attrs.PartName = `/${path}`;
        type.attrs.ContentType =
          'application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml';
      }
    write('[Content_Types].xml', XML_DECLARATION + xmlSerialize(types));
  }
}
