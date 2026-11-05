import type {ChartPart} from '../../core/chart.ts';
import {type AnchorPoint, type ImageAnchor, isOneCellAnchor} from '../../core/image.ts';
import type {Worksheet} from '../../core/worksheet.ts';
import {AuthoringError} from '../../errors.ts';
import {elementRange, openElements, parseXml} from '../../xml/xml-read.ts';
import {escapeAttr, XML_DECLARATION} from '../../xml/xml.ts';
import {RELATIONSHIPS_NS} from '../opc/namespaces.ts';
import {relativePartPath, relsPathFor, resolveRelativePart} from '../opc/part-paths.ts';
import {parseRelationshipRecords} from '../opc/relationship-records.ts';
import {relsPartXml} from '../opc/rels.ts';
import {CHART_NS, chartXml} from './chart-xml.ts';
import {DRAWINGML_NS, XDR_NS} from './namespaces.ts';
import {RelationshipLedger} from './package-plan.ts';
import {splicePreservedPart} from './preserved-splices.ts';

const CHART_REL = `${RELATIONSHIPS_NS}/chart`;
const DRAWING_REL = `${RELATIONSHIPS_NS}/drawing`;
const CHART_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml';
const DRAWING_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.drawing+xml';
const encoder = new TextEncoder();
const decoder = new TextDecoder();

function appendChild(xml: string, local: string, child: string): string {
  const root = elementRange(xml, [local]);
  if (root === undefined) throw new AuthoringError(`chart package part has no ${local} root`);
  if (xml.slice(root.start, root.contentStart).endsWith('/>')) {
    return (
      xml.slice(0, root.contentStart - 2) + '>' + child + `</${root.name}>` + xml.slice(root.end)
    );
  }
  return xml.slice(0, root.contentEnd) + child + xml.slice(root.contentEnd);
}

function pointXml(point: AnchorPoint): string {
  const values = [point.col, point.colOff ?? 0, point.row, point.rowOff ?? 0];
  if (
    values.some((value) => !Number.isSafeInteger(value) || value < 0) ||
    point.col > 16383 ||
    point.row > 1048575
  ) {
    throw new AuthoringError(
      'chart anchor coordinates must be nonnegative integers within the worksheet',
    );
  }
  return `<xdr:col>${point.col}</xdr:col><xdr:colOff>${point.colOff ?? 0}</xdr:colOff><xdr:row>${point.row}</xdr:row><xdr:rowOff>${point.rowOff ?? 0}</xdr:rowOff>`;
}

function anchorXml(anchor: ImageAnchor, relationshipId: string, id: number, name: string): string {
  const namespaces = ` xmlns:xdr="${XDR_NS}" xmlns:a="${DRAWINGML_NS}" xmlns:r="${RELATIONSHIPS_NS}" xmlns:c="${CHART_NS}"`;
  const frame = `<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${id}" name="${escapeAttr(name)}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="${CHART_NS}"><c:chart r:id="${escapeAttr(relationshipId)}"/></a:graphicData></a:graphic></xdr:graphicFrame>`;
  const from = `<xdr:from>${pointXml(anchor.from)}</xdr:from>`;
  if (isOneCellAnchor(anchor)) {
    if (![anchor.ext.cx, anchor.ext.cy].every((value) => Number.isSafeInteger(value) && value > 0))
      throw new AuthoringError('chart extents must be positive integer EMUs');
    return `<xdr:oneCellAnchor${namespaces}>${from}<xdr:ext cx="${anchor.ext.cx}" cy="${anchor.ext.cy}"/>${frame}<xdr:clientData/></xdr:oneCellAnchor>`;
  }
  const editAs = anchor.editAs ?? 'twoCell';
  if (!['twoCell', 'oneCell', 'absolute'].includes(editAs))
    throw new AuthoringError('invalid chart anchor edit mode');
  return `<xdr:twoCellAnchor${namespaces} editAs="${editAs}">${from}<xdr:to>${pointXml(anchor.to)}</xdr:to>${frame}<xdr:clientData/></xdr:twoCellAnchor>`;
}

function assertChartPart(part: ChartPart): void {
  parseXml(part.xml, {onOpen() {}});
  const root = elementRange(part.xml, ['chartSpace']);
  if (root === undefined) throw new AuthoringError('chart XML must have a chartSpace root');
  const prefix = root.name.includes(':') ? `xmlns:${root.name.split(':')[0]}` : 'xmlns';
  if (root.attrs[prefix] !== CHART_NS)
    throw new AuthoringError('chart XML must use the standard DrawingML chart namespace');
  const seen = new Set<string>();
  for (const related of part.relatedParts ?? []) {
    if (!/^[A-Za-z_][\w.-]*$/.test(related.relationshipId) || seen.has(related.relationshipId))
      throw new AuthoringError('chart relationship ids must be unique XML names');
    seen.add(related.relationshipId);
    parseXml(related.xml, {onOpen() {}});
  }
}

/**
 * Add authored charts after generated and preserved parts have their final paths.
 * One drawing per sheet is an OOXML invariant: images, loaded graphics and new charts share it.
 */
export function appendChartParts(
  files: Record<string, Uint8Array>,
  sheets: readonly Worksheet[],
): void {
  const read = (path: string): string => decoder.decode(files[path]);
  const write = (path: string, xml: string): void => {
    files[path] = encoder.encode(xml);
  };
  const allocate = (directory: string, base: string): string => {
    let number = 1;
    while (Object.hasOwn(files, `${directory}/${base}${number}.xml`)) number++;
    return `${directory}/${base}${number}.xml`;
  };
  const override = (path: string, contentType: string): void => {
    write(
      '[Content_Types].xml',
      appendChild(
        read('[Content_Types].xml'),
        'Types',
        `<Override PartName="/${escapeAttr(path)}" ContentType="${escapeAttr(contentType)}"/>`,
      ),
    );
  };
  sheets.forEach((sheet, index) => {
    const charts = sheet.charts;
    if (charts.length === 0) return;
    const sheetPath = `xl/worksheets/sheet${index + 1}.xml`;
    const sheetRelsPath = relsPathFor(sheetPath);
    const sheetRels = parseRelationshipRecords(read(sheetRelsPath));
    const drawing = sheetRels.find((rel) => rel.type === DRAWING_REL && !rel.external);
    let drawingPath: string;
    if (drawing === undefined) {
      drawingPath = allocate('xl/drawings', 'drawing');
      write(drawingPath, `${XML_DECLARATION}<xdr:wsDr xmlns:xdr="${XDR_NS}"/>`);
      override(drawingPath, DRAWING_CONTENT_TYPE);
      const ledger = new RelationshipLedger(sheetPath, new Set(sheetRels.map((rel) => rel.id)));
      const id = ledger.add(DRAWING_REL, drawingPath);
      write(sheetRelsPath, relsPartXml([...sheetRels, ...ledger.relationships]));
      const xml = read(sheetPath);
      const element = `<drawing xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="${RELATIONSHIPS_NS}" r:id="${id}"/>`;
      const after = new Set([
        'legacyDrawing',
        'legacyDrawingHF',
        'picture',
        'oleObjects',
        'controls',
        'webPublishItems',
        'tableParts',
        'extLst',
      ]);
      const next = [...after]
        .map((name) => elementRange(xml, ['worksheet', name]))
        .filter((range) => range !== undefined)
        .sort((a, b) => a.start - b.start)[0];
      // The worksheet writer's trailing drawing/table elements are schema ordered already.
      const offset = next?.start ?? elementRange(xml, ['worksheet'])?.contentEnd;
      if (offset === undefined) throw new AuthoringError('chart worksheet has no worksheet root');
      write(sheetPath, xml.slice(0, offset) + element + xml.slice(offset));
    } else {
      drawingPath = resolveRelativePart(sheetPath, drawing.target);
      if (files[drawingPath] === undefined)
        throw new AuthoringError('chart worksheet drawing target is missing');
    }
    const drawingRelsPath = relsPathFor(drawingPath);
    const existingRels = parseRelationshipRecords(read(drawingRelsPath));
    const ledger = new RelationshipLedger(drawingPath, new Set(existingRels.map((rel) => rel.id)));
    let drawingXml = read(drawingPath);
    let shapeId = 0;
    for (const tag of openElements(drawingXml, 'cNvPr'))
      shapeId = Math.max(shapeId, Number(tag.attrs.id) || 0);
    for (const chart of charts) {
      const part: ChartPart =
        chart.kind === 'chart'
          ? {
              xml: chartXml(chart.options),
              ...(chart.options.title === undefined ? {} : {name: chart.options.title}),
            }
          : chart.part;
      assertChartPart(part);
      const path = allocate('xl/charts', 'chart');
      write(
        path,
        chart.kind === 'xml'
          ? decoder.decode(
              splicePreservedPart(
                encoder.encode(part.xml),
                CHART_CONTENT_TYPE,
                chart.splices ?? [],
              ),
            )
          : part.xml,
      );
      override(path, CHART_CONTENT_TYPE);
      const relatedRels = [];
      for (const related of part.relatedParts ?? []) {
        const target = allocate('xl/drawings', 'chartShapes');
        write(target, related.xml);
        override(target, related.contentType);
        relatedRels.push({
          id: related.relationshipId,
          type: related.relationshipType,
          target: relativePartPath(path, target),
        });
      }
      if (relatedRels.length > 0) write(relsPathFor(path), relsPartXml(relatedRels));
      drawingXml = appendChild(
        drawingXml,
        'wsDr',
        anchorXml(
          chart.anchor,
          ledger.add(CHART_REL, path),
          ++shapeId,
          part.name ?? `Chart ${shapeId}`,
        ),
      );
    }
    write(drawingPath, drawingXml);
    write(drawingRelsPath, relsPartXml([...existingRels, ...ledger.relationships]));
  });
}
