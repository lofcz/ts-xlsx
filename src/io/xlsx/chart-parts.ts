import {resolveChartCaches, type ChartCacheDiagnostic} from '../../core/chart-cache.ts';
import {isExtendedChart, type ChartPart} from '../../core/chart.ts';
import type {ImageAnchor} from '../../core/image.ts';
import {INTERNAL} from '../../core/internal.ts';
import {modelEqual} from '../../core/model-patch.ts';
import type {Workbook} from '../../core/workbook.ts';
import type {Worksheet} from '../../core/worksheet.ts';
import {AuthoringError} from '../../errors.ts';
import {mergeXml} from '../../xml/xml-merge.ts';
import {directChildRange, elementRange, openElements, parseXml} from '../../xml/xml-read.ts';
import {xmlChild, xmlChildren, xmlSerialize, xmlTree} from '../../xml/xml-tree.ts';
import {escapeAttr, XML_DECLARATION} from '../../xml/xml.ts';
import {RELATIONSHIPS_NS} from '../opc/namespaces.ts';
import {relativePartPath, relsPathFor, resolveRelativePart} from '../opc/part-paths.ts';
import {parseRelationshipRecords} from '../opc/relationship-records.ts';
import {relsPartXml} from '../opc/rels.ts';
import {CHARTEX_NS, CHARTEX_REL, CHARTEX_CONTENT_TYPE} from './chart-ex-xml.ts';
import {CHART_NS, chartXml} from './chart-xml.ts';
import {drawingAnchorXml} from './drawing-anchor.ts';
import {DRAWINGML_NS, XDR_NS} from './namespaces.ts';
import {RelationshipLedger} from './package-plan.ts';
import {splicePreservedPart} from './preserved-splices.ts';
import {shapeAnchors} from './shape-xml.ts';

const CHART_REL = `${RELATIONSHIPS_NS}/chart`;
const DRAWING_REL = `${RELATIONSHIPS_NS}/drawing`;
const CHART_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml';
const DRAWING_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.drawing+xml';
function chartFrame(node: import('../../xml/xml-tree.ts').XmlNode) {
  return (
    xmlChild(node, 'graphicFrame') ?? xmlChild(node, 'AlternateContent', 'Choice', 'graphicFrame')
  );
}
function chartId(node: import('../../xml/xml-tree.ts').XmlNode): string | undefined {
  return xmlChild(chartFrame(node), 'graphic', 'graphicData', 'chart')?.attrs['r:id'];
}
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

function anchorXml(
  anchor: ImageAnchor,
  relationshipId: string,
  id: number,
  name: string,
  extended = false,
): string {
  const graphic = `<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${id}" name="${escapeAttr(name)}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="${extended ? CHARTEX_NS : CHART_NS}">${extended ? `<cx:chart xmlns:cx="${CHARTEX_NS}" r:id="${escapeAttr(relationshipId)}"/>` : `<c:chart xmlns:c="${CHART_NS}" r:id="${escapeAttr(relationshipId)}"/>`}</a:graphicData></a:graphic></xdr:graphicFrame>`;
  const fallback = `<xdr:sp><xdr:nvSpPr><xdr:cNvPr id="${id}" name="${escapeAttr(name)}"/><xdr:cNvSpPr txBox="1"/></xdr:nvSpPr><xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr><xdr:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>This chart requires Excel 2016 or later.</a:t></a:r></a:p></xdr:txBody></xdr:sp>`;
  const frame = extended
    ? `<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice xmlns:cx2="http://schemas.microsoft.com/office/drawing/2015/10/21/chartex" Requires="cx2">${graphic}</mc:Choice><mc:Fallback>${fallback}</mc:Fallback></mc:AlternateContent>`
    : graphic;
  return drawingAnchorXml(anchor, frame);
}

function assertChartPart(part: ChartPart): void {
  parseXml(part.xml, {onOpen() {}});
  const root = elementRange(part.xml, ['chartSpace']);
  if (root === undefined) throw new AuthoringError('chart XML must have a chartSpace root');
  const prefix = root.name.includes(':') ? `xmlns:${root.name.split(':')[0]}` : 'xmlns';
  if (root.attrs[prefix] !== CHART_NS && root.attrs[prefix] !== CHARTEX_NS)
    throw new AuthoringError('chart XML must use the DrawingML chart or chartEx namespace');
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
  workbook: Workbook,
  diagnostics: ChartCacheDiagnostic[] = [],
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
    const removed = new Set(sheet[INTERNAL].removedChartIds);
    if (
      charts.length === 0 &&
      removed.size === 0 &&
      sheet.shapes.length === 0 &&
      sheet[INTERNAL].removedShapeIds.length === 0
    )
      return;
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
        .map((name) => directChildRange(xml, 'worksheet', name))
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
    const existingRels = parseRelationshipRecords(read(drawingRelsPath)).filter(
      (rel) => !removed.has(rel.id),
    );
    const ledger = new RelationshipLedger(drawingPath, new Set(existingRels.map((rel) => rel.id)));
    let drawingXml = read(drawingPath);
    if (removed.size > 0) {
      const tree = xmlTree(drawingXml);
      tree.children = tree.children.filter(
        (child) => typeof child === 'string' || !removed.has(chartId(child) ?? ''),
      );
      drawingXml = XML_DECLARATION + xmlSerialize(tree);
    }
    let shapeId = 0;
    for (const tag of openElements(drawingXml, 'cNvPr'))
      shapeId = Math.max(shapeId, Number(tag.attrs.id) || 0);
    for (const chart of charts) {
      const options =
        chart.kind === 'chart'
          ? resolveChartCaches(workbook, sheet.name, chart.options, diagnostics)
          : undefined;
      if (chart.kind === 'chart' && chart.source?.relationshipId && options) {
        const rel = existingRels.find(
          (item) =>
            item.id === chart.source?.relationshipId &&
            (item.type === CHART_REL || item.type === CHARTEX_REL) &&
            !item.external,
        );
        if (!rel) throw new AuthoringError('imported chart drawing relationship is missing');
        const path = resolveRelativePart(drawingPath, rel.target);
        const source = xmlTree(chart.source.editedXml ?? read(path));
        const extended = isExtendedChart(options.type);
        const converted = (rel.type === CHARTEX_REL) !== extended;
        const updated = converted
          ? xmlTree(chartXml(options))
          : mergeXml(source, xmlTree(chartXml(chart.source.options)), xmlTree(chartXml(options)));
        if (converted) {
          existingRels[existingRels.indexOf(rel)] = {
            ...rel,
            type: extended ? CHARTEX_REL : CHART_REL,
          };
          const types = xmlTree(read('[Content_Types].xml'));
          const entry = xmlChildren(types, 'Override').find(
            (item) => item.attrs.PartName === `/${path}`,
          );
          if (entry) entry.attrs.ContentType = extended ? CHARTEX_CONTENT_TYPE : CHART_CONTENT_TYPE;
          write('[Content_Types].xml', XML_DECLARATION + xmlSerialize(types));
        }
        updated.attrs['xmlns:c'] = CHART_NS;
        updated.attrs['xmlns:a'] = DRAWINGML_NS;
        write(path, XML_DECLARATION + xmlSerialize(updated));
        if (converted || !modelEqual(chart.anchor, chart.source.anchor)) {
          const tree = xmlTree(drawingXml);
          const previous = xmlChildren(tree).find((node) => chartId(node) === rel.id);
          if (previous) {
            const replacement = xmlTree(
              anchorXml(
                chart.anchor,
                rel.id,
                Number(xmlChild(chartFrame(previous), 'nvGraphicFramePr', 'cNvPr')?.attrs.id) ||
                  ++shapeId,
                options.title ?? 'Chart',
                extended,
              ),
            );
            if (!converted) {
              const frame = xmlChild(previous, 'AlternateContent') ?? chartFrame(previous);
              const newFrame = xmlChild(replacement, 'AlternateContent') ?? chartFrame(replacement);
              if (frame && newFrame)
                replacement.children[replacement.children.indexOf(newFrame)] = frame;
            }
            tree.children[tree.children.indexOf(previous)] = replacement;
            drawingXml = XML_DECLARATION + xmlSerialize(tree);
          }
        }
        continue;
      }
      const part: ChartPart =
        chart.kind === 'chart'
          ? {
              xml:
                chart.source &&
                isExtendedChart(chart.source.options.type) === isExtendedChart(chart.options.type)
                  ? XML_DECLARATION +
                    xmlSerialize(
                      mergeXml(
                        xmlTree(chart.source.xml),
                        xmlTree(chartXml(chart.source.options)),
                        xmlTree(chartXml(options ?? chart.options)),
                      ),
                    )
                  : chartXml(options ?? chart.options),
              ...(chart.options.title === undefined ? {} : {name: chart.options.title}),
            }
          : chart.part;
      assertChartPart(part);
      const extended = Object.values(xmlTree(part.xml).attrs).includes(CHARTEX_NS);
      const contentType = extended ? CHARTEX_CONTENT_TYPE : CHART_CONTENT_TYPE;
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
      override(path, contentType);
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
          ledger.add(extended ? CHARTEX_REL : CHART_REL, path),
          ++shapeId,
          part.name ?? `Chart ${shapeId}`,
          extended,
        ),
      );
    }
    const drawingTree = xmlTree(drawingXml);
    const removedShapes = new Set(sheet[INTERNAL].removedShapeIds);
    const sourceId = (anchor: import('../../xml/xml-tree.ts').XmlNode): string | undefined => {
      const body = xmlChild(anchor, 'sp') ?? xmlChild(anchor, 'cxnSp') ?? xmlChild(anchor, 'grpSp');
      return xmlChild(body ? xmlChildren(body)[0] : undefined, 'cNvPr')?.attrs.id;
    };
    drawingTree.children = drawingTree.children.filter(
      (child) => typeof child === 'string' || !removedShapes.has(sourceId(child) ?? ''),
    );
    for (const shape of sheet.shapes) {
      if (shape.kind === 'xml' && shape.sourceId) {
        const existing = xmlChildren(drawingTree).find(
          (anchor) => sourceId(anchor) === shape.sourceId,
        );
        if (existing) {
          if (shape.splices?.length || shape.modified) {
            const replacement = shapeAnchors(shape, () => ++shapeId, true)[0];
            if (replacement)
              drawingTree.children[drawingTree.children.indexOf(existing)] = replacement;
          }
          continue;
        }
      }
      drawingTree.children.push(...shapeAnchors(shape, () => ++shapeId));
    }
    write(drawingPath, XML_DECLARATION + xmlSerialize(drawingTree));
    write(drawingRelsPath, relsPartXml([...existingRels, ...ledger.relationships]));
  });
}
