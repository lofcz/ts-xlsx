import {shiftIndex} from '../../core/grid-shift.ts';
import type {WorksheetShape} from '../../core/sheet-features.ts';
import {AuthoringError} from '../../errors.ts';
import {localName} from '../../xml/xml-scan.ts';
import {
  xmlChild,
  xmlChildren,
  xmlSerialize,
  xmlText,
  xmlTree,
  type XmlNode,
} from '../../xml/xml-tree.ts';
import {escapeAttr, escapeText} from '../../xml/xml.ts';
import {RELATIONSHIPS_NS} from '../opc/namespaces.ts';
import {drawingAnchorXml} from './drawing-anchor.ts';
import {DRAWINGML_NS, XDR_NS} from './namespaces.ts';

function color(value: string): string {
  if (!/^#?[\da-f]{6}$/i.test(value))
    throw new AuthoringError('shape colors must be six hexadecimal RGB digits');
  return `<a:solidFill><a:srgbClr val="${value.replace(/^#/, '').toUpperCase()}"/></a:solidFill>`;
}

export function shapeAnchors(
  shape: WorksheetShape,
  nextId: () => number,
  preserveIds = false,
): XmlNode[] {
  if (shape.kind === 'xml') {
    const root = xmlTree(
      `<xdr:wsDr xmlns:xdr="${XDR_NS}" xmlns:a="${DRAWINGML_NS}" xmlns:r="${RELATIONSHIPS_NS}">${shape.xml}</xdr:wsDr>`,
    );
    const visit = (node: XmlNode) => {
      if (localName(node.name) === 'cNvPr') node.attrs.id = String(nextId());
      for (const child of xmlChildren(node)) visit(child);
    };
    for (const anchor of xmlChildren(root)) {
      if (
        !['oneCellAnchor', 'twoCellAnchor', 'absoluteAnchor'].includes(localName(anchor.name)) ||
        !['sp', 'cxnSp', 'grpSp'].some((name) => xmlChild(anchor, name))
      )
        throw new AuthoringError('shape XML must contain complete DrawingML shape anchors');
      if (!preserveIds) visit(anchor);
      if (localName(anchor.name) !== 'absoluteAnchor' && anchor.attrs.editAs !== 'absolute')
        for (const edit of shape.splices ?? [])
          for (const end of ['from', 'to']) {
            const coordinate = xmlChild(anchor, end, edit.splice.axis);
            if (coordinate)
              coordinate.children = [
                String(shiftIndex(Number(xmlText(coordinate)) + 1, edit.splice) - 1),
              ];
          }
      Object.assign(anchor.attrs, {
        'xmlns:xdr': XDR_NS,
        'xmlns:a': DRAWINGML_NS,
        'xmlns:r': RELATIONSHIPS_NS,
      });
    }
    return xmlChildren(root);
  }
  const options = shape.options,
    id = nextId(),
    geometry = options.geometry ?? 'rect';
  if (!/^[a-zA-Z][a-zA-Z0-9]*$/.test(geometry))
    throw new AuthoringError('invalid shape preset geometry');
  for (const value of [options.rotation, options.outline?.width])
    if (value !== undefined && !Number.isFinite(value))
      throw new AuthoringError('shape dimensions must be finite');
  const fill = options.fill ? color(options.fill) : '<a:noFill/>';
  const line = options.outline
    ? `<a:ln w="${Math.round((options.outline.width ?? 1) * 9525)}">${color(options.outline.color)}</a:ln>`
    : '';
  const runs = typeof options.text === 'string' ? [{text: options.text}] : (options.text ?? []);
  const text = runs
    .map((run) => {
      const font = run.font;
      if (font?.size !== undefined && (!Number.isFinite(font.size) || font.size <= 0))
        throw new AuthoringError('shape text size must be positive');
      return `<a:r><a:rPr${font?.size === undefined ? '' : ` sz="${Math.round(font.size * 100)}"`}${font?.bold ? ' b="1"' : ''}${font?.italic ? ' i="1"' : ''}>${font?.color ? color(font.color) : ''}${font?.name ? `<a:latin typeface="${escapeAttr(font.name)}"/>` : ''}</a:rPr><a:t>${escapeText(run.text)}</a:t></a:r>`;
    })
    .join('');
  const body = `<xdr:sp><xdr:nvSpPr><xdr:cNvPr id="${id}" name="${escapeAttr(options.name ?? `Shape ${id}`)}"${options.alt ? ` descr="${escapeAttr(options.alt)}"` : ''}/><xdr:cNvSpPr${options.textBox ? ' txBox="1"' : ''}/></xdr:nvSpPr><xdr:spPr>${options.rotation === undefined ? '' : `<a:xfrm rot="${Math.round(options.rotation * 60000)}"/>`}<a:prstGeom prst="${geometry}"><a:avLst/></a:prstGeom>${fill}${line}</xdr:spPr><xdr:txBody><a:bodyPr/><a:lstStyle/><a:p>${text}</a:p></xdr:txBody></xdr:sp>`;
  return [xmlTree(drawingAnchorXml(shape.anchor, body))];
}

/** Opaque imported shape anchors stay attached to their original drawing object identity. */
export function importedShapes(xml: string): WorksheetShape[] {
  const root = xmlTree(xml);
  return xmlChildren(root)
    .filter((anchor) => ['sp', 'cxnSp', 'grpSp'].some((name) => xmlChild(anchor, name)))
    .map((anchor) => {
      const body = xmlChild(anchor, 'sp') ?? xmlChild(anchor, 'cxnSp') ?? xmlChild(anchor, 'grpSp');
      const nv = body ? xmlChildren(body)[0] : undefined,
        id = xmlChild(nv, 'cNvPr')?.attrs.id;
      const owned = structuredClone(anchor);
      Object.assign(owned.attrs, root.attrs);
      return {kind: 'xml', xml: xmlSerialize(owned), ...(id ? {sourceId: id} : {})};
    });
}
