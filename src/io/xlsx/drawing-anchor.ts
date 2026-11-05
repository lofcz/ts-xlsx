import {type AnchorPoint, type ImageAnchor, isOneCellAnchor} from '../../core/image.ts';
import {AuthoringError} from '../../errors.ts';
import {RELATIONSHIPS_NS} from '../opc/namespaces.ts';
import {DRAWINGML_NS, XDR_NS} from './namespaces.ts';

function pointXml(point: AnchorPoint): string {
  if (
    [point.col, point.row, point.colOff ?? 0, point.rowOff ?? 0].some(
      (value) => !Number.isSafeInteger(value) || value < 0,
    ) ||
    point.col > 16383 ||
    point.row > 1048575
  )
    throw new AuthoringError(
      'drawing anchor coordinates must be nonnegative integers within the worksheet',
    );
  return `<xdr:col>${point.col}</xdr:col><xdr:colOff>${point.colOff ?? 0}</xdr:colOff><xdr:row>${point.row}</xdr:row><xdr:rowOff>${point.rowOff ?? 0}</xdr:rowOff>`;
}

export function drawingAnchorXml(anchor: ImageAnchor, body: string): string {
  const ns = ` xmlns:xdr="${XDR_NS}" xmlns:a="${DRAWINGML_NS}" xmlns:r="${RELATIONSHIPS_NS}"`,
    from = `<xdr:from>${pointXml(anchor.from)}</xdr:from>`;
  if (isOneCellAnchor(anchor)) {
    if (![anchor.ext.cx, anchor.ext.cy].every((value) => Number.isSafeInteger(value) && value > 0))
      throw new AuthoringError('drawing extents must be positive integer EMUs');
    return `<xdr:oneCellAnchor${ns}>${from}<xdr:ext cx="${anchor.ext.cx}" cy="${anchor.ext.cy}"/>${body}<xdr:clientData/></xdr:oneCellAnchor>`;
  }
  const editAs = anchor.editAs ?? 'twoCell';
  if (!['twoCell', 'oneCell', 'absolute'].includes(editAs))
    throw new AuthoringError('invalid drawing anchor edit mode');
  return `<xdr:twoCellAnchor${ns} editAs="${editAs}">${from}<xdr:to>${pointXml(anchor.to)}</xdr:to>${body}<xdr:clientData/></xdr:twoCellAnchor>`;
}
