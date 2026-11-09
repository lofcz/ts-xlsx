import {MAX_COLUMN, MAX_ROW, numberToColumn} from '../../core/address.ts';
import {XmlParseError} from '../../xml/errors.ts';
import {elementRanges, parseXml} from '../../xml/xml-read.ts';
import {localName, markupAt, parseAttributes, tagAt} from '../../xml/xml-scan.ts';
import {openSpreadsheetPackage, readPartRelationships} from '../opc/read-opc.ts';
import type {ReadPackageOptions} from '../opc/read-options.ts';
import {unsupportedWorkbookPart} from '../opc/sniff-format.ts';
import {workbookSheetsPass} from './read-workbook-xml.ts';

/** An immutable worksheet source suitable for structured cloning to another thread.
 * Values and styles remain encoded until the consumer requests a row.
 */
export interface WorksheetSource {
  readonly xml: string;
  /** Worksheet XML with ordinary cells omitted. Formulas and row properties remain. */
  readonly metadataXml: string;
  /** Triples of 1-based row number, start offset, end offset into {@link xml}. */
  readonly rows: Uint32Array;
  /** Triples of 1-based column, start offset and end offset for each cell. */
  readonly cells: Uint32Array;
  /** Cell indexes at the beginning of each indexed row, with one final end index. */
  readonly rowCellStarts: Uint32Array;
  readonly rowCount: number;
  readonly columnCount: number;
  readonly cellCount: number;
}

/** A package opened without building a mutable workbook or decoding cell values. */
export interface XlsxSource {
  readonly files: Readonly<Record<string, Uint8Array>>;
  readonly worksheets: readonly (WorksheetSource & {
    readonly name: string;
    readonly path: string;
  })[];
}

/** Pairs of start/end offsets for shared string items in document order.
 * Consumers can slice one complete `si` element without decoding the rest of the pool.
 */
export function indexSharedStrings(xml: string): Uint32Array {
  const offsets: number[] = [];
  for (const item of elementRanges(xml, 'si')) offsets.push(item.start, item.end);
  return Uint32Array.from(offsets);
}

/** Open the package once and index worksheet rows for on-demand cell decoding. */
export function readXlsxSource(data: Uint8Array, options: ReadPackageOptions = {}): XlsxSource {
  const opened = openSpreadsheetPackage(data, options.maxUncompressedBytes);
  if (opened.workbookXml === undefined)
    throw unsupportedWorkbookPart(
      opened.pkg.partBytes,
      opened.documentPath,
      'the binary .xlsb format cannot be indexed as XML; read it with readXlsx or readXlsb',
    );
  const pass = workbookSheetsPass();
  parseXml(opened.workbookXml, pass.handlers);
  const rels = readPartRelationships(opened.documentPath, opened.pkg.partText);
  const worksheets: XlsxSource['worksheets'][number][] = [];
  for (const sheet of pass.result()) {
    const rel = rels.byId(sheet.relId);
    if (rel === undefined || rel.external || !rel.type.endsWith('/worksheet')) continue;
    const path = opened.pkg.partKey(rels.pathOf(rel.target));
    const xml = path === undefined ? undefined : opened.pkg.partText(path);
    if (path !== undefined && xml !== undefined) {
      worksheets.push({...indexWorksheetSource(xml), name: sheet.name, path});
    }
  }
  return {files: opened.files, worksheets};
}

/** Index source locations without parsing strings, numbers or style records. */
export function indexWorksheetSource(xml: string): WorksheetSource {
  const rows: number[] = [];
  const cells: number[] = [];
  const rowCellStarts: number[] = [];
  const metadata: string[] = [];
  let dataStart = -1;
  let dataEnd = -1;
  let rowStart = -1;
  let rowName = '';
  let rowOpen = '';
  let keepRow = false;
  let row = 0;
  let column = 0;
  let rowCount = 0;
  let columnCount = 0;
  let cellCount = 0;
  let cellStart = -1;
  let keepCell = false;
  let cellOpenEnd = 0;
  let cellOpen = '';
  let rowCells: string[] = [];
  const completeRow = (end: number): void => {
    rows.push(row, rowStart, end);
    if (keepRow || rowCells.length > 0) {
      metadata.push(rowOpen, ...rowCells, `</${rowName}>`);
    }
    rowStart = -1;
    rowCells = [];
  };
  // Canonical cell bodies cannot contain literal '<' in text. Markup or
  // prefixed tags use the full scanner below; ordinary values can be skipped
  // in one jump without allocating events for every <v>/<is>/<t> descendant.
  const declarationEnd = /^\s*<\?xml[\t\n\r ]/.test(xml) ? xml.indexOf('?>') + 2 : 0;
  const simpleCells =
    !xml.includes('<!--') &&
    !xml.includes('<![CDATA[') &&
    !/<\/?[\w.-]+:/.test(xml) &&
    xml.indexOf('<?', declarationEnd) < 0;
  let formulaAt = simpleCells ? xml.indexOf('<f') : -1;
  const completeCell = (end: number): void => {
    if (cellStart >= 0) cells.push(column, cellStart, end);
    if (keepCell && cellStart >= 0) rowCells.push(cellOpen + xml.slice(cellOpenEnd, end));
    cellStart = -1;
  };
  for (let offset = 0; offset < xml.length;) {
    const lt = xml.indexOf('<', offset);
    if (lt < 0) break;
    const markup = markupAt(xml, lt);
    if (markup !== undefined) {
      offset = markup.next;
      continue;
    }
    const tag = tagAt(xml, lt);
    offset = tag.next;
    const local = localName(tag.name);
    if (local === 'sheetData') {
      if (tag.close) {
        dataEnd = lt;
        break;
      }
      if (tag.selfClosing) break;
      dataStart = offset;
      continue;
    }
    if (dataStart < 0) continue;
    if (local === 'row') {
      if (tag.close) {
        if (rowStart < 0) throw new XmlParseError('Worksheet row closes without an opening tag');
        completeRow(offset);
      } else {
        if (rowStart >= 0) throw new XmlParseError('Worksheet rows must not nest');
        const attrs = parseAttributes(tag.attrSource);
        const next = attrs.r === undefined ? row + 1 : Number(attrs.r);
        if (!Number.isInteger(next) || next < 1 || next > MAX_ROW || next <= row) {
          throw new XmlParseError('Worksheet source rows must have increasing valid row numbers');
        }
        row = next;
        rowCellStarts.push(cells.length / 3);
        rowCount = Math.max(rowCount, row);
        column = 0;
        rowStart = lt;
        rowName = tag.name;
        rowOpen =
          attrs.r === undefined
            ? `<${tag.name} r="${row}"${tag.attrSource}${tag.selfClosing ? '/>' : '>'}`
            : xml.slice(lt, offset);
        keepRow = Object.keys(attrs).some((key) => key !== 'r' && key !== 'spans');
        if (tag.selfClosing) {
          rows.push(row, lt, offset);
          if (keepRow) metadata.push(rowOpen);
          rowStart = -1;
        }
      }
    } else if (rowStart >= 0 && local === 'c') {
      if (tag.close) {
        completeCell(offset);
      } else {
        const attrs = parseAttributes(tag.attrSource);
        const ref = attrs.r;
        if (ref === undefined) column += 1;
        else {
          column = 0;
          const upper = ref.toUpperCase();
          for (let i = 0; i < upper.length; i++) {
            const ch = upper.charCodeAt(i);
            if (ch < 65 || ch > 90) break;
            column = column * 26 + ch - 64;
          }
        }
        if (column < 1 || column > MAX_COLUMN)
          throw new XmlParseError('Invalid worksheet source column');
        columnCount = Math.max(columnCount, column);
        cellCount++;
        cellStart = tag.selfClosing ? -1 : lt;
        if (tag.selfClosing) cells.push(column, lt, offset);
        cellOpenEnd = offset;
        cellOpen =
          ref === undefined
            ? `<${tag.name} r="${numberToColumn(column)}${row}"${tag.attrSource}${tag.selfClosing ? '/>' : '>'}`
            : xml.slice(lt, offset);
        // Cell metadata links images and rich values; consumers need those at open time.
        keepCell = attrs.vm !== undefined || attrs.cm !== undefined;
        if (tag.selfClosing && keepCell) rowCells.push(cellOpen);
        if (!tag.selfClosing && simpleCells) {
          const close = `</${tag.name}>`;
          const end = xml.indexOf(close, offset);
          if (end < 0) throw new XmlParseError('Unclosed worksheet source cell');
          while (formulaAt >= 0 && formulaAt < offset) formulaAt = xml.indexOf('<f', formulaAt + 2);
          if (formulaAt >= 0 && formulaAt < end) keepCell = true;
          offset = end + close.length;
          completeCell(offset);
        }
      }
    } else if (cellStart >= 0 && local === 'f' && !tag.close) keepCell = true;
  }
  if (rowStart >= 0 || (dataStart >= 0 && dataEnd < 0)) {
    throw new XmlParseError('Unclosed worksheet source sheetData');
  }
  return {
    xml,
    metadataXml:
      dataStart < 0 ? xml : xml.slice(0, dataStart) + metadata.join('') + xml.slice(dataEnd),
    rows: Uint32Array.from(rows),
    cells: Uint32Array.from(cells),
    rowCellStarts: Uint32Array.from([...rowCellStarts, cells.length / 3]),
    rowCount,
    columnCount,
    cellCount,
  };
}

/** Read one raw row, using its 1-based worksheet number; absent rows return undefined. */
export function worksheetSourceRow(source: WorksheetSource, row: number): string | undefined {
  const index = sourceRowIndex(source, row);
  return index < 0
    ? undefined
    : source.xml.slice(source.rows[index * 3 + 1], source.rows[index * 3 + 2]);
}

function sourceRowIndex(source: WorksheetSource, row: number): number {
  let low = 0;
  let high = source.rows.length / 3 - 1;
  while (low <= high) {
    const mid = (low + high) >>> 1;
    const number = source.rows[mid * 3];
    if (number === row) return mid;
    if (number !== undefined && number < row) low = mid + 1;
    else high = mid - 1;
  }
  return -1;
}

/** Read a raw cell by 1-based row and column without decoding its neighbours. */
export function worksheetSourceCell(
  source: WorksheetSource,
  row: number,
  column: number,
): string | undefined {
  const index = sourceRowIndex(source, row);
  if (index < 0) return undefined;
  const start = source.rowCellStarts[index] ?? 0;
  const end = source.rowCellStarts[index + 1] ?? start;
  for (let i = end - 1; i >= start; i--) {
    if (source.cells[i * 3] === column)
      return source.xml.slice(source.cells[i * 3 + 1], source.cells[i * 3 + 2]);
  }
  return undefined;
}
