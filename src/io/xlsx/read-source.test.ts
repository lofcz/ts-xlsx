import assert from 'node:assert/strict';
import {test} from 'node:test';

import {Workbook} from '../../core/workbook.ts';
import {foreignPackage, foreignSheet, SHEET1} from './package.test-support.ts';
import {
  indexWorksheetSource,
  indexSharedStrings,
  readXlsxSource,
  worksheetSourceRow,
  worksheetSourceCell,
} from './read-source.ts';
import {writeXlsx} from './write.ts';

test('shared string offsets preserve rich text, empty items and literal markup without decoding', () => {
  const items = [
    '<s:si><s:r><s:t>A &amp; B</s:t></s:r></s:si>',
    '<s:si/>',
    '<s:si><s:t><![CDATA[<si/>]]></s:t></s:si>',
  ];
  const xml = `<s:sst xmlns:s="urn:test"><!-- <si/> -->${items.join('')}</s:sst>`;
  const offsets = indexSharedStrings(xml);
  assert.equal(offsets.length, items.length * 2);
  assert.deepEqual(
    items.map((_, i) => xml.slice(offsets[i * 2], offsets[i * 2 + 1])),
    items,
  );
  assert.equal(indexSharedStrings('<sst/>').length, 0);
  assert.throws(() => indexSharedStrings('<sst><si><t>unfinished'));
});

test('opens a workbook as source without constructing cell values', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('Data');
  sheet.getCell('A1').value = 'hello';
  sheet.getCell('J50000').value = 42;
  const source = readXlsxSource(writeXlsx(wb));
  const indexed = source.worksheets[0];
  assert.ok(indexed);
  assert.equal(indexed.name, 'Data');
  assert.equal(indexed.rowCount, 50000);
  assert.equal(indexed.columnCount, 10);
  assert.equal(indexed.cellCount, 2);
  assert.equal(indexed.rows.length, 6);
  assert.match(worksheetSourceRow(indexed, 50000) ?? '', /42/);
  assert.equal(worksheetSourceRow(indexed, 2), undefined);
  assert.match(worksheetSourceCell(indexed, 50000, 10) ?? '', /<v>42<\/v>/);
  assert.equal(worksheetSourceCell(indexed, 50000, 9), undefined);
  assert.doesNotMatch(indexed.metadataXml, /<c[\s>]/);
});

test('metadata retains row properties, formulas and all worksheet features', () => {
  const xml = foreignSheet(
    '<row r="4" ht="25" customHeight="1"><c r="A4"><v>7</v></c>' +
      '<c r="B4"><f>A4*2</f><v>14</v></c></row>',
  );
  const indexed = indexWorksheetSource(xml);
  assert.match(indexed.metadataXml, /ht="25"/);
  assert.match(indexed.metadataXml, /<f>A4\*2<\/f>/);
  assert.doesNotMatch(indexed.metadataXml, /<v>7<\/v>/);
  assert.equal(indexed.rowCount, 4);
  assert.equal(indexed.columnCount, 2);
  assert.match(worksheetSourceRow(indexed, 4) ?? '', /<v>7<\/v>/);
});

test('indexes namespaced, omitted and empty rows without matching markup inside comments or CDATA', () => {
  const xml =
    '<s:worksheet xmlns:s="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
    '<s:sheetData><!-- <row r="500"/> --><s:row><s:c t="inlineStr"><s:is><s:t>' +
    '<![CDATA[<row r="999"/><c><f>not a formula</f></c>]]></s:t></s:is></s:c></s:row>' +
    '<s:row r="4"/><s:row><s:c/></s:row></s:sheetData></s:worksheet>';
  const indexed = indexWorksheetSource(xml);
  assert.deepEqual([indexed.rowCount, indexed.columnCount, indexed.cellCount], [5, 1, 2]);
  assert.equal(indexed.rows.length, 9);
  assert.equal(worksheetSourceRow(indexed, 4), '<s:row r="4"/>');
  assert.equal(worksheetSourceCell(indexed, 4, 1), undefined);
  assert.equal(worksheetSourceCell(indexed, 5, 1), '<s:c/>');
  assert.doesNotMatch(indexed.metadataXml, /not a formula/);
});

test('source reads enforce package expansion limits and reject truncated rows', () => {
  const data = foreignPackage({[SHEET1]: foreignSheet('<row><c><v>1</v></c></row>')});
  assert.throws(() => readXlsxSource(data, {maxUncompressedBytes: 1}));
  assert.throws(() => indexWorksheetSource('<worksheet><sheetData><row><c/>'));
});

test('retained metadata cells keep implicit row and column addresses after ordinary cells are removed', () => {
  const xml = foreignSheet(
    '<row><c><v>1</v></c></row><row><c><v>2</v></c>' + '<c><f>A2*2</f><v>4</v></c></row>',
  );
  const indexed = indexWorksheetSource(xml);
  assert.match(indexed.metadataXml, /<row r="2">/);
  assert.match(indexed.metadataXml, /<c r="B2"><f>A2\*2/);
});

test('processing instructions inside cells cannot truncate a source span', () => {
  const cell = '<c r="A1"><v>1<?payload end="</c>"?></v></c>';
  const xml = foreignSheet(`<row r="1">${cell}</row>`).replace(/^<\?xml[^?]*\?>/, '');
  const indexed = indexWorksheetSource(xml);
  assert.equal(worksheetSourceCell(indexed, 1, 1), cell);
});

test('duplicate cell declarations resolve to the last source value', () => {
  const indexed = indexWorksheetSource(
    foreignSheet('<row r="1"><c r="A1"><v>1</v></c><c r="A1"><v>2</v></c></row>'),
  );
  assert.equal(worksheetSourceCell(indexed, 1, 1), '<c r="A1"><v>2</v></c>');
});

test('compact cell indexing agrees with the general scanner across cell encodings', () => {
  const cells = [
    '<c r="A3" s="2" t="n"><v>17</v></c>',
    '<c r="B3" t="inlineStr"><is><t>A &amp; B</t></is></c>',
    '<c r="C3" t="s" s="1"><v>2</v></c>',
    '<c r="D3" />',
    '<c r="E3"><f>C3*2</f><v>4</v></c>',
    '<c r="F3"><x:f xmlns:x="urn:test">C3*3</x:f><v>6</v></c>',
    '<c r="G3" cm="1"/>',
    '<c r="H3" vm="2"><v>8</v></c>',
    "<c s='3' r='I3'><v>9</v></c>",
    '<c r="&#74;3"><v>10</v></c>',
    '<c r="k3"><v>11</v></c>',
    '<c><f>K3+1</f><v>12</v></c>',
    '<c r="M3" custom="a > b"><v>13</v></c>',
    '<c r="N3" t="str"><v>&lt;f&gt;literal&lt;/f&gt;</v></c>',
    '<c r="N3"><v>last</v></c>',
  ];
  const body = `<row r="3" ht="24">${cells.join('')}</row><row r="7"><c r="XFD7"/></row>`;
  const plain = indexWorksheetSource(foreignSheet(body));
  const general = indexWorksheetSource(
    foreignSheet(body).replace('<sheetData>', '<sheetData><!-- general -->'),
  );
  assert.deepEqual(
    [plain.rowCount, plain.columnCount, plain.cellCount],
    [7, 16384, cells.length + 1],
  );
  assert.deepEqual(
    [plain.rowCount, plain.columnCount, plain.cellCount],
    [general.rowCount, general.columnCount, general.cellCount],
  );
  assert.equal(plain.metadataXml, general.metadataXml);
  for (let c = 1; c <= 14; c++)
    assert.equal(worksheetSourceCell(plain, 3, c), worksheetSourceCell(general, 3, c));
  assert.equal(worksheetSourceCell(plain, 7, 16384), '<c r="XFD7"/>');
  assert.match(plain.metadataXml, /<x:f/);
  assert.match(plain.metadataXml, /r="L3"/);
  assert.doesNotMatch(plain.metadataXml, /literal/);
});

test('compact cell indexing rejects invalid columns and truncated cell bodies', () => {
  assert.throws(() => indexWorksheetSource(foreignSheet('<row><c r="XFE1"/></row>')), /column/);
  assert.throws(
    () => indexWorksheetSource(foreignSheet('<row><c r="A1"><v>1</v></row>')),
    /Unclosed/,
  );
});
