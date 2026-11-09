import assert from 'node:assert/strict';
import {test} from 'node:test';

import {Workbook} from '../../core/workbook.ts';
import {AuthoringError} from '../../errors.ts';
import {xmlChild, xmlChildren, xmlText, xmlTree} from '../../xml/xml-tree.ts';
import {nativeFeatureWorkbook} from './native-features.test-support.ts';
import {assertRelationshipsWired, partMatching} from './package.test-support.ts';
import {readXlsx} from './read.ts';
import {writeXlsx} from './write.ts';

test('large checkbox sheets author and round-trip within a bounded budget', () => {
  const start = performance.now();
  const workbook = new Workbook(),
    sheet = workbook.addWorksheet('Checks');
  for (let row = 1; row <= 50000; row++) sheet.addCheckbox(`A${row}`, row % 2 === 0);
  sheet.addCheckbox('A1', true);
  assert.equal(sheet.checkboxes.length, 50000);
  const bytes = writeXlsx(workbook);
  assert.ok(
    performance.now() - start < 10000,
    '50k checkboxes must not copy the full XML per cell',
  );
  const loaded = readXlsx(bytes).requireWorksheet('Checks');
  assert.equal(loaded.checkboxes.length, 50000);
  loaded.spliceRows(1, 1);
  assert.equal(loaded.checkboxes.length, 49999);
  loaded.removeCheckbox('A49999');
  loaded.addCheckbox('A49999', true);
  assert.equal(loaded.checkboxes.length, 49999);
  assert.equal(loaded.getCell('A49999').value, true);
});

test('large feature collections survive model assignment and grid edits', () => {
  const workbook = new Workbook(),
    sheet = workbook.addWorksheet('Checks');
  const model = sheet.model;
  model.checkboxes = Array.from({length: 150000}, (_, i) => `A${i + 1}`);
  sheet.model = model;
  sheet.spliceRows(1, 1);
  assert.equal(sheet.checkboxes.length, 149999);
  assert.equal(sheet.checkboxes.at(-1), 'A149999');
  sheet.addCheckbox('A149999');
  assert.equal(sheet.checkboxes.length, 149999);
});

test('native features share drawing, media, styles and metadata through repeated round-trips', () => {
  let bytes = writeXlsx(nativeFeatureWorkbook());
  for (let i = 0; i < 3; i++) {
    assertRelationshipsWired(bytes);
    const loaded = readXlsx(bytes),
      sheet = loaded.requireWorksheet('Features');
    assert.equal(sheet.cellImages.length, 1);
    assert.equal(sheet.cellImages[0]?.alt, 'A tiny image');
    assert.deepEqual(sheet.checkboxes, ['C4']);
    assert.equal(sheet.getCell('C4').value, true);
    assert.equal(sheet.getCell('C4').font?.bold, true);
    assert.equal(sheet.sparklineGroups[0]?.sparklines[0]?.formula, 'Features!A2:A3');
    assert.equal(sheet.sparklineGroups[0]?.colors?.series, 'FF123456');
    assert.equal(sheet.shapes.length, 1);
    assert.equal(sheet.conditionalFormattings.length, 1);
    assert.equal(sheet.charts.length, 1);
    assert.deepEqual(sheet.getCell('E2').value, {error: '#SPILL!'});
    const array = sheet.getCell('D2').value;
    assert.ok(array && typeof array === 'object' && 'dynamic' in array && array.dynamic);
    assert.match(partMatching(bytes, /^xl\/drawings\/[^/]+\.xml$/), /Rich /);
    bytes = writeXlsx(loaded);
  }
});

test('removing native shape, checkbox, image and sparkline removes their exported feature', () => {
  const loaded = readXlsx(writeXlsx(nativeFeatureWorkbook())),
    sheet = loaded.requireWorksheet('Features');
  sheet.removeShape(0);
  sheet.removeCheckbox('C4');
  sheet.removeCellImage('C2');
  sheet.removeSparklineGroup(0);
  const reread = readXlsx(writeXlsx(loaded)).requireWorksheet('Features');
  assert.equal(reread.shapes.length, 0);
  assert.equal(reread.cellImages.length, 0);
  assert.equal(reread.sparklineGroups.length, 0);
  assert.equal(reread.checkboxes.length, 0);
  assert.equal(reread.getCell('C4').value, true);
});

test('copied raw shape anchors receive unique identities and follow grid edits', () => {
  const original = new Workbook(),
    sheet = original.addWorksheet('Shapes');
  sheet.addShape(
    {name: 'Moving', text: 'Editable annotation'},
    {
      tl: {col: 2, row: 3},
      br: {col: 5, row: 8},
    },
  );
  sheet.addShape(
    {name: 'Fixed', text: 'Fixed annotation'},
    {
      tl: {col: 2, row: 3},
      br: {col: 5, row: 8},
      editAs: 'absolute',
    },
  );
  const imported = readXlsx(writeXlsx(original)).requireWorksheet('Shapes');
  const copied = new Workbook(),
    destination = copied.addWorksheet('Copies');
  for (const shape of imported.shapes) {
    assert.equal(shape.kind, 'xml');
    if (shape.kind === 'xml') destination.addDrawingXml(shape.xml);
  }
  destination.addShape(
    {name: 'New', text: 'New annotation'},
    {
      tl: {col: 7, row: 3},
      ext: {width: 80, height: 40},
    },
  );
  destination.spliceRows(2, 0, [], []);
  destination.spliceColumns(2, 0, []);
  const bytes = writeXlsx(copied);
  assertRelationshipsWired(bytes);
  const drawing = xmlTree(partMatching(bytes, /^xl\/drawings\/[^/]+\.xml$/));
  const anchors = xmlChildren(drawing);
  const ids = anchors.map((anchor) => xmlChild(anchor, 'sp', 'nvSpPr', 'cNvPr')?.attrs.id);
  assert.equal(new Set(ids).size, 3);
  assert.ok(ids.every(Boolean));
  const position = (name: string, end: string) => {
    const anchor = anchors.find(
      (item) => xmlChild(item, 'sp', 'nvSpPr', 'cNvPr')?.attrs.name === name,
    );
    assert.ok(anchor);
    return ['col', 'row'].map((axis) => Number(xmlText(xmlChild(anchor, end, axis))));
  };
  assert.deepEqual(position('Moving', 'from'), [3, 5]);
  assert.deepEqual(position('Moving', 'to'), [6, 10]);
  assert.deepEqual(position('Fixed', 'from'), [2, 3]);
  assert.deepEqual(position('Fixed', 'to'), [5, 8]);
  assert.deepEqual(position('New', 'from'), [8, 5]);
  const restored = readXlsx(bytes);
  restored.requireWorksheet('Copies').spliceRows(2, 1);
  const rewritten = writeXlsx(restored);
  const after = xmlChildren(xmlTree(partMatching(rewritten, /^xl\/drawings\/[^/]+\.xml$/)));
  assert.deepEqual(
    after.map((anchor) => xmlChild(anchor, 'sp', 'nvSpPr', 'cNvPr')?.attrs.id),
    ids,
  );
  assert.equal(xmlText(xmlChild(after[0], 'from', 'row')), '4');
  assert.equal(xmlText(xmlChild(after[1], 'from', 'row')), '3');
});

test('raw shapes reject incomplete anchors instead of writing broken drawings', () => {
  for (const xml of ['<xdr:sp/>', '<xdr:oneCellAnchor><xdr:pic/></xdr:oneCellAnchor>']) {
    const workbook = new Workbook();
    workbook.addWorksheet('Invalid').addDrawingXml(xml);
    assert.throws(() => writeXlsx(workbook), AuthoringError);
  }
});
