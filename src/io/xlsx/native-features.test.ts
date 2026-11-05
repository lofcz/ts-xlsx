import assert from 'node:assert/strict';
import {test} from 'node:test';

import {nativeFeatureWorkbook} from './native-features.test-support.ts';
import {assertRelationshipsWired, partMatching} from './package.test-support.ts';
import {readXlsx} from './read.ts';
import {writeXlsx} from './write.ts';

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
