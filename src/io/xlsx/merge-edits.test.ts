import assert from 'node:assert/strict';
import {test} from 'node:test';

import {strToU8, unzipSync, zipSync} from 'fflate';

import {INTERNAL} from '../../core/internal.ts';
import {Workbook} from '../../core/workbook.ts';
import {xmlTree, xmlChild, xmlChildren, xmlText} from '../../xml/xml-tree.ts';
import {mergeXlsxEdits} from './merge-edits.ts';
import {nativeFeatureWorkbook} from './native-features.test-support.ts';
import {patchParts, partMatching, assertRelationshipsWired} from './package.test-support.ts';
import {readXlsx} from './read.ts';
import {writeXlsx} from './write.ts';

test('editing an expanded shared formula replaces its source group membership', () => {
  const editor = new Workbook();
  const sheet = editor.addWorksheet('Data');
  sheet.getCell('B1').value = {formula: 'A1*2', result: 2};
  sheet.getCell('B2').value = {formula: 'A2*2', result: 4};
  const baseline = writeXlsx(editor);
  const original = patchParts(baseline, {
    'xl/worksheets/sheet1.xml': (xml) =>
      xml
        .replace('<f>A1*2</f>', '<f t="shared" ref="B1:B2" si="0">A1*2</f>')
        .replace('<f>A2*2</f>', '<f t="shared" si="0"/>'),
  });
  sheet.getCell('B2').value = {formula: 'A2*3', result: 6};
  const result = mergeXlsxEdits(original, baseline, writeXlsx(editor), {
    sheets: [{name: 'Data', originalName: 'Data', baselineName: 'Data'}],
  });
  const value = readXlsx(result.data).requireWorksheet('Data').getCell('B2').value;
  assert.deepEqual(value, {formula: 'A2*3', result: 6});
  sheet.getCell('B1').value = 5;
  sheet.getCell('B2').value = {formula: 'A2*2', result: 4};
  const removedMaster = mergeXlsxEdits(original, baseline, writeXlsx(editor), {
    sheets: [{name: 'Data', originalName: 'Data', baselineName: 'Data'}],
  });
  assert.deepEqual(readXlsx(removedMaster.data).requireWorksheet('Data').getCell('B2').value, {
    formula: 'A2*2',
    result: 4,
  });
});

test('rewriting a workbook discards the source calculation chain and requests recalculation', () => {
  const workbook = new Workbook();
  workbook.addWorksheet('Data').getCell('A1').value = {formula: '1+1', result: 2};
  const parts = unzipSync(writeXlsx(workbook));
  parts['xl/calcChain.xml'] = strToU8(
    '<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><c r="A1" i="7"/></calcChain>',
  );
  const original = patchParts(zipSync(parts), {
    'xl/_rels/workbook.xml.rels': (xml) =>
      xml.replace(
        '</Relationships>',
        '<Relationship Id="rIdChain" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" Target="calcChain.xml"/></Relationships>',
      ),
    '[Content_Types].xml': (xml) =>
      xml.replace(
        '</Types>',
        '<Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/></Types>',
      ),
  });
  const loaded = readXlsx(original);
  loaded.requireWorksheet('Data').getCell('A1').value = 5;
  const output = writeXlsx(loaded);
  assert.doesNotMatch(partMatching(output, /^xl\/_rels\/workbook.xml.rels$/), /calcChain/);
  assert.doesNotMatch(partMatching(output, /^\[Content_Types\].xml$/), /calcChain/);
  assert.match(partMatching(output, /^xl\/workbook.xml$/), /fullCalcOnLoad="1"/);
});

test('picture movement in a mixed drawing preserves charts, shapes and their relationships', () => {
  const workbook = nativeFeatureWorkbook();
  const original = writeXlsx(workbook);
  const baseline = writeXlsx(workbook);
  const sheet = workbook.requireWorksheet('Features');
  const image = sheet.images[0]!;
  sheet.removeImage(image.imageId);
  sheet.addImage(image.imageId, {tl: {col: 2, row: 10}, ext: {width: 80, height: 60}});
  const result = mergeXlsxEdits(original, baseline, writeXlsx(workbook), {
    sheets: [{name: 'Features', originalName: 'Features', baselineName: 'Features'}],
  });
  assertRelationshipsWired(result.data);
  const loaded = readXlsx(result.data).requireWorksheet('Features');
  assert.equal(loaded.charts.length, 1);
  assert.equal(loaded.shapes.length, 1);
  const drawing = xmlTree(partMatching(result.data, /^xl\/drawings\/[^/]+\.xml$/));
  const picture = xmlChildren(drawing).find((node) => xmlChild(node, 'pic'));
  assert.equal(xmlText(xmlChild(picture, 'from', 'row')), '10');
  assert.equal(xmlChildren(drawing).length, 3);
});

test('chart formatting outside the typed model and differential-style identities survive editor merging', () => {
  const workbook = new Workbook(),
    sheet = workbook.addWorksheet('Data');
  sheet.getCell('A1').value = 1;
  sheet.addChart(
    {type: 'column', title: 'Sales', series: [{name: 'S', values: {values: [1]}}]},
    {tl: {col: 2, row: 0}, ext: {width: 200, height: 100}},
  );
  workbook[INTERNAL].restoreDifferentialStyles(['<dxf><font><b/></font></dxf>']);
  const original = writeXlsx(workbook);
  const baseline = writeXlsx(workbook);
  const edited = patchParts(baseline, {
    'xl/charts/chart1.xml': (xml) =>
      xml.replace(
        '</c:chartSpace>',
        '<c:spPr><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></c:spPr></c:chartSpace>',
      ),
  });
  const current = readXlsx(edited);
  current[INTERNAL].restoreDifferentialStyles(['<dxf><font><i/></font></dxf>']);
  current.requireWorksheet('Data').addConditionalFormatting({
    ref: 'A1',
    rules: [{type: 'expression', formulae: ['A1>0'], dxfId: '0'}],
  });
  const result = mergeXlsxEdits(original, baseline, writeXlsx(current), {
    sheets: [{name: 'Data', originalName: 'Data', baselineName: 'Data'}],
  });
  assert.match(partMatching(result.data, /^xl\/charts\/[^/]+\.xml$/), /FF0000/);
  const out = readXlsx(result.data),
    rule = out.requireWorksheet('Data').conditionalFormattings[0]!.rules[0]!;
  assert.ok('dxfId' in rule && rule.dxfId !== undefined);
  assert.match(out.differentialStyles[Number(rule.dxfId)]!, /<i\//);
});

test('editor baseline omissions cannot delete original features or flatten an untouched theme style', () => {
  const source = nativeFeatureWorkbook();
  source.requireWorksheet('Features').getCell('A2').font = {
    name: 'Aptos',
    bold: true,
    color: {theme: 4},
  };
  const original = patchParts(writeXlsx(source), {
    'xl/worksheets/sheet1.xml': (xml) =>
      xml.replace(
        '</worksheet>',
        '<ignoredErrors><ignoredError sqref="A2" numberStoredAsText="1"/></ignoredErrors></worksheet>',
      ),
  });
  const editor = new Workbook(),
    sheet = editor.addWorksheet('Features');
  sheet.addRows([
    ['A', 'B'],
    [1, 2],
    [3, 4],
  ]);
  sheet.getCell('A2').font = {name: 'Aptos', bold: true, color: {argb: 'FF4472C4'}};
  const baseline = writeXlsx(editor);
  sheet.getCell('A2').value = 12;
  const result = mergeXlsxEdits(original, baseline, writeXlsx(editor), {
    sheets: [{name: 'Features', originalName: 'Features', baselineName: 'Features'}],
  });
  assertRelationshipsWired(result.data);
  const loaded = readXlsx(result.data).requireWorksheet('Features');
  assert.equal(loaded.getCell('A2').value, 12);
  assert.deepEqual(loaded.getCell('A2').font?.color, {theme: 4});
  assert.equal(loaded.cellImages.length, 1);
  assert.equal(loaded.shapes.length, 1);
  assert.equal(loaded.charts.length, 1);
  assert.equal(loaded.sparklineGroups.length, 1);
  assert.deepEqual(loaded.checkboxes, ['C4']);
  assert.match(
    partMatching(result.data, /^xl\/worksheets\/sheet1.xml$/),
    /ignoredError sqref="A2"/,
  );
  const chart = loaded.charts[0];
  assert.ok(chart?.kind === 'chart');
  assert.deepEqual(chart.options.series[0]?.values.values, [12, 3]);
});

test('stable sheet mapping supports swapped names, reordered tabs, new tabs and deleted tabs', () => {
  const original = new Workbook();
  original.addWorksheet('A').getCell('A1').value = 'A';
  original.addWorksheet('B').getCell('A1').value = 'B';
  original.addWorksheet('Deleted');
  const baseline = writeXlsx(original),
    edited = new Workbook();
  edited.addWorksheet('A').getCell('A1').value = 'B edited';
  edited.addWorksheet('B').getCell('A1').value = 'A edited';
  edited.addWorksheet('New').getCell('A1').value = 'new';
  const result = mergeXlsxEdits(baseline, baseline, writeXlsx(edited), {
    sheets: [
      {name: 'A', originalName: 'B', baselineName: 'B'},
      {name: 'B', originalName: 'A', baselineName: 'A'},
      {name: 'New'},
    ],
  });
  const loaded = readXlsx(result.data);
  assert.deepEqual(
    loaded.worksheets.map((sheet) => sheet.name),
    ['A', 'B', 'New'],
  );
  assert.equal(loaded.requireWorksheet('A').getCell('A1').value, 'B edited');
  assert.equal(loaded.requireWorksheet('B').getCell('A1').value, 'A edited');
  assert.ok(
    result.diagnostics.some((item) => item.code === 'sheet-deleted' && item.sheet === 'Deleted'),
  );
});

test('native features follow insertions and deletion of their owning cells', () => {
  const workbook = nativeFeatureWorkbook(),
    sheet = workbook.requireWorksheet('Features');
  sheet.spliceRows(2, 0, []);
  assert.equal(sheet.cellImages[0]?.ref, 'C3');
  assert.deepEqual(sheet.checkboxes, ['C5']);
  assert.equal(sheet.sparklineGroups[0]?.sparklines[0]?.ref, 'B6');
  assert.equal(sheet.sparklineGroups[0]?.sparklines[0]?.formula, 'Features!A3:A4');
  const loaded = readXlsx(writeXlsx(workbook)).requireWorksheet('Features');
  assert.equal(loaded.cellImages[0]?.ref, 'C3');
  loaded.spliceRows(3, 1);
  assert.equal(loaded.cellImages.length, 0);
});
