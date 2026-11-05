import assert from 'node:assert/strict';
import {test} from 'node:test';

import {Workbook} from '../../core/workbook.ts';
import {partMatching, patchParts, assertRelationshipsWired} from './package.test-support.ts';
import {readXlsx} from './read.ts';
import {writeXlsx} from './write.ts';

test('imported chartEx anchors move, change chart vocabulary and delete without stale objects', () => {
  const workbook = new Workbook();
  workbook
    .addWorksheet('Data')
    .addChart(
      {type: 'waterfall', series: [{name: 'Sales', values: {values: [1, 2]}}]},
      {tl: {col: 0, row: 0}, ext: {width: 300, height: 200}},
    );
  const loaded = readXlsx(writeXlsx(workbook));
  const sheet = loaded.requireWorksheet('Data');
  const model = sheet.model;
  const chart = model.charts[0]!;
  sheet.model = {...model, charts: [{...chart, anchor: {...chart.anchor, from: {col: 5, row: 4}}}]};
  let next = readXlsx(writeXlsx(loaded)).requireWorksheet('Data').charts[0]!;
  assert.equal(next.anchor.from.col, 5);
  sheet.updateChart(0, {type: 'column'});
  let bytes = writeXlsx(loaded);
  assertRelationshipsWired(bytes);
  next = readXlsx(bytes).requireWorksheet('Data').charts[0]!;
  assert.ok(next.kind === 'chart');
  assert.equal(next.options.type, 'column');
  sheet.updateChart(0, {type: 'funnel'});
  bytes = writeXlsx(loaded);
  assertRelationshipsWired(bytes);
  sheet.removeChart(0);
  bytes = writeXlsx(loaded);
  assertRelationshipsWired(bytes);
  assert.equal(readXlsx(bytes).requireWorksheet('Data').charts.length, 0);
});

test('model replacement removes old chart anchors and copies imported charts to another sheet', () => {
  const workbook = new Workbook();
  const original = workbook.addWorksheet('Source');
  original.addChart(
    {type: 'line', series: [{name: 'A', values: {values: [1]}}]},
    {tl: {col: 1, row: 1}, ext: {width: 300, height: 200}},
  );
  const loaded = readXlsx(writeXlsx(workbook));
  const source = loaded.requireWorksheet('Source');
  loaded.addWorksheet('Copy').model = source.model;
  source.model = {...source.model, charts: []};
  const bytes = writeXlsx(loaded);
  assertRelationshipsWired(bytes);
  const out = readXlsx(bytes);
  assert.equal(out.requireWorksheet('Source').charts.length, 0);
  assert.equal(out.requireWorksheet('Copy').charts.length, 1);
});

test('a rejected chart sheet cannot leave a tab behind', () => {
  const workbook = new Workbook();
  assert.throws(() => workbook.addChartSheet('Invalid', {type: 'column', series: []}));
  assert.equal(workbook.worksheets.length, 0);
});

test('loaded charts edit their existing part, refresh cell caches and retain foreign formatting', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('Data');
  sheet.addRows([
    ['Name', 'Value'],
    ['A', 5],
    ['B', 6],
  ]);
  sheet.addChart(
    {
      type: 'column',
      title: 'Before',
      categories: {formula: 'Data!A2:A3', values: ['A', 'B']},
      series: [{name: 'Values', values: {formula: 'Data!B2:B3', values: [5, 6]}}],
    },
    {tl: {col: 3, row: 0}, ext: {width: 300, height: 200}},
  );
  const bytes = patchParts(writeXlsx(wb), {
    'xl/charts/chart1.xml': (xml) =>
      xml.replace(
        '</c:chartSpace>',
        '<c:extLst><c:ext uri="urn:custom"><custom:payload xmlns:custom="urn:custom">kept</custom:payload></c:ext></c:extLst></c:chartSpace>',
      ),
  });
  const loaded = readXlsx(bytes);
  const target = loaded.requireWorksheet('Data');
  assert.equal(target.charts.length, 1);
  target.updateChart(0, {title: 'After'});
  target.getCell('B2').value = 19;
  const out = writeXlsx(loaded);
  const xml = partMatching(out, /^xl\/charts\/[^/]+\.xml$/);
  assert.match(xml, /After/);
  assert.match(xml, /<c:v>19<\/c:v>/);
  assert.match(xml, /custom:payload/);
  assertRelationshipsWired(out);
  target.removeChart(0);
  assert.equal(readXlsx(writeXlsx(loaded)).requireWorksheet('Data').charts.length, 0);
});
