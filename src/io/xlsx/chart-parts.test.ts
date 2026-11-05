import assert from 'node:assert/strict';
import {test} from 'node:test';

import type {ChartOptions, ChartType} from '../../core/chart.ts';
import {Workbook} from '../../core/workbook.ts';
import {resolveRelativePart} from '../opc/part-paths.ts';
import {parseRelationshipRecords} from '../opc/read-opc.ts';
import {chartXml} from './chart-xml.ts';
import {readXlsx} from './read.ts';
import {buildPackageParts, writeXlsx, writeXlsxAsync} from './write.ts';

const anchor = {tl: {col: 3, row: 1}, ext: {width: 480, height: 320}};
const options: ChartOptions = {
  type: 'column',
  title: 'Sales & <profit>',
  categories: {formula: "'Data'!$A$2:$A$4", values: ['A', 'B', 'C']},
  series: [
    {name: 'Sales', values: {formula: "'Data'!$B$2:$B$4", values: [12, null, 0]}, color: '#123456'},
  ],
};
const text = (parts: Record<string, Uint8Array>, path: string): string =>
  new TextDecoder().decode(parts[path]);

test('authored charts have linked parts, cached data and escaped titles', () => {
  const workbook = new Workbook();
  const sheet = workbook.addWorksheet('Data');
  sheet.addChart(options, anchor);
  const parts = buildPackageParts(workbook);
  const chart = text(parts, 'xl/charts/chart1.xml');
  assert.match(chart, /Sales &amp; &lt;profit&gt;/);
  assert.match(chart, /<c:numRef><c:f>'Data'!\$B\$2:\$B\$4<\/c:f>/);
  assert.match(
    chart,
    /<c:ptCount val="3"\/><c:pt idx="0"><c:v>12<\/c:v><\/c:pt><c:pt idx="2"><c:v>0<\/c:v>/,
  );
  assert.match(text(parts, 'xl/drawings/drawing1.xml'), /<xdr:oneCellAnchor/);
  assert.match(text(parts, 'xl/drawings/drawing1.xml'), /cx="4572000" cy="3048000"/);
  for (const [path, bytes] of Object.entries(parts)) {
    if (!path.endsWith('.rels')) continue;
    const owner = path === '_rels/.rels' ? '' : path.replace('/_rels/', '/').slice(0, -5);
    for (const rel of parseRelationshipRecords(new TextDecoder().decode(bytes))) {
      if (!rel.external)
        assert.ok(parts[resolveRelativePart(owner, rel.target)], `${path}: ${rel.target}`);
    }
  }
});

test('all chart families, combinations and secondary axes serialize', () => {
  const types: ChartType[] = [
    'column',
    'bar',
    'line',
    'area',
    'pie',
    'doughnut',
    'scatter',
    'bubble',
    'radar',
  ];
  for (const type of types) {
    const xml = chartXml({
      type,
      showValues: true,
      series: [
        {
          name: 'S',
          values: {values: [1, 2]},
          xValues: {values: [0, 1]},
          bubbleSizes: {values: [3, 4]},
        },
      ],
    });
    assert.match(xml, new RegExp(`<c:${type === 'column' ? 'bar' : type}Chart>`));
  }
  const xml = chartXml({
    ...options,
    series: [
      ...options.series,
      {name: 'Margin', type: 'line', secondaryAxis: true, values: {values: [1, 2, 3]}},
    ],
    secondaryValueAxis: {min: 0, max: 10, majorUnit: 2, title: 'Margin'},
  });
  assert.match(xml, /<c:lineChart>/);
  assert.match(xml, /<c:axId val="4"\/>/);
  assert.match(xml, /<c:crossAx val="3"\/>/);
  assert.match(xml, /<c:majorUnit val="2"\/>/);
});

test('images, loaded charts and newly authored charts share one drawing without collisions', () => {
  const workbook = new Workbook();
  const sheet = workbook.addWorksheet('Data');
  sheet.addChart(options, anchor);
  const image = workbook.addImage({
    buffer: new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10]),
    extension: 'png',
  });
  sheet.addImage(image, anchor);
  const loaded = readXlsx(writeXlsx(workbook));
  const target = loaded.worksheets[0]!;
  target.addChart({...options, title: 'Second'}, {tl: {col: 0, row: 20}, br: {col: 8, row: 35}});
  const parts = buildPackageParts(loaded);
  assert.equal(
    Object.keys(parts).filter((path) => /^xl\/charts\/chart\d+.xml$/.test(path)).length,
    2,
  );
  const drawingPaths = Object.keys(parts).filter((path) => /^xl\/drawings\/[^/]+.xml$/.test(path));
  assert.equal(drawingPaths.length, 1);
  const drawing = text(parts, drawingPaths[0]!);
  assert.equal([...drawing.matchAll(/<c:chart\b/g)].length, 2);
  const ids = [...drawing.matchAll(/<[^:>]+:cNvPr id="(\d+)"/g)].map((match) => match[1]);
  assert.equal(new Set(ids).size, ids.length);
  const reread = readXlsx(writeXlsx(loaded));
  assert.equal(reread.worksheets.length, 1);
});

test('chart ownership survives caller mutation and model replacement', () => {
  const workbook = new Workbook();
  const sheet = workbook.addWorksheet('Data');
  const values = [1, 2];
  sheet.addChart({type: 'line', series: [{name: 'S', values: {values}}]}, anchor);
  values[0] = 999;
  const copy = workbook.addWorksheet('Copy');
  copy.model = sheet.model;
  sheet.removeChart(0);
  assert.equal(copy.charts.length, 1);
  assert.match(text(buildPackageParts(workbook), 'xl/charts/chart1.xml'), /<c:v>1<\/c:v>/);
  copy.model = sheet.model;
  assert.equal(copy.charts.length, 0);
  assert.throws(() => copy.removeChart(0), /outside/);
});

test('complete chart XML can bring a related drawing and remains editable', () => {
  const workbook = new Workbook();
  const sheet = workbook.addWorksheet('Data');
  sheet.addChartPart({xml: chartXml(options)}, {from: {col: 0, row: 0}, to: {col: 5, row: 10}});
  assert.match(text(buildPackageParts(workbook), 'xl/charts/chart1.xml'), /<c:barChart>/);
  sheet.addChartPart({xml: '<chartSpace/>'}, {from: {col: 0, row: 0}, to: {col: 5, row: 10}});
  assert.throws(() => writeXlsx(workbook), /namespace/);
});

test('invalid data and geometry fail before a corrupt package is returned', () => {
  const workbook = new Workbook();
  const sheet = workbook.addWorksheet('Data');
  assert.throws(() => sheet.addChart({type: 'line', series: []}, anchor), /at least one/);
  assert.throws(
    () => sheet.addChart({type: 'scatter', series: [{name: 'S', values: {values: [1]}}]}, anchor),
    /xValues/,
  );
  assert.throws(
    () =>
      sheet.addChart(
        {type: 'column', series: [{name: 'S', values: {values: [Number.NaN]}}]},
        anchor,
      ),
    /finite/,
  );
  sheet.addChart(options, {tl: {col: -1, row: 0}, ext: {width: 480, height: 320}});
  assert.throws(() => writeXlsx(workbook), /anchor/);
});

test('async and sync chart writers produce identical packages', async () => {
  const workbook = new Workbook();
  workbook.addWorksheet('Data').addChart(options, anchor);
  assert.deepEqual(await writeXlsxAsync(workbook), writeXlsx(workbook));
});

test('structural edits move chart references and anchors with their owning sheets', () => {
  const workbook = new Workbook();
  const data = workbook.addWorksheet('Data');
  const report = workbook.addWorksheet('Report');
  report.addChart(options, anchor);
  report.addChartPart({xml: chartXml(options)}, {from: {col: 3, row: 1}, to: {col: 8, row: 20}});
  data.insertRow(2, [42]);
  report.insertRow(1, []);
  const parts = buildPackageParts(workbook);
  for (const number of [1, 2]) {
    assert.match(text(parts, `xl/charts/chart${number}.xml`), /'Data'!\$B\$3:\$B\$5/);
  }
  assert.equal(report.charts[0]!.anchor.from.row, 2);
  assert.equal(report.charts[1]!.anchor.from.row, 2);
});
