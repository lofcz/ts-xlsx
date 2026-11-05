import assert from 'node:assert/strict';
import {test} from 'node:test';

import type {ChartOptions, ChartType} from '../../core/chart.ts';
import {Workbook} from '../../core/workbook.ts';
import {assertRelationshipsWired, partMatching} from './package.test-support.ts';
import {readXlsx} from './read.ts';
import {writeXlsx} from './write.ts';

const anchor = {tl: {col: 0, row: 0}, ext: {width: 640, height: 400}};
test('date/log axes, trendlines, custom error bars and labels remain editable after import', () => {
  const workbook = new Workbook(),
    sheet = workbook.addWorksheet('Data');
  const options: ChartOptions = {
    type: 'line',
    title: 'Analysis',
    categories: {values: [46000, 46001, 46002], formatCode: 'yyyy-mm-dd'},
    categoryAxis: {type: 'date', baseTimeUnit: 'days', majorTimeUnit: 'days', majorUnit: 1},
    valueAxis: {logBase: 10, min: 1, max: 1000},
    dataLabels: {showCategory: true, separator: ' / ', formatCode: '0.0', position: 't'},
    series: [
      {
        name: 'Values',
        values: {values: [10, 20, 30]},
        trendlines: [{type: 'poly', order: 2, displayEquation: true, forward: 3}],
        errorBars: [{type: 'cust', plus: {values: [1, 2, 3]}, minus: {values: [2, 3, 4]}}],
      },
    ],
  };
  sheet.addChart(options, anchor);
  const loaded = readXlsx(writeXlsx(workbook));
  const chart = loaded.requireWorksheet('Data').charts[0];
  assert.ok(chart?.kind === 'chart');
  assert.equal(chart.options.categoryAxis?.type, 'date');
  assert.equal(chart.options.valueAxis?.logBase, 10);
  assert.equal(chart.options.series[0]?.trendlines?.[0]?.order, 2);
  assert.deepEqual(chart.options.series[0]?.errorBars?.[0]?.plus?.values, [1, 2, 3]);
  assert.equal(chart.options.dataLabels?.separator, ' / ');
  loaded.requireWorksheet('Data').updateChart(0, {title: 'Edited'});
  const bytes = writeXlsx(loaded);
  assertRelationshipsWired(bytes);
  const xml = partMatching(bytes, /^xl\/charts\/[^/]+\.xml$/);
  assert.match(xml, /<c:dateAx>/);
  assert.match(xml, /Edited/);
});

test('Office 2016 families round-trip and edit the original chartEx part', () => {
  const types: ChartType[] = [
    'waterfall',
    'histogram',
    'pareto',
    'funnel',
    'treemap',
    'sunburst',
    'boxWhisker',
  ];
  for (const type of types) {
    const workbook = new Workbook(),
      sheet = workbook.addWorksheet('Data');
    sheet.addChart(
      {
        type,
        series: [
          {name: 'S', values: {values: [10, 20, 30]}, categories: {values: ['A', 'B', 'C']}},
        ],
        ...(type === 'histogram' ? {binning: {width: 5}} : {}),
        ...(type === 'waterfall' ? {waterfallTotals: [2]} : {}),
      },
      anchor,
    );
    const bytes = writeXlsx(workbook);
    assertRelationshipsWired(bytes);
    const loaded = readXlsx(bytes),
      target = loaded.requireWorksheet('Data');
    assert.equal(target.charts.length, 1, type);
    const chart = target.charts[0];
    assert.ok(chart?.kind === 'chart');
    assert.equal(chart.options.type, type);
    target.updateChart(0, {title: 'Changed'});
    const out = writeXlsx(loaded);
    assert.match(partMatching(out, /^xl\/charts\/[^/]+\.xml$/), /Changed/);
    assertRelationshipsWired(out);
  }
});

test('chart sheets retain their tab kind and chart through two read/write cycles', () => {
  const workbook = new Workbook();
  workbook.addWorksheet('Data').getCell('A1').value = 8;
  workbook.addChartSheet('Dashboard', {
    type: 'column',
    series: [{name: 'S', values: {values: [8]}}],
  });
  let bytes = writeXlsx(workbook);
  for (let i = 0; i < 2; i++) {
    assertRelationshipsWired(bytes);
    assert.match(partMatching(bytes, /^xl\/chartsheets\/[^/]+\.xml$/), /<chartsheet/);
    const loaded = readXlsx(bytes);
    assert.equal(loaded.requireWorksheet('Dashboard').kind, 'chartsheet');
    assert.equal(loaded.requireWorksheet('Dashboard').charts.length, 1);
    bytes = writeXlsx(loaded);
  }
});

test('unresolved formula caches are retained and reported without pretending to calculate', () => {
  const workbook = new Workbook(),
    sheet = workbook.addWorksheet('Data');
  sheet.getCell('A1').value = {formula: 'SUM(B1:B3)'};
  sheet.addChart(
    {type: 'column', series: [{name: 'S', values: {formula: 'Data!A1', values: [10]}}]},
    anchor,
  );
  const diagnostics: import('../../core/chart-cache.ts').ChartCacheDiagnostic[] = [];
  const bytes = writeXlsx(workbook, {chartCacheDiagnostics: diagnostics});
  assert.equal(diagnostics[0]?.reason, 'uncalculated-formula');
  assert.match(partMatching(bytes, /^xl\/charts\/[^/]+\.xml$/), /<c:v>10<\/c:v>/);
});
