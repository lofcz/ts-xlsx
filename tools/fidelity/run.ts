import assert from 'node:assert/strict';
import {spawnSync} from 'node:child_process';
import {mkdirSync, readFileSync, readdirSync, writeFileSync} from 'node:fs';
import {join, resolve} from 'node:path';
import {pathToFileURL} from 'node:url';

import {scratchDir} from '../../scripts/repo.ts';
import {readXlsx} from '../../src/io/xlsx/read.ts';
import {writeXlsx} from '../../src/io/xlsx/write.ts';

const output = scratchDir('fidelity-render');
const fixture = resolve(
  process.argv[2] ?? 'test/corpus/fixtures/combined-native-features/sample.xlsx',
);
const original = readFileSync(fixture);
writeFileSync(join(output, 'original.xlsx'), original);
writeFileSync(join(output, 'rewritten.xlsx'), writeXlsx(readXlsx(original)));
const run = (command: string, args: string[]) => {
  const result = spawnSync(command, args, {encoding: 'utf8', timeout: 60000});
  if (result.error || result.status !== 0 || /Error:|failed/i.test(result.stderr))
    throw new Error(`${command}: ${result.error?.message ?? result.stderr ?? result.stdout}`);
  return result.stdout;
};
const version = run('libreoffice', ['--version']).trim();
const profile = `-env:UserInstallation=${pathToFileURL(join(output, 'lo-profile')).href}`;
for (const name of ['original', 'rewritten']) {
  run('libreoffice', [
    profile,
    '--headless',
    '--convert-to',
    'pdf:calc_pdf_Export',
    '--outdir',
    output,
    join(output, `${name}.xlsx`),
  ]);
  assert.ok(readFileSync(join(output, `${name}.pdf`)).length > 100);
  run('pdftoppm', ['-r', '96', '-png', join(output, `${name}.pdf`), join(output, name)]);
}
const pages = readdirSync(output).filter((name) => /^original-\d+\.png$/.test(name));
assert.ok(pages.length > 0);
assert.equal(
  readdirSync(output).filter((name) => /^rewritten-\d+\.png$/.test(name)).length,
  pages.length,
);
const comparisons = pages.map((name) => {
  const other = name.replace('original-', 'rewritten-');
  const comparison = spawnSync(
    'compare',
    [
      '-metric',
      'AE',
      join(output, name),
      join(output, other),
      join(output, name.replace('original-', 'diff-')),
    ],
    {encoding: 'utf8', timeout: 60000},
  );
  if (comparison.error || comparison.status === null || comparison.status > 1)
    throw new Error(comparison.error?.message ?? comparison.stderr);
  return {page: name, changedPixels: Number.parseFloat(comparison.stderr.trim())};
});
const saved = join(output, 'lo-resaved');
mkdirSync(saved);
run('libreoffice', [
  profile,
  '--headless',
  '--convert-to',
  'xlsx:Calc MS Excel 2007 XML',
  '--outdir',
  saved,
  join(output, 'rewritten.xlsx'),
]);
const features = (bytes: Uint8Array) =>
  readXlsx(bytes).worksheets.map((sheet) => ({
    name: sheet.name,
    charts: sheet.charts.length,
    shapes: sheet.shapes.length,
    cellImages: sheet.cellImages.length,
    sparklines: sheet.sparklineGroups.length,
    checkboxes: sheet.checkboxes.length,
  }));
const report = {
  application: version,
  fixture,
  comparisons,
  before: features(original),
  afterLibreOfficeSave: features(readFileSync(join(saved, 'rewritten.xlsx'))),
  excelDesktop: 'not run',
  repairCheck:
    'LibreOffice command failures checked; Excel repair detection requires the Windows COM oracle',
};
writeFileSync(join(output, 'report.json'), JSON.stringify(report, null, 2));
process.stdout.write(JSON.stringify({output, ...report}, null, 2) + '\n');
assert.ok(
  comparisons.every((page) => page.changedPixels === 0),
  'the library rewrite changed rendered pixels; inspect the diff images',
);
