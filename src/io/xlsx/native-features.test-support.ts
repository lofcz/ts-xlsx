import {Workbook} from '../../core/workbook.ts';

export function nativeFeatureWorkbook(): Workbook {
  const workbook = new Workbook(),
    sheet = workbook.addWorksheet('Features');
  sheet.addRows([
    ['A', 'B'],
    [1, 2],
    [3, 4],
  ]);
  sheet.addConditionalFormatting({
    ref: 'A2:A3',
    rules: [
      {
        type: 'dataBar',
        cfvo: [{type: 'autoMin'}, {type: 'autoMax'}],
        color: {argb: 'FF336699'},
        gradient: true,
        negativeFillColor: {argb: 'FFFF0000'},
      },
    ],
  });
  const image = workbook.addImage({
    extension: 'png',
    buffer: Uint8Array.from(
      Buffer.from(
        'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR4nGMwTpv5HwAENAIyWy0K4AAAAABJRU5ErkJggg==',
        'base64',
      ),
    ),
  });
  sheet.addImage(image, {tl: {col: 5, row: 1}, ext: {width: 40, height: 40}});
  sheet.addCellImage(image, 'C2', 'A tiny image');
  sheet.getCell('D2').value = {
    formula: 'SEQUENCE(2)',
    shareType: 'array',
    ref: 'D2:D3',
    dynamic: true,
    result: 1,
  };
  sheet.getCell('E2').value = {error: '#SPILL!'};
  sheet.addCheckbox('C4', true);
  sheet.getCell('C4').font = {bold: true};
  sheet.addSparklineGroup({
    type: 'line',
    sparklines: [{ref: 'B5', formula: 'Features!A2:A3'}],
    markers: true,
    colors: {series: 'FF123456'},
  });
  sheet.addShape(
    {
      geometry: 'roundRect',
      fill: '#123456',
      text: [{text: 'Rich ', font: {bold: true}}, {text: 'text'}],
      textBox: true,
    },
    {tl: {col: 7, row: 1}, ext: {width: 120, height: 60}},
  );
  sheet.addChart(
    {
      type: 'column',
      categories: {values: ['A', 'B']},
      series: [{name: 'S', values: {formula: 'Features!A2:A3', values: [1, 3]}}],
    },
    {tl: {col: 9, row: 1}, ext: {width: 320, height: 240}},
  );
  return workbook;
}
