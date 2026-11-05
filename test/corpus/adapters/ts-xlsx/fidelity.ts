import {readFixture, readXlsx, writeXlsx} from './runtime.ts';

export const fidelity = {
  combinedFeaturesThroughSaves(fixture: string, cycles: number) {
    let workbook = readFixture(fixture);
    const counts = [];
    for (let cycle = 0; cycle < cycles; cycle++) {
      workbook = readXlsx(writeXlsx(workbook));
      const sheet = workbook.requireWorksheet('Features');
      const chart = sheet.charts[0];
      counts.push({
        charts: sheet.charts.length,
        shapes: sheet.shapes.length,
        cellImages: sheet.cellImages.length,
        sparklines: sheet.sparklineGroups.length,
        checkboxes: sheet.checkboxes,
        conditionalFormats: sheet.conditionalFormattings.length,
        chartValues: chart?.kind === 'chart' ? chart.options.series[0]?.values.values : undefined,
        array: sheet.getCell('D2').value,
        error: sheet.getCell('E2').value,
      });
    }
    return counts;
  },
};
