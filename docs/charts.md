# Authoring charts

The `@lofcz/ts-xlsx` fork adds native, editable column, bar, line, area, pie,
doughnut, scatter, bubble, radar and stock charts, plus waterfall, histogram,
Pareto, funnel, treemap, sunburst and box-and-whisker chartEx charts. Column/line/area series can be combined
and assigned a secondary axis. Titles, legends, colors, data labels, axis bounds,
number formats, literal data and worksheet references are supported.

```ts
import {Workbook, writeXlsx} from '@lofcz/ts-xlsx';

const workbook = new Workbook();
const sheet = workbook.addWorksheet('Sales');
sheet.addRows([['Month', 'Revenue'], ['Jan', 120], ['Feb', 150]]);
sheet.addChart({
  type: 'column',
  title: 'Revenue',
  categories: {formula: 'Sales!$A$2:$A$3', values: ['Jan', 'Feb']},
  series: [{
    name: 'Revenue',
    values: {formula: 'Sales!$B$2:$B$3', values: [120, 150]},
    color: '#336699',
  }],
  valueAxis: {title: 'EUR', min: 0},
}, {tl: {col: 3, row: 1}, ext: {width: 480, height: 320}});
const bytes = writeXlsx(workbook);
```

`values` supplies cached chart data. `formula` optionally links that data to
worksheet cells. Writing refreshes caches from cell values, cached formula results
and defined names; it does not evaluate formulas. Unresolved caches are retained
and reported through `WriteOptions.chartCacheDiagnostics`. Null values
are gaps, and all series arrays must have matching lengths. Scatter charts require
`xValues`; bubble charts additionally require `bubbleSizes`.

Anchors use the same fractional-cell/pixel input as `addImage`. The library copies
chart options and anchors when they are added. `sheet.charts` returns snapshots;
`updateChart(index, patch)` changes imported or authored options, and
`removeChart(index)` removes the chart. Worksheet model copies include
authored charts. Row and column edits move their anchors and formula references.

`sheet.addChartPart(part, anchor)` accepts a complete standard DrawingML chart XML
part for producers with more elaborate models, including chart-related XML parts
such as user-shape drawings. This lower-level anchor uses zero-based cells and EMU
offsets, like `addImageAnchor`. The writer validates XML well-formedness and owns
package paths, content types, relationship IDs and placement in the worksheet's
single drawing. The caller owns the supplied XML's schema conformance.

Supported imported charts expose `ChartOptions` while retaining their original
XML. Edits update the existing part and retain formatting outside the typed model.
Unsupported charts remain preserved parts. New charts share the worksheet drawing
with pictures and shapes. Both buffered and streaming writers use that assembly.

`Workbook.addChartSheet(name, options, {width, height})` creates a chart-only tab.
Standard charts support numeric date categories, date units, logarithmic value
axes, trendlines, custom/fixed/percentage/statistical error bars, per-series labels
and point colors. Stock series are ordered high/low/close or open/high/low/close.
ChartEx adds hierarchy category levels, waterfall totals/connectors, histogram
binning and box-and-whisker display options. Its axes support bounds and units;
the chartEx schema has no date or logarithmic scaling.

See [workbook fidelity](workbook-fidelity.md) for native drawing features,
preserving editor imports, diagnostics, and interoperability checks.
