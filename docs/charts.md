# Authoring charts

The `@lofcz/ts-xlsx` fork adds native, editable column, bar, line, area, pie,
doughnut, scatter, bubble and radar charts. Column/line/area series can be combined
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
worksheet cells; this library does not evaluate the referenced cells. Null values
are gaps, and all series arrays must have matching lengths. Scatter charts require
`xValues`; bubble charts additionally require `bubbleSizes`.

Anchors use the same fractional-cell/pixel input as `addImage`. The library copies
chart options and anchors when they are added. `sheet.charts` returns snapshots;
`removeChart(index)` removes an authored chart. Worksheet model copies include
authored charts. Row and column edits move their anchors and formula references.

`sheet.addChartPart(part, anchor)` accepts a complete standard DrawingML chart XML
part for producers with more elaborate models, including chart-related XML parts
such as user-shape drawings. This lower-level anchor uses zero-based cells and EMU
offsets, like `addImageAnchor`. The writer validates XML well-formedness and owns
package paths, content types, relationship IDs and placement in the worksheet's
single drawing. The caller owns the supplied XML's schema conformance.

Existing charts remain preserved XML on import; they are not converted into
`ChartOptions`. New charts can join that same drawing alongside images without
overwriting its parts or relationships. Both buffered writers and the streaming
writer share chart package assembly. Office 2016 chart extensions and chart sheets
are not part of the typed authoring API.
