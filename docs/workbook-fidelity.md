# Workbook fidelity

The fork owns drawings, media, relationship IDs, styles and cell metadata together.
The following APIs can coexist in one workbook without competing ZIP patches:

```ts
sheet.addCellImage(workbook.addImage({buffer: pngBytes, extension: 'png'}), 'C2', 'Logo');
sheet.addCheckbox('D2', true);
sheet.addSparklineGroup({
  type: 'line', markers: true,
  sparklines: [{ref: 'E2', formula: 'Data!A2:A12'}],
  colors: {series: 'FF336699'},
});
sheet.addShape({text: 'Quarterly revenue', textBox: true, fill: 'EEF4FF'}, {
  tl: {col: 6, row: 1}, ext: {width: 220, height: 60},
});
```

Use `removeCellImage`, `removeCheckbox`, `removeSparklineGroup` and `removeShape`
to remove those features. `addDrawingXml` accepts complete shape anchors for
producers with richer DrawingML models. Imported shapes retain their raw markup.
Checkbox formatting does not coerce a pre-existing non-Boolean value. Cell images
use Excel rich-value metadata, sharing its indices with dynamic arrays and errors.
Structural edits move the modeled anchors, cell locations and formula references.

## Editing imported workbooks

An editor that represents only part of Excel's model must keep both the original
package and its first export. `mergeXlsxEdits` applies changes between that baseline
and the current export to the original. A feature missing from both editor exports
is retained, rather than interpreted as a user deletion.

```ts
import {mergeXlsxEdits} from '@lofcz/ts-xlsx/xlsx';
const {data, diagnostics} = mergeXlsxEdits(original, baseline, current, {
  sheets: [{name: 'Renamed', originalName: 'Data', baselineName: 'Data'}],
});
```

The sheet mapping must come from stable editor identities; its order is the final
tab order. New tabs omit source/baseline names. Omitted original tabs are deleted.
Preserved chart parts are edited in place, caches refresh, theme-linked untouched
styles remain linked, and workbook/worksheet extension markup and unconsumed
relationship closures survive. Cell images, differential styles and comment people
are transferred into the destination workbook's registries.

Diagnostics report renamed/deleted sheets, picture replacement when the editor
exposed fewer picture anchors, and caches retained because no authoritative value
was available. Surface these to users. Opaque objects can retain references to old
sheet names, and unsupported embedded objects cannot be fully edited through a
partial editor. This is preservation with explicit limits, not a claim of complete
Excel feature parity. Original packages and baselines contain workbook content;
persist them with the same access controls as the workbook.

## Verification

`pnpm run test:ooxml` checks combined native features and rewritten files with the
Microsoft Open XML SDK. The regression corpus includes a combined-feature fixture.
`node tools/fidelity/run.ts` compares LibreOffice-rendered original/rewrite pages
and records its re-save losses separately. The Excel COM oracle can check the same
fixture on Windows: `node tools/excel-oracle/run.ts --xlsx <file> --cells A2,C4`.
Its repair/readback results are evidence only when run on a real Excel host.

The 4.x model adds `kind`, `charts`, `shapes`, `sparklines` and `checkboxes` facets.
Construct models through `Worksheet.model` and apply changes by spreading that
snapshot, so new facets survive application-level copies.
