import type {ChartOptions} from '../core/chart.ts';
import type {
  CellImage,
  ShapeOptions,
  SparklineGroup,
  WorksheetShape,
} from '../core/sheet-features.ts';
import type {Workbook} from '../core/workbook.ts';
import type {PixelAnchor} from '../core/worksheet-pictures.ts';
import type {Worksheet} from '../core/worksheet.ts';
import type {
  FidelityDiagnostic,
  mergeXlsxEdits,
  MergeXlsxEditsOptions,
  MergeXlsxEditsResult,
} from '../io/xlsx/merge-edits.ts';
import type {Equal, Expect} from './expect.ts';

export type FidelityContracts = [
  Expect<
    Equal<Parameters<Worksheet['updateChart']>, [index: number, patch: Partial<ChartOptions>]>
  >,
  Expect<Equal<Worksheet['cellImages'], readonly CellImage[]>>,
  Expect<Equal<Worksheet['sparklineGroups'], readonly SparklineGroup[]>>,
  Expect<Equal<Worksheet['shapes'], readonly WorksheetShape[]>>,
  Expect<Equal<Parameters<Worksheet['addShape']>, [options: ShapeOptions, anchor: PixelAnchor]>>,
  Expect<Equal<ReturnType<Workbook['addChartSheet']>, Worksheet>>,
  Expect<
    Equal<
      Parameters<typeof mergeXlsxEdits>,
      [
        originalBytes: Uint8Array,
        baselineBytes: Uint8Array,
        editedBytes: Uint8Array,
        options: MergeXlsxEditsOptions,
      ]
    >
  >,
  Expect<Equal<ReturnType<typeof mergeXlsxEdits>, MergeXlsxEditsResult>>,
  Expect<Equal<MergeXlsxEditsResult['diagnostics'], readonly FidelityDiagnostic[]>>,
];
