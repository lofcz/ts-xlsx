import type {
  ReadPackageOptions,
  readXlsxSource,
  indexSharedStrings,
  WorksheetSource,
  worksheetSourceRow,
  XlsxSource,
} from '../index.ts';
import type {Equal, Expect} from './expect.ts';

export type SourceContracts = [
  Expect<
    Equal<
      Parameters<typeof readXlsxSource>,
      [data: Uint8Array, options?: ReadPackageOptions | undefined]
    >
  >,
  Expect<Equal<ReturnType<typeof readXlsxSource>, XlsxSource>>,
  Expect<Equal<WorksheetSource['rows'], Uint32Array>>,
  Expect<Equal<ReturnType<typeof indexSharedStrings>, Uint32Array>>,
  Expect<Equal<ReturnType<typeof worksheetSourceRow>, string | undefined>>,
];
