# Xlsx Read Source

<!-- Generated from the public types by `pnpm run docs`. Do not edit by hand. -->

### `indexWorksheetSource`

<sub>function</sub>

Index source locations without parsing strings, numbers or style records.

```ts
function indexWorksheetSource(xml: string): WorksheetSource;
```

---

### `readXlsxSource`

<sub>function</sub>

Open the package once and index worksheet rows for on-demand cell decoding.

```ts
function readXlsxSource(data: Uint8Array, options: ReadPackageOptions = {}): XlsxSource;
```

---

### `WorksheetSource`

<sub>interface</sub>

An immutable worksheet source suitable for structured cloning to another thread.
Values and styles remain encoded until the consumer requests a row.

```ts
interface WorksheetSource {
  readonly xml: string;
  /** Worksheet XML with ordinary cells omitted. Formulas and row properties remain. */
  readonly metadataXml: string;
  /** Triples of 1-based row number, start offset, end offset into {@link xml}. */
  readonly rows: Uint32Array;
  /** Triples of 1-based column, start offset and end offset for each cell. */
  readonly cells: Uint32Array;
  /** Cell indexes at the beginning of each indexed row, with one final end index. */
  readonly rowCellStarts: Uint32Array;
  readonly rowCount: number;
  readonly columnCount: number;
  readonly cellCount: number;
}
```

---

### `worksheetSourceCell`

<sub>function</sub>

Read a raw cell by 1-based row and column without decoding its neighbours.

```ts
function worksheetSourceCell(
  source: WorksheetSource,
  row: number,
  column: number,
): string | undefined;
```

---

### `worksheetSourceRow`

<sub>function</sub>

Read one raw row, using its 1-based worksheet number; absent rows return undefined.

```ts
function worksheetSourceRow(source: WorksheetSource, row: number): string | undefined;
```

---

### `XlsxSource`

<sub>interface</sub>

A package opened without building a mutable workbook or decoding cell values.

```ts
interface XlsxSource {
  readonly files: Readonly<Record<string, Uint8Array>>;
  readonly worksheets: readonly (WorksheetSource & {
    readonly name: string;
    readonly path: string;
  })[];
}
```
