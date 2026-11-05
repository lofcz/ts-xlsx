# Xlsx Merge Edits

<!-- Generated from the public types by `pnpm run docs`. Do not edit by hand. -->

### `FidelityDiagnostic`

<sub>interface</sub>

```ts
interface FidelityDiagnostic {
  readonly code:
    | 'drawing-replaced'
    | 'sheet-deleted'
    | 'sheet-renamed'
    | 'chart-cache-retained'
    | 'unmodeled-structure';
  readonly message: string;
  readonly sheet?: string;
}
```

---

### `mergeXlsxEdits`

<sub>function</sub>

Merge an editor's changes into its original workbook. The baseline is the editor's first export,
so a feature it never understood is not mistaken for a user deletion. Inputs are independently
bounded ZIP packages; the result includes actionable loss diagnostics.

```ts
function mergeXlsxEdits(
  originalBytes: Uint8Array,
  baselineBytes: Uint8Array,
  editedBytes: Uint8Array,
  options: MergeXlsxEditsOptions,
): MergeXlsxEditsResult;
```

---

### `MergeXlsxEditsOptions`

<sub>interface</sub>

```ts
interface MergeXlsxEditsOptions {
  readonly sheets: readonly WorkbookEditSheet[];
  readonly maxUncompressedBytes?: number;
}
```

---

### `MergeXlsxEditsResult`

<sub>interface</sub>

```ts
interface MergeXlsxEditsResult {
  readonly data: Uint8Array;
  readonly diagnostics: readonly FidelityDiagnostic[];
}
```

---

### `WorkbookEditSheet`

<sub>interface</sub>

Map stable editor sheet identities to source and baseline names; array order is the final tab order.

```ts
interface WorkbookEditSheet {
  readonly name: string;
  readonly originalName?: string;
  readonly baselineName?: string;
}
```
