# Chart Cache

<!-- Generated from the public types by `pnpm run docs`. Do not edit by hand. -->

### `ChartCacheDiagnostic`

<sub>interface</sub>

A reference whose cache was retained because no bounded local range could be resolved.

```ts
interface ChartCacheDiagnostic {
  readonly sheet: string;
  readonly formula: string;
  readonly reason: 'unresolved-reference' | 'range-limit' | 'uncalculated-formula';
}
```

---

### `resolveChartCaches`

<sub>function</sub>

Refresh chart references using stored cell values and formula results, without evaluating formulas.

```ts
function resolveChartCaches(
  workbook: Workbook,
  home: string,
  options: ChartOptions,
  diagnostics: ChartCacheDiagnostic[] = [],
): ChartOptions;
```
