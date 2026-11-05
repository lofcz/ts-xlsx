# Chart

<!-- Generated from the public types by `pnpm run docs`. Do not edit by hand. -->

### `ChartAxis`

<sub>interface</sub>

```ts
interface ChartAxis {
  readonly title?: string;
  readonly min?: number;
  readonly max?: number;
  readonly majorUnit?: number;
  readonly formatCode?: string;
}
```

---

### `ChartData`

<sub>interface</sub>

Values are the chart's cached data; a formula optionally links them to worksheet cells.

```ts
interface ChartData<T extends string | number> {
  readonly values: readonly (T | null)[];
  readonly formula?: string;
  readonly formatCode?: string;
}
```

---

### `ChartOptions`

<sub>interface</sub>

```ts
interface ChartOptions {
  readonly type: ChartType;
  readonly title?: string;
  readonly series: readonly ChartSeries[];
  readonly categories?: ChartData<string>;
  readonly grouping?: 'clustered' | 'stacked' | 'percentStacked';
  readonly legend?: false | 'top' | 'bottom' | 'left' | 'right';
  readonly categoryAxis?: ChartAxis;
  readonly valueAxis?: ChartAxis;
  readonly secondaryValueAxis?: ChartAxis;
  readonly showValues?: boolean;
  readonly varyColors?: boolean;
  readonly holeSize?: number;
}
```

---

### `ChartPart`

<sub>interface</sub>

A complete standard DrawingML chart for producers with a richer chart model.
The writer owns package paths, drawing anchors, relationship ids and content types.
XML is validated for well-formedness; its schema conformance remains the author's responsibility.

```ts
interface ChartPart {
  readonly xml: string;
  readonly name?: string;
  readonly relatedParts?: readonly ChartRelatedPart[];
}
```

---

### `ChartRelatedPart`

<sub>interface</sub>

A related XML part, such as a chart's user-shapes drawing. Paths are assigned by the writer.

```ts
interface ChartRelatedPart {
  readonly relationshipId: string;
  readonly relationshipType: string;
  readonly contentType: string;
  readonly xml: string;
}
```

---

### `ChartSeries`

<sub>interface</sub>

```ts
interface ChartSeries {
  readonly name: string;
  readonly nameFormula?: string;
  readonly values: ChartData<number>;
  readonly categories?: ChartData<string>;
  readonly xValues?: ChartData<number>;
  readonly bubbleSizes?: ChartData<number>;
  /** Six hexadecimal RGB digits, with an optional leading #. */
  readonly color?: string;
  /** Per-series type enables column/line/area combinations. */
  readonly type?: 'column' | 'line' | 'area';
  readonly secondaryAxis?: boolean;
}
```

---

### `ChartType`

<sub>type</sub>

Chart families authored as native DrawingML, editable by spreadsheet applications.

```ts
type ChartType =
  | 'column'
  | 'bar'
  | 'line'
  | 'area'
  | 'pie'
  | 'doughnut'
  | 'scatter'
  | 'bubble'
  | 'radar';
```

---

### `isChartType`

<sub>const</sub>

```ts
const isChartType: (value: string) => value is "area" | "bar" | "bubble" | "column" | "doughnut" | "line" | "pie" | "radar" | "scatter"
```

---

### `WorksheetChart`

<sub>type</sub>

```ts
type WorksheetChart =
  | {readonly kind: 'chart'; readonly options: ChartOptions; readonly anchor: ImageAnchor}
  | {
      readonly kind: 'xml';
      readonly part: ChartPart;
      readonly anchor: ImageAnchor;
      readonly splices?: readonly SheetSplice[];
    };
```
