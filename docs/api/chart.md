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
  readonly type?: 'category' | 'date' | 'value';
  readonly logBase?: number;
  readonly reversed?: boolean;
  readonly minorUnit?: number;
  readonly baseTimeUnit?: 'days' | 'months' | 'years';
  readonly majorTimeUnit?: 'days' | 'months' | 'years';
  readonly minorTimeUnit?: 'days' | 'months' | 'years';
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

### `ChartDataLabels`

<sub>interface</sub>

```ts
interface ChartDataLabels {
  readonly showValue?: boolean;
  readonly showCategory?: boolean;
  readonly showSeriesName?: boolean;
  readonly showPercent?: boolean;
  readonly showLegendKey?: boolean;
  readonly showBubbleSize?: boolean;
  readonly position?: 'bestFit' | 'b' | 'ctr' | 'inBase' | 'inEnd' | 'l' | 'outEnd' | 'r' | 't';
  readonly formatCode?: string;
  readonly separator?: string;
  readonly lastPointOnly?: boolean;
}
```

---

### `ChartErrorBars`

<sub>interface</sub>

```ts
interface ChartErrorBars {
  readonly type: 'fixedVal' | 'percentage' | 'stdDev' | 'stdErr' | 'cust';
  readonly direction?: 'x' | 'y';
  readonly sides?: 'both' | 'plus' | 'minus';
  readonly value?: number;
  readonly plus?: ChartData<number>;
  readonly minus?: ChartData<number>;
  readonly endCap?: boolean;
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
  readonly categories?: ChartData<string | number>;
  readonly grouping?: 'clustered' | 'stacked' | 'percentStacked';
  readonly legend?: false | 'top' | 'bottom' | 'left' | 'right';
  readonly categoryAxis?: ChartAxis;
  readonly valueAxis?: ChartAxis;
  readonly secondaryValueAxis?: ChartAxis;
  readonly showValues?: boolean;
  readonly varyColors?: boolean;
  readonly holeSize?: number;
  readonly dataLabels?: ChartDataLabels;
  /** Zero-based waterfall points rendered as totals. */
  readonly waterfallTotals?: readonly number[];
  readonly waterfallConnectors?: boolean;
  readonly binning?: {
    readonly width?: number;
    readonly count?: number;
    readonly underflow?: number;
    readonly overflow?: number;
  };
  readonly boxWhisker?: {
    readonly quartileMethod?: 'inclusive' | 'exclusive';
    readonly showMeanLine?: boolean;
    readonly showMeanMarker?: boolean;
    readonly showInnerPoints?: boolean;
    readonly showOutlierPoints?: boolean;
  };
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
  readonly categories?: ChartData<string | number>;
  /** Outer-to-inner category levels for hierarchy charts. */
  readonly categoryLevels?: readonly ChartData<string>[];
  readonly xValues?: ChartData<number>;
  readonly bubbleSizes?: ChartData<number>;
  /** Six hexadecimal RGB digits, with an optional leading #. */
  readonly color?: string;
  /** Per-series type enables column/line/area combinations. */
  readonly type?: 'column' | 'line' | 'area';
  readonly secondaryAxis?: boolean;
  readonly trendlines?: readonly ChartTrendline[];
  readonly errorBars?: readonly ChartErrorBars[];
  readonly dataLabels?: ChartDataLabels;
  readonly pointColors?: readonly (string | null)[];
  readonly smooth?: boolean;
}
```

---

### `ChartSource`

<sub>interface</sub>

Original drawing identity and markup retained while an imported chart is edited.

```ts
interface ChartSource {
  readonly relationshipId?: string;
  readonly name?: string;
  readonly editedXml?: string;
  readonly xml: string;
  readonly options: ChartOptions;
  readonly anchor: ImageAnchor;
}
```

---

### `ChartTrendline`

<sub>interface</sub>

```ts
interface ChartTrendline {
  readonly type: 'linear' | 'exp' | 'log' | 'poly' | 'power' | 'movingAvg';
  readonly name?: string;
  readonly order?: number;
  readonly period?: number;
  readonly forward?: number;
  readonly backward?: number;
  readonly intercept?: number;
  readonly displayEquation?: boolean;
  readonly displayRSquared?: boolean;
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
  | 'radar'
  | 'stock'
  | 'waterfall'
  | 'histogram'
  | 'pareto'
  | 'funnel'
  | 'treemap'
  | 'sunburst'
  | 'boxWhisker';
```

---

### `isChartType`

<sub>const</sub>

```ts
const isChartType: (value: string) => value is "area" | "bar" | "boxWhisker" | "bubble" | "column" | "doughnut" | "funnel" | "histogram" | "line" | "pareto" | "pie" | "radar" | "scatter" | "stock" | "sunburst" | "treemap" | "waterfall"
```

---

### `isExtendedChart`

<sub>function</sub>

Office 2016 charts use the chartEx vocabulary and relationship type.

```ts
function isExtendedChart(type: ChartType): boolean;
```

---

### `WorksheetChart`

<sub>type</sub>

```ts
type WorksheetChart =
  | {
      readonly kind: 'chart';
      readonly options: ChartOptions;
      readonly anchor: ImageAnchor;
      readonly source?: ChartSource;
    }
  | {
      readonly kind: 'xml';
      readonly part: ChartPart;
      readonly anchor: ImageAnchor;
      readonly splices?: readonly SheetSplice[];
    };
```
