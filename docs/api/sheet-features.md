# Sheet Features

<!-- Generated from the public types by `pnpm run docs`. Do not edit by hand. -->

### `CellImage`

<sub>interface</sub>

```ts
interface CellImage {
  readonly ref: string;
  readonly imageId: number;
  readonly alt?: string;
}
```

---

### `ShapeOptions`

<sub>interface</sub>

```ts
interface ShapeOptions {
  readonly name?: string;
  readonly alt?: string;
  /** DrawingML preset geometry, e.g. rect, roundRect, ellipse, rightArrow. */
  readonly geometry?: string;
  readonly fill?: string;
  readonly outline?: {readonly color: string; readonly width?: number};
  readonly rotation?: number;
  readonly text?: string | readonly ShapeTextRun[];
  readonly textBox?: boolean;
}
```

---

### `ShapeTextRun`

<sub>interface</sub>

```ts
interface ShapeTextRun {
  readonly text: string;
  readonly font?: {
    readonly name?: string;
    readonly size?: number;
    readonly bold?: boolean;
    readonly italic?: boolean;
    readonly color?: string;
  };
}
```

---

### `SparklineGroup`

<sub>interface</sub>

Excel's in-cell sparkline group; addresses and formulas use A1 notation.

```ts
interface SparklineGroup {
  readonly type?: 'line' | 'column' | 'winloss';
  readonly sparklines: readonly {readonly ref: string; readonly formula: string}[];
  readonly dateAxis?: string;
  readonly lineWeight?: number;
  readonly markers?: boolean;
  readonly high?: boolean;
  readonly low?: boolean;
  readonly first?: boolean;
  readonly last?: boolean;
  readonly negative?: boolean;
  readonly displayXAxis?: boolean;
  readonly displayHidden?: boolean;
  readonly rightToLeft?: boolean;
  readonly displayEmptyCellsAs?: 'gap' | 'zero' | 'span';
  readonly minAxisType?: 'individual' | 'group' | 'custom';
  readonly maxAxisType?: 'individual' | 'group' | 'custom';
  readonly manualMin?: number;
  readonly manualMax?: number;
  /** ARGB colors, matching worksheet style colors. */
  readonly colors?: Partial<
    Readonly<
      Record<'series' | 'negative' | 'axis' | 'markers' | 'first' | 'last' | 'high' | 'low', string>
    >
  >;
}
```

---

### `WorksheetShape`

<sub>type</sub>

```ts
type WorksheetShape =
  | {readonly kind: 'shape'; readonly options: ShapeOptions; readonly anchor: ImageAnchor}
  | {
      readonly kind: 'xml';
      readonly xml: string;
      readonly sourceId?: string;
      readonly modified?: boolean;
      readonly splices?: readonly SheetSplice[];
    };
```
