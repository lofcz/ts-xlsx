# Formula References

<!-- Generated from the public types by `pnpm run docs`. Do not edit by hand. -->

### `SheetSplice`

<sub>interface</sub>

A splice as a formula sees it: which sheet's lines moved, and how.

```ts
interface SheetSplice {
  /** The spliced sheet's name, matched in any case. */
  readonly sheet: string;
  readonly splice: AxisSplice;
}
```
