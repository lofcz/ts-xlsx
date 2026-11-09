import {decodeAddress, encodeAddress} from './address.ts';
import {replaceContents} from './containers.ts';
import {type SheetSplice, spliceFormula} from './formula-references.ts';
import {isDeletedSpan, shiftIndex} from './grid-shift.ts';
import {type ImageAnchor, isOneCellAnchor} from './image.ts';

export interface CellImage {
  readonly ref: string;
  readonly imageId: number;
  readonly alt?: string;
}

/** Excel's in-cell sparkline group; addresses and formulas use A1 notation. */
export interface SparklineGroup {
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

export interface ShapeTextRun {
  readonly text: string;
  readonly font?: {
    readonly name?: string;
    readonly size?: number;
    readonly bold?: boolean;
    readonly italic?: boolean;
    readonly color?: string;
  };
}

export interface ShapeOptions {
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

export type WorksheetShape =
  | {readonly kind: 'shape'; readonly options: ShapeOptions; readonly anchor: ImageAnchor}
  | {
      readonly kind: 'xml';
      readonly xml: string;
      readonly sourceId?: string;
      readonly modified?: boolean;
      readonly splices?: readonly SheetSplice[];
    };

export interface SheetFeatureStorage {
  readonly cellImages: CellImage[];
  readonly checkboxes: Set<string>;
  readonly sparklines: SparklineGroup[];
  readonly shapes: WorksheetShape[];
}

export function spliceSheetFeatures(
  storage: SheetFeatureStorage,
  home: string,
  edit: SheetSplice,
): void {
  const local = home.toLowerCase() === edit.sheet.toLowerCase(),
    axis = edit.splice.axis;
  const move = (ref: string): string | undefined => {
    if (!local) return ref;
    const position = decodeAddress(ref);
    if (position.col === undefined || position.row === undefined) return undefined;
    const value = position[axis];
    if (value === undefined) return undefined;
    if (isDeletedSpan(value, value, edit.splice)) return undefined;
    return encodeAddress(
      axis === 'col' ? shiftIndex(position.col, edit.splice) : position.col,
      axis === 'row' ? shiftIndex(position.row, edit.splice) : position.row,
    );
  };
  const images = storage.cellImages.flatMap((image) => {
    const ref = move(image.ref);
    return ref ? [{...image, ref}] : [];
  });
  replaceContents(storage.cellImages, images);
  const checkboxes = [...storage.checkboxes].map(move).filter((ref) => ref !== undefined);
  storage.checkboxes.clear();
  for (const ref of checkboxes) storage.checkboxes.add(ref);
  const groups = storage.sparklines
    .map((group) => ({
      ...group,
      ...(group.dateAxis ? {dateAxis: spliceFormula(group.dateAxis, home, edit)} : {}),
      sparklines: group.sparklines.flatMap((item) => {
        const ref = move(item.ref);
        return ref ? [{ref, formula: spliceFormula(item.formula, home, edit)}] : [];
      }),
    }))
    .filter((group) => group.sparklines.length);
  replaceContents(storage.sparklines, groups);
  if (!local) return;
  const shapes = storage.shapes.map((shape) => {
    if (shape.kind === 'xml') return {...shape, splices: [...(shape.splices ?? []), edit]};
    const anchor = shape.anchor;
    if (!isOneCellAnchor(anchor) && anchor.editAs === 'absolute') return shape;
    const point = (value: ImageAnchor['from']) => ({
      ...value,
      [axis]: shiftIndex(value[axis] + 1, edit.splice) - 1,
    });
    return {
      ...shape,
      anchor: isOneCellAnchor(anchor)
        ? {...anchor, from: point(anchor.from)}
        : {...anchor, from: point(anchor.from), to: point(anchor.to)},
    };
  });
  replaceContents(storage.shapes, shapes);
}
