import type {WorksheetChart} from '../../core/chart.ts';
import {INTERNAL} from '../../core/internal.ts';
import {modelEqual, patchKeyed, patchModel} from '../../core/model-patch.ts';
import type {Workbook} from '../../core/workbook.ts';
import type {Worksheet} from '../../core/worksheet.ts';
import {AuthoringError} from '../../errors.ts';
import {mergeXml} from '../../xml/xml-merge.ts';
import {xmlChildren, xmlSerialize, xmlTree} from '../../xml/xml-tree.ts';
import {mergeDrawingPictures} from './merge-pictures.ts';
import {readXlsx} from './read.ts';
import {writeXlsx} from './write.ts';

export interface FidelityDiagnostic {
  readonly code:
    | 'drawing-replaced'
    | 'sheet-deleted'
    | 'sheet-renamed'
    | 'chart-cache-retained'
    | 'unmodeled-structure';
  readonly message: string;
  readonly sheet?: string;
}

/** Map stable editor sheet identities to source and baseline names; array order is the final tab order. */
export interface WorkbookEditSheet {
  readonly name: string;
  readonly originalName?: string;
  readonly baselineName?: string;
}

export interface MergeXlsxEditsOptions {
  readonly sheets: readonly WorkbookEditSheet[];
  readonly maxUncompressedBytes?: number;
}

export interface MergeXlsxEditsResult {
  readonly data: Uint8Array;
  readonly diagnostics: readonly FidelityDiagnostic[];
}

function mergeCharts(source: Worksheet, baseline: Worksheet, edited: Worksheet): WorksheetChart[] {
  const originals = source.charts,
    before = baseline.charts,
    after = edited.charts;
  if (modelEqual(before, after)) return [...originals];
  const identity = (chart: WorksheetChart) =>
    chart.kind === 'chart' ? chart.source?.name : chart.part.name;
  const available = new Set(originals);
  const matched = new Map<WorksheetChart, WorksheetChart>();
  for (const previous of before) {
    const original =
      [...available].find((item) => identity(item) === identity(previous)) ??
      [...available].find(
        (item) =>
          item.kind === 'chart' &&
          previous.kind === 'chart' &&
          item.options.title === previous.options.title &&
          item.options.type === previous.options.type,
      ) ??
      (originals.length === before.length ? originals[before.indexOf(previous)] : undefined);
    if (original && available.delete(original)) matched.set(previous, original);
  }
  const used = new Set<WorksheetChart>();
  const result = after.map((next) => {
    const previous = before.find((item) => !used.has(item) && identity(item) === identity(next));
    if (previous) used.add(previous);
    const original = previous && matched.get(previous);
    if (original?.kind === 'chart' && previous?.kind === 'chart' && next.kind === 'chart') {
      const editedXml =
        original.source && previous.source && next.source
          ? xmlSerialize(
              mergeXml(
                xmlTree(original.source.xml),
                xmlTree(previous.source.xml),
                xmlTree(next.source.xml),
              ),
            )
          : undefined;
      return {
        ...original,
        ...(original.source && editedXml ? {source: {...original.source, editedXml}} : {}),
        options: patchModel(original.options, previous.options, next.options) ?? next.options,
        anchor: patchModel(original.anchor, previous.anchor, next.anchor) ?? next.anchor,
      };
    }
    return next.kind === 'chart'
      ? {kind: 'chart' as const, options: next.options, anchor: next.anchor}
      : next;
  });
  // Objects the editor never exposed are still owned by the source workbook.
  result.push(...available);
  return result;
}

function cellImages(workbook: Workbook, sheet: Worksheet): unknown {
  return sheet.cellImages.map((image) => ({...image, image: workbook.getImage(image.imageId)}));
}

function mergeSheet(
  source: Worksheet,
  baseline: Worksheet,
  edited: Worksheet,
  original: Workbook,
  before: Workbook,
  after: Workbook,
  diagnostics: FidelityDiagnostic[],
): void {
  if (
    (baseline.rowCount !== edited.rowCount || baseline.columnCount !== edited.columnCount) &&
    (source.preservedReferences.length ||
      source[INTERNAL].preservedXml.some((xml) => /(?:sqref|ref)=/.test(xml)))
  )
    diagnostics.push({
      code: 'unmodeled-structure',
      sheet: source.name,
      message:
        'The used grid changed; references inside preserved objects outside the editor model may need review.',
    });
  const current = source.model,
    previous = baseline.model,
    next = edited.model;
  const cells = patchKeyed(
    current.cells,
    previous.cells,
    next.cells,
    (cell) => `${cell.row}:${cell.col}`,
  );
  const rows = patchKeyed(current.rows, previous.rows, next.rows, (row) => String(row.number));
  const columns = patchKeyed(current.columns, previous.columns, next.columns, (col) =>
    String(col.index),
  );
  const merged = patchModel(current, previous, next) ?? next;
  const charts = mergeCharts(source, baseline, edited);
  const shapeNodes = (xml: string) => {
    const nodes: ReturnType<typeof xmlTree>[] = [];
    const walk = (node: ReturnType<typeof xmlTree>) => {
      if (node.name.split(':').at(-1) === 'cNvPr') nodes.push(node);
      for (const child of xmlChildren(node)) walk(child);
    };
    const root = xmlTree(xml);
    walk(root);
    return {root, nodes};
  };
  const shapeKey = (shape: (typeof merged.shapes)[number]) =>
    shape.kind === 'xml' ? shapeNodes(shape.xml).nodes[0]?.attrs.name : shape.options.name;
  const unseen = new Set(current.shapes);
  const originalByPrevious = new Map(
    previous.shapes.map((shape, index) => {
      const originalShape =
        [...unseen].find((item) => shapeKey(item) === shapeKey(shape)) ??
        (current.shapes.length === previous.shapes.length ? current.shapes[index] : undefined);
      if (originalShape) unseen.delete(originalShape);
      return [shapeKey(shape), originalShape];
    }),
  );
  const shapes = next.shapes.map((shape) => {
    const old = previous.shapes.find((item) => shapeKey(item) === shapeKey(shape)),
      originalShape = originalByPrevious.get(shapeKey(shape));
    if (shape.kind !== 'xml' || old?.kind !== 'xml' || originalShape?.kind !== 'xml')
      return shape.kind === 'xml' ? {kind: 'xml' as const, xml: shape.xml} : shape;
    const originalTree = shapeNodes(originalShape.xml),
      oldTree = shapeNodes(old.xml),
      nextTree = shapeNodes(shape.xml);
    for (const [index, node] of originalTree.nodes.entries()) {
      const a = oldTree.nodes[index],
        b = nextTree.nodes[index];
      if (a && node.attrs.id) a.attrs.id = node.attrs.id;
      if (b && node.attrs.id) b.attrs.id = node.attrs.id;
    }
    return {
      ...originalShape,
      xml: xmlSerialize(mergeXml(originalTree.root, oldTree.root, nextTree.root)),
      modified: !modelEqual(old, shape),
    };
  });
  shapes.push(...unseen);
  const remapStyles = <T>(value: T): T => {
    if (Array.isArray(value)) return value.map(remapStyles) as T;
    if (!value || typeof value !== 'object') return value;
    return Object.fromEntries(
      Object.entries(value).map(([key, item]) => {
        const fragment =
          key.toLowerCase().endsWith('dxfid') &&
          (typeof item === 'number' || typeof item === 'string')
            ? after.differentialStyles[Number(item)]
            : undefined;
        return [
          key,
          fragment === undefined
            ? remapStyles(item)
            : typeof item === 'string'
              ? String(original[INTERNAL].adoptDifferentialStyle(fragment))
              : original[INTERNAL].adoptDifferentialStyle(fragment),
        ];
      }),
    ) as T;
  };
  const conditionalFormattings = modelEqual(
    previous.conditionalFormattings,
    next.conditionalFormattings,
  )
    ? current.conditionalFormattings
    : remapStyles(next.conditionalFormattings);
  const tables = modelEqual(previous.tables, next.tables)
    ? current.tables
    : remapStyles(next.tables);
  source.model = {...merged, cells, rows, columns, charts, shapes, conditionalFormattings, tables};
  if (!modelEqual(baseline.commentThreads, edited.commentThreads))
    source[INTERNAL].restoreCommentThreads(edited.commentThreads);
  if (!modelEqual(cellImages(before, baseline), cellImages(after, edited))) {
    source[INTERNAL].cellImages.splice(0);
    for (const image of edited.cellImages) {
      const media = after.getImage(image.imageId);
      if (media)
        source.addCellImage(
          original.addImage({buffer: media.data, extension: media.extension}),
          image.ref,
          image.alt,
        );
    }
  }
  const beforeBackground = before.exportImages(baseline).background,
    afterBackground = after.exportImages(edited).background;
  if (!modelEqual(beforeBackground, afterBackground)) {
    source.removeBackgroundImage();
    if (afterBackground)
      source.addBackgroundImage(
        original.addImage({buffer: afterBackground.data, extension: afterBackground.extension}),
      );
  }
  if (mergeDrawingPictures(source, baseline, edited, original, before, after))
    diagnostics.push({
      code: 'drawing-replaced',
      sheet: source.name,
      message:
        'Picture edits replaced picture anchors the editor did not fully expose; charts and shapes were retained.',
    });
}

/**
 * Merge an editor's changes into its original workbook. The baseline is the editor's first export,
 * so a feature it never understood is not mistaken for a user deletion. Inputs are independently
 * bounded ZIP packages; the result includes actionable loss diagnostics.
 */
export function mergeXlsxEdits(
  originalBytes: Uint8Array,
  baselineBytes: Uint8Array,
  editedBytes: Uint8Array,
  options: MergeXlsxEditsOptions,
): MergeXlsxEditsResult {
  const readOptions =
    options.maxUncompressedBytes === undefined
      ? {}
      : {maxUncompressedBytes: options.maxUncompressedBytes};
  const original = readXlsx(originalBytes, readOptions),
    before = readXlsx(baselineBytes, readOptions),
    after = readXlsx(editedBytes, readOptions);
  const diagnostics: FidelityDiagnostic[] = [],
    ordered: Worksheet[] = [],
    names = new Set<string>();
  const sources = new Map(original.worksheets.map((sheet) => [sheet.name.toLowerCase(), sheet]));
  for (const person of after.persons)
    if (!original.getPerson(person.id)) original.addPerson(person);
  for (const mapping of options.sheets) {
    if (names.has(mapping.name.toLowerCase()))
      throw new AuthoringError('edited sheet names must be unique');
    names.add(mapping.name.toLowerCase());
    const edited = after.requireWorksheet(mapping.name),
      baseline = mapping.baselineName ? before.getWorksheet(mapping.baselineName) : undefined;
    let source = mapping.originalName ? sources.get(mapping.originalName.toLowerCase()) : undefined;
    if (source && ordered.includes(source))
      throw new AuthoringError('an original sheet cannot be mapped twice');
    if (source && baseline)
      mergeSheet(source, baseline, edited, original, before, after, diagnostics);
    else {
      let temporary = `Imported ${original.worksheets.length + 1}`;
      while (original.getWorksheet(temporary)) temporary += 'x';
      source = original.addWorksheet(temporary);
      source.model = edited.model;
      original.importImages(source, after.exportImages(edited));
      source[INTERNAL].replacePreservedReferences(edited.preservedReferences);
      source[INTERNAL].restoreCharts(edited.charts);
      source[INTERNAL].restoreCommentThreads(edited.commentThreads);
      for (const image of edited.cellImages) {
        const media = after.getImage(image.imageId);
        if (media)
          source.addCellImage(
            original.addImage({buffer: media.data, extension: media.extension}),
            image.ref,
            image.alt,
          );
      }
    }
    if (mapping.originalName && mapping.originalName !== mapping.name)
      diagnostics.push({
        code: 'sheet-renamed',
        sheet: mapping.name,
        message: `Renamed ${mapping.originalName}; references in opaque embedded objects may still use the old name.`,
      });
    source[INTERNAL].rename(mapping.name);
    ordered.push(source);
  }
  for (const sheet of original.worksheets)
    if (!ordered.includes(sheet))
      diagnostics.push({
        code: 'sheet-deleted',
        sheet: sheet.name,
        message: 'The sheet was deleted; embedded objects that reference it may need updating.',
      });
  if (ordered.length === 0) throw new AuthoringError('an edited workbook needs at least one sheet');
  original[INTERNAL].restoreWorksheets(ordered);
  const fragmentKey = (xml: string) => xmlTree(xml).name.split(':').at(-1) ?? '';
  const fragments = new Map(original[INTERNAL].preservedXml.map((xml) => [fragmentKey(xml), xml]));
  const previousFragments = new Map(
    before[INTERNAL].preservedXml.map((xml) => [fragmentKey(xml), xml]),
  );
  const nextFragments = new Map(after[INTERNAL].preservedXml.map((xml) => [fragmentKey(xml), xml]));
  for (const key of new Set([...previousFragments.keys(), ...nextFragments.keys()])) {
    const previous = previousFragments.get(key),
      next = nextFragments.get(key),
      source = fragments.get(key);
    if (previous === next) continue;
    if (!next) fragments.delete(key);
    else
      fragments.set(
        key,
        previous && source
          ? xmlSerialize(mergeXml(xmlTree(source), xmlTree(previous), xmlTree(next)))
          : next,
      );
  }
  original[INTERNAL].preservedXml.splice(0, Infinity, ...fragments.values());
  original.protection = patchModel(original.protection, before.protection, after.protection);
  original.fullCalcOnLoad =
    patchModel(original.fullCalcOnLoad, before.fullCalcOnLoad, after.fullCalcOnLoad) ?? false;
  Object.assign(original.view, patchModel(original.view, before.view, after.view));
  original[INTERNAL].restoreDefinedNames(
    patchKeyed(
      original.definedNames,
      before.definedNames,
      after.definedNames,
      (name) => `${name.scope ?? ''}:${name.name}`,
    ),
  );
  const cacheDiagnostics: import('../../core/chart-cache.ts').ChartCacheDiagnostic[] = [];
  const data = writeXlsx(original, {chartCacheDiagnostics: cacheDiagnostics});
  diagnostics.push(
    ...cacheDiagnostics.map((item) => ({
      code: 'chart-cache-retained' as const,
      sheet: item.sheet,
      message: `Kept chart cache for ${item.formula}: ${item.reason}.`,
    })),
  );
  return {data, diagnostics};
}
