import {INTERNAL} from '../../core/internal.ts';
import {modelEqual} from '../../core/model-patch.ts';
import type {PreservedPart, PreservedWorksheetReference} from '../../core/preserved.ts';
import {Workbook} from '../../core/workbook.ts';
import type {Worksheet} from '../../core/worksheet.ts';
import {xmlChild, xmlChildren, xmlSerialize, xmlTree, type XmlNode} from '../../xml/xml-tree.ts';
import {
  capturePartClosure,
  contentTypeResolver,
  openSpreadsheetPackage,
  readPartRelationships,
} from '../opc/read-opc.ts';
import {writeXlsx} from './write.ts';

const encoder = new TextEncoder(),
  decoder = new TextDecoder();

function drawingOf(workbook: Workbook, sheet: Worksheet): PreservedWorksheetReference | undefined {
  const preserved = sheet.preservedReferences.find((ref) => ref.element === 'drawing');
  if (preserved || sheet.images.length === 0) return preserved;
  const temporary = new Workbook(),
    target = temporary.addWorksheet('Images');
  temporary.importImages(target, workbook.exportImages(sheet));
  const {pkg} = openSpreadsheetPackage(writeXlsx(temporary), undefined);
  const entryPath = readPartRelationships('xl/worksheets/sheet1.xml', pkg.partText).targetPath(
    'drawing',
  );
  if (!entryPath) return undefined;
  const parts = capturePartClosure(
    entryPath,
    pkg.partText,
    pkg.partBytes,
    contentTypeResolver(pkg.partText('[Content_Types].xml') ?? ''),
  );
  return parts
    ? {
        element: 'drawing',
        relType: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing',
        entryPath,
        parts,
      }
    : undefined;
}

function pictureState(ref: PreservedWorksheetReference | undefined) {
  const entry = ref?.parts.find((part) => part.path === ref.entryPath);
  const tree = entry ? xmlTree(decoder.decode(entry.bytes)) : undefined;
  const pictures = tree ? xmlChildren(tree).filter((node) => xmlChild(node, 'pic')) : [];
  const normalize = (node: XmlNode): unknown => ({
    name: node.name.split(':').at(-1),
    attrs: Object.fromEntries(
      Object.entries(node.attrs)
        .filter(([key]) => !key.startsWith('xmlns'))
        .map(([key, value]) => {
          const rel = entry?.rels.find((item) => item.id === value && key.startsWith('r:'));
          return [
            key,
            rel
              ? rel.external
                ? rel.targetPath
                : ref?.parts.find((part) => part.path === rel.targetPath)?.bytes
              : value,
          ];
        }),
    ),
    children: node.children.map((child) => (typeof child === 'string' ? child : normalize(child))),
  });
  return {entry, tree, pictures, comparable: pictures.map(normalize)};
}

/** Picture edits replace only picture anchors; the original chart/shape graph keeps its ownership. */
export function mergeDrawingPictures(
  source: Worksheet,
  baseline: Worksheet,
  edited: Worksheet,
  original: Workbook,
  before: Workbook,
  after: Workbook,
): boolean {
  const initial = drawingOf(original, source),
    previous = drawingOf(before, baseline),
    current = drawingOf(after, edited);
  const a = pictureState(initial),
    b = pictureState(previous),
    c = pictureState(current);
  if (modelEqual(b.comparable, c.comparable)) return false;
  if (!initial || !a.entry || !a.tree) {
    if (!current) return false;
    original.importImages(source, {
      anchored: [],
      background: original.exportImages(source).background,
    });
    const pictureOnly = {
      ...current,
      parts: current.parts.map((part) =>
        part.path === current.entryPath && c.tree
          ? {...part, bytes: encoder.encode(xmlSerialize({...c.tree, children: c.pictures}))}
          : part,
      ),
    };
    source[INTERNAL].replacePreservedReferences([
      ...source.preservedReferences.filter((ref) => ref.element !== 'drawing'),
      pictureOnly,
    ]);
    return false;
  }
  const parts = [...initial.parts],
    rels = [...a.entry.rels];
  const copied = new Map<string, string>();
  const copyPart = (path: string): string => {
    const known = copied.get(path);
    if (known) return known;
    const part = current?.parts.find((item) => item.path === path);
    if (!part) return path;
    let index = parts.length + copied.size + 1;
    const extension = path.split('.').at(-1) ?? 'bin';
    let target = `xl/fidelityEdits/part${index}.${extension}`;
    while (parts.some((item) => item.path === target))
      target = `xl/fidelityEdits/part${++index}.${extension}`;
    copied.set(path, target);
    const result: PreservedPart = {
      ...part,
      path: target,
      rels: part.rels.map((rel) =>
        rel.external ? rel : {...rel, targetPath: copyPart(rel.targetPath)},
      ),
    };
    parts.push(result);
    return target;
  };
  const rewrite = (node: XmlNode) => {
    for (const [key, value] of Object.entries(node.attrs)) {
      if (!key.startsWith('r:')) continue;
      const relationship = c.entry?.rels.find((rel) => rel.id === value);
      if (!relationship) continue;
      let n = rels.length + 1;
      while (rels.some((rel) => rel.id === `rId${n}`)) n++;
      const id = `rId${n}`;
      rels.push({
        ...relationship,
        id,
        targetPath: relationship.external
          ? relationship.targetPath
          : copyPart(relationship.targetPath),
      });
      node.attrs[key] = id;
    }
    for (const child of xmlChildren(node)) rewrite(child);
  };
  const replacements = c.pictures.map((node) => {
    const clone = structuredClone(node);
    Object.assign(clone.attrs, c.tree?.attrs);
    rewrite(clone);
    return clone;
  });
  a.tree.children = a.tree.children.filter(
    (node) => typeof node === 'string' || !xmlChild(node, 'pic'),
  );
  a.tree.children.push(...replacements);
  let shapeId = 0;
  const ids = (node: XmlNode, assign = false) => {
    if (node.name.split(':').at(-1) === 'cNvPr') {
      if (assign) node.attrs.id = String(++shapeId);
      else shapeId = Math.max(shapeId, Number(node.attrs.id) || 0);
    }
    for (const child of xmlChildren(node)) ids(child, assign);
  };
  ids(a.tree);
  for (const node of replacements) ids(node, true);
  const usedRelationships = new Set<string>();
  const collect = (node: XmlNode) => {
    for (const [key, value] of Object.entries(node.attrs))
      if (key.startsWith('r:')) usedRelationships.add(value);
    for (const child of xmlChildren(node)) collect(child);
  };
  collect(a.tree);
  const liveRels = rels.filter(
    (rel) => !rel.type.endsWith('/image') || usedRelationships.has(rel.id),
  );
  const drawingBytes = encoder.encode(xmlSerialize(a.tree));
  const replacement = {
    ...initial,
    parts: parts.map((part) =>
      part.path === a.entry?.path ? {...part, bytes: drawingBytes, rels: liveRels} : part,
    ),
  };
  original.importImages(source, {
    anchored: [],
    background: original.exportImages(source).background,
  });
  source[INTERNAL].replacePreservedReferences([
    ...source.preservedReferences.filter((ref) => ref.element !== 'drawing'),
    replacement,
  ]);
  return a.pictures.length !== b.pictures.length;
}
