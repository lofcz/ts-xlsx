import {AuthoringError} from '../../errors.ts';
import {xmlChildren, xmlSerialize, xmlTree, type XmlNode} from '../../xml/xml-tree.ts';
import {XML_DECLARATION} from '../../xml/xml.ts';
import {relativePartPath, relsPathFor} from '../opc/part-paths.ts';
import {parseRelationshipRecords} from '../opc/relationship-records.ts';
import {relsPartXml} from '../opc/rels.ts';

/** One collision-aware owner for late-added package parts and relationships. */
export class PackageEditor {
  readonly #parts: Record<string, Uint8Array>;
  constructor(parts: Record<string, Uint8Array>) {
    this.#parts = parts;
  }
  read(path: string): string {
    return new TextDecoder().decode(this.#parts[path]);
  }
  write(path: string, xml: string): void {
    this.#parts[path] = new TextEncoder().encode(xml);
  }
  tree(path: string): XmlNode {
    return xmlTree(this.read(path));
  }
  writeTree(path: string, tree: XmlNode): void {
    this.write(path, XML_DECLARATION + xmlSerialize(tree));
  }
  override(path: string, contentType: string): void {
    const types = this.tree('[Content_Types].xml');
    let item = xmlChildren(types, 'Override').find((node) => node.attrs.PartName === `/${path}`);
    if (!item) {
      item = {name: 'Override', attrs: {PartName: `/${path}`}, children: []};
      types.children.push(item);
    }
    item.attrs.ContentType = contentType;
    this.writeTree('[Content_Types].xml', types);
  }
  relationship(owner: string, type: string, target: string): string {
    const path = relsPathFor(owner),
      rels = parseRelationshipRecords(this.read(path));
    const relative = relativePartPath(owner, target);
    const existing = rels.find(
      (rel) => rel.type === type && rel.target === relative && !rel.external,
    );
    if (existing) return existing.id;
    let id = 1;
    const used = new Set(rels.map((rel) => rel.id));
    while (used.has(`rId${id}`)) id++;
    rels.push({id: `rId${id}`, type, target: relative, external: false});
    this.write(path, relsPartXml(rels));
    return `rId${id}`;
  }
  require(path: string): void {
    if (!Object.hasOwn(this.#parts, path))
      throw new AuthoringError(`required package part is missing: ${path}`);
  }
}
