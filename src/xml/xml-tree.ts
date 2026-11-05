import {XmlParseError} from './errors.ts';
import {parseXml} from './xml-read.ts';
import {localName, type XmlAttributes} from './xml-scan.ts';
import {escapeAttr, escapeText} from './xml.ts';

/** A bounded XML fragment tree for editing small metadata parts, never worksheet cell grids. */
export interface XmlNode {
  name: string;
  attrs: Record<string, string>;
  children: (XmlNode | string)[];
}

export function xmlTree(xml: string, maxNodes = 100000): XmlNode {
  let root: XmlNode | undefined;
  const stack: XmlNode[] = [];
  let count = 0;
  parseXml(xml, {
    onOpen(name: string, attrs: XmlAttributes, closed: boolean) {
      if (++count > maxNodes) throw new XmlParseError('XML fragment exceeds the node limit');
      const node: XmlNode = {name, attrs: {...attrs}, children: []};
      const parent = stack.at(-1);
      if (parent) parent.children.push(node);
      else root = node;
      if (!closed) stack.push(node);
    },
    onText(text) {
      stack.at(-1)?.children.push(text);
    },
    onClose() {
      stack.pop();
    },
  });
  if (!root) throw new XmlParseError('XML fragment has no root');
  return root;
}

export function xmlChildren(node: XmlNode, name?: string): XmlNode[] {
  return node.children.filter(
    (child): child is XmlNode =>
      typeof child !== 'string' && (name === undefined || localName(child.name) === name),
  );
}

export function xmlChild(node: XmlNode | undefined, ...path: string[]): XmlNode | undefined {
  let current = node;
  for (const name of path) current = current && xmlChildren(current, name)[0];
  return current;
}

export function xmlText(node: XmlNode | undefined): string {
  return (
    node?.children.map((child) => (typeof child === 'string' ? child : xmlText(child))).join('') ??
    ''
  );
}

export function xmlSerialize(node: XmlNode): string {
  const attrs = Object.entries(node.attrs)
    .map(([key, value]) => ` ${key}="${escapeAttr(value)}"`)
    .join('');
  return node.children.length === 0
    ? `<${node.name}${attrs}/>`
    : `<${node.name}${attrs}>${node.children.map((child) => (typeof child === 'string' ? escapeText(child) : xmlSerialize(child))).join('')}</${node.name}>`;
}
