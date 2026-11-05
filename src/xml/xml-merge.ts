import {localName} from './xml-scan.ts';
import {xmlChildren, xmlSerialize, type XmlNode} from './xml-tree.ts';

/** Apply only the authored differences, retaining markup outside the author's model. */
export function mergeXml(original: XmlNode, before: XmlNode, after: XmlNode): XmlNode {
  if (xmlSerialize(before) === xmlSerialize(after)) return original;
  const result = structuredClone(original);
  for (const key of new Set([...Object.keys(before.attrs), ...Object.keys(after.attrs)])) {
    if (before.attrs[key] === after.attrs[key]) continue;
    const value = after.attrs[key];
    if (value === undefined) delete result.attrs[key];
    else result.attrs[key] = value;
  }
  const oldChildren = xmlChildren(before),
    newChildren = xmlChildren(after);
  if (oldChildren.length === 0 && newChildren.length === 0) {
    result.children = structuredClone(after.children);
    return result;
  }
  const names = new Set([...oldChildren, ...newChildren].map((child) => localName(child.name)));
  for (const name of names) {
    const oldGroup = xmlChildren(before, name),
      newGroup = xmlChildren(after, name);
    const originals = xmlChildren(result, name);
    for (let i = 0; i < Math.max(oldGroup.length, newGroup.length); i++) {
      const oldNode = oldGroup[i],
        newNode = newGroup[i],
        originalNode = originals[i];
      if (oldNode && newNode && xmlSerialize(oldNode) === xmlSerialize(newNode)) continue;
      const index = originalNode ? result.children.indexOf(originalNode) : -1;
      if (!newNode) {
        if (index >= 0) result.children.splice(index, 1);
      } else if (index >= 0 && originalNode) {
        result.children[index] = oldNode
          ? mergeXml(originalNode, oldNode, newNode)
          : structuredClone(newNode);
      } else {
        // Insert before the next modeled sibling so extension lists stay last.
        const following = newChildren
          .slice(newChildren.indexOf(newNode) + 1)
          .map((child) => localName(child.name));
        const next = result.children.findIndex(
          (child) => typeof child !== 'string' && following.includes(localName(child.name)),
        );
        const ext = result.children.findIndex(
          (child) => typeof child !== 'string' && localName(child.name) === 'extLst',
        );
        result.children.splice(
          next >= 0 ? next : ext >= 0 ? ext : result.children.length,
          0,
          structuredClone(newNode),
        );
      }
    }
  }
  return result;
}
