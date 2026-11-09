/** Join ordered, non-overlapping edits without copying the document for each edit. */
export function applyXmlEdits(
  xml: string,
  edits: readonly {start: number; end: number; text: string}[],
): string {
  if (edits.length === 0) return xml;
  const parts: string[] = [];
  let end = 0;
  for (const edit of edits) {
    parts.push(xml.slice(end, edit.start), edit.text);
    end = edit.end;
  }
  parts.push(xml.slice(end));
  return parts.join('');
}
