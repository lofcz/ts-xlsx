/** Semantic equality for detached workbook snapshots, including dates and binary media. */
export function modelEqual(a: unknown, b: unknown): boolean {
  if (Object.is(a, b)) return true;
  if (a instanceof Date && b instanceof Date) return a.getTime() === b.getTime();
  if (a instanceof Uint8Array && b instanceof Uint8Array)
    return a.length === b.length && a.every((value, index) => value === b[index]);
  if (
    !a ||
    !b ||
    typeof a !== 'object' ||
    typeof b !== 'object' ||
    Array.isArray(a) !== Array.isArray(b)
  )
    return false;
  const aa = a as Record<string, unknown>,
    bb = b as Record<string, unknown>;
  const keys = new Set([...Object.keys(aa), ...Object.keys(bb)]);
  return [...keys].every((key) => modelEqual(aa[key], bb[key]));
}

/** Apply only changes visible to an editor; original fields the editor omitted remain owned by the source. */
export function patchModel<T>(
  original: T,
  before: T | undefined,
  after: T | undefined,
): T | undefined {
  if (modelEqual(before, after)) return structuredClone(original);
  if (
    !before ||
    !after ||
    typeof before !== 'object' ||
    typeof after !== 'object' ||
    Array.isArray(before) ||
    Array.isArray(after) ||
    before instanceof Date ||
    after instanceof Date ||
    before instanceof Uint8Array ||
    after instanceof Uint8Array
  )
    return structuredClone(after);
  // T's shape is retained: only its own keys are copied, patched or removed, recursively.
  const result: Record<string, unknown> =
    original && typeof original === 'object' ? {...original} : {};
  const previous = before as Record<string, unknown>,
    current = after as Record<string, unknown>;
  for (const key of new Set([...Object.keys(previous), ...Object.keys(current)])) {
    if (key === '__proto__' || key === 'constructor' || key === 'prototype') continue;
    if (!modelEqual(previous[key], current[key])) {
      const value = patchModel(result[key], previous[key], current[key]);
      if (value === undefined) delete result[key];
      else result[key] = value;
    }
  }
  return result as T;
}

export function patchKeyed<T>(
  original: readonly T[],
  before: readonly T[],
  after: readonly T[],
  key: (value: T) => string,
): T[] {
  const result = new Map(original.map((value) => [key(value), value])),
    old = new Map(before.map((value) => [key(value), value])),
    next = new Map(after.map((value) => [key(value), value]));
  for (const id of new Set([...old.keys(), ...next.keys()])) {
    if (modelEqual(old.get(id), next.get(id))) continue;
    const value = patchModel(result.get(id), old.get(id), next.get(id));
    if (value === undefined) result.delete(id);
    else result.set(id, value);
  }
  return [...result.values()];
}
