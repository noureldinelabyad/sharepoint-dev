export function splitToItems(value: string): string[] {
  return value.split(/[,;•\n]/).map(entry => entry.trim()).filter(Boolean);
}

export function uniqueKeepOrder(arr: string[]): string[] {
  const seen = new Set<string>();
  const out: string[] = [];
  for (const value of arr) {
    const trimmed = value.trim();
    if (!trimmed) continue;
    if (seen.has(trimmed)) continue;
    seen.add(trimmed);
    out.push(trimmed);
  }
  return out;
}
