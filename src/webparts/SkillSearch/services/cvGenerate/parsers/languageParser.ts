import { splitToItems, uniqueKeepOrder } from './listHelpers';

export function extractLanguages(tables: HTMLTableElement[], lines: string[]): string[] {
  for (const t of tables) {
    const headerCells = Array.from(t.rows && t.rows[0] ? t.rows[0].cells : [])
      .map(c => (c.textContent || '').trim().toLowerCase());

    const isLangTable = headerCells.some(h => h.includes('sprachen')) && headerCells.some(h => h.includes('niveau'));
    if (!isLangTable) continue;

    const res: string[] = [];
    for (const tr of Array.from(t.rows).slice(1)) {
      const lang = (tr.cells && tr.cells[0] ? tr.cells[0].textContent : '') || '';
      const level = (tr.cells && tr.cells[1] ? tr.cells[1].textContent : '') || '';
      const language = lang.trim();
      const languageLevel = level.trim();
      if (language) res.push(languageLevel ? `${language} (${languageLevel})` : language);
    }
    return uniqueKeepOrder(res);
  }

  const idx = lines.findIndex(l => /^sprachen\b/i.test(l) || /\bsprachen\b/i.test(l));
  if (idx >= 0) {
    const out: string[] = [];
    const firstLine = lines[idx];
    const after = firstLine.replace(/^.*?\bsprachen\b\s*:?\s*/i, '').trim();
    if (after) out.push(...splitToItems(after));

    const stopRe = /(kompetenzen|skills|kenntnisse|projekte|projektreferenzen|tätigkeitsbeschreibung|ausbildung|zertifikat|branchen|qualifikationen)/i;
    for (let i = idx + 1; i < lines.length; i++) {
      const line = (lines[i] || '').trim();
      if (!line) break;
      if (stopRe.test(line)) break;
      out.push(...splitToItems(line));
    }
    return uniqueKeepOrder(out).filter(Boolean);
  }

  return [];
}
