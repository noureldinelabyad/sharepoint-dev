import { SkillGroup } from '../types';

export function extractSkills(htmlDoc: Document): { flatSkills: string[]; skillGroups: SkillGroup[] } {
  const tables = Array.from(htmlDoc.getElementsByTagName('table')) as HTMLTableElement[];

  for (const t of tables) {
    const header = Array.from(t.rows && t.rows[0] ? t.rows[0].cells : [])
      .map(c => (c.textContent || '').trim().toLowerCase());

    if (!header.some(h => h.includes('beschreibung'))) continue;

    const groups: SkillGroup[] = [];
    for (const tr of Array.from(t.rows).slice(1)) {
      const cells = Array.from(tr.cells)
        .map(td => (td.textContent || '').replace(/\s+/g, ' ').trim())
        .filter(Boolean);

      if (cells.length < 2) continue;

      const category = cells[0];
      const items = cells[1].split(/,\s*/).map(x => x.trim()).filter(Boolean);
      if (category && items.length) groups.push({ category, items });
    }

    const flat = groups.reduce((acc: string[], g: SkillGroup) => {
      acc.push(g.category);
      for (const it of g.items) acc.push(it);
      return acc;
    }, []);

    return { flatSkills: flat, skillGroups: groups };
  }

  return { flatSkills: [], skillGroups: [] };
}
