import { Person, Skill } from "../services/models";

/** UI-facing labels for levels */
export type SkillLevel = "Expert" | "Advanced" | "Associate" | "Foundation" | "Beginner";

export interface FilterState {
  depts: Set<string>;      // normalized department names (accents removed, lowercase)
  levels: Set<SkillLevel>; // selected level labels
}

export const emptyFilterState = (): FilterState => ({
  depts: new Set<string>(),
  levels: new Set<SkillLevel>()
});

// ---------- helpers (same semantics as your skill ranking) ----------
const RULES: Array<{ rank: number; rx: RegExp }> = [
  { rank: 5, rx: /\b(expert|experte|expertin|principal|architect)\b/i },
  { rank: 4, rx: /\b(advanced|fortgeschritten|senior|professional|profi|specialist)\b/i },
  { rank: 3, rx: /\b(associate|intermediate|mittel|mittelstufe)\b/i },
  { rank: 2, rx: /\b(foundation|fundamentals|basic|grundkenntnisse)\b/i },
  { rank: 1, rx: /\b(junior|beginner|einsteiger|newbie)\b/i }
];
const TRAILING_EXPERT = /[\s:\-–—]\s*expert\b/i;
const TRAILING_ADV    = /[\s:\-–—]\s*(advanced|fortgeschritten)\b/i;

const norm = (s?: string) =>
  (s || "").toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");

function rankFromText(t?: string): number {
  if (!t) return 2;
  if (TRAILING_EXPERT.test(t)) return 5;
  if (TRAILING_ADV.test(t))    return 4;
  for (const r of RULES) if (r.rx.test(t)) return r.rank;
  return 2;
}

function rankForSkill(s: Skill): number {
  const p = rankFromText(s.proficiency);
  return p !== 2 ? p : rankFromText(s.displayName);
}

const labelForRank = (r: number): SkillLevel =>
  r >= 5 ? "Expert" :
  r === 4 ? "Advanced" :
  r === 3 ? "Associate" :
  r === 2 ? "Foundation" : "Beginner";

// ---------- public helpers ----------
/** true if a person has any skill in one of the selected levels */
export function personMatchesLevels(p: Person, selected: Set<SkillLevel>): boolean {
  if (!selected.size) return true; // no level filter -> allow
  for (const s of (p.skills || [])) {
    const lvl = labelForRank(rankForSkill(s));
    if (selected.has(lvl)) return true;
  }
  return false;
}

/** true if a person's department is among selected (case/diacritics-insensitive) */
export function personMatchesDepartments(p: Person, selected: Set<string>): boolean {
  if (!selected.size) return true; // no dept filter -> allow
  const d = norm(p.department);
  return d ? selected.has(d) : false;
}

/** Apply both filters */
export function applyFilters(people: Person[], state: FilterState): Person[] {
  return people.filter(p => personMatchesDepartments(p, state.depts) &&
                            personMatchesLevels(p, state.levels));
}

/** Build a unique, display-ready department list for the UI. */
export function collectDepartments(people: Person[]): string[] {
  const byKey = new Map<string, string>();
  for (const p of people) {
    if (!p.department) continue;
    const k = norm(p.department);
    if (k && !byKey.has(k)) byKey.set(k, p.department); // keep 1st original-case variant
  }
  return Array.from(byKey.values()).sort((a, b) =>
    a.localeCompare(b, undefined, { sensitivity: "base" }));
}

/** Level options for the UI (ordered high→low). */
export const LEVEL_OPTIONS: SkillLevel[] = [
  "Expert", "Advanced", "Associate", "Foundation", "Beginner"
];

/** Normalize a department string to key form (exported for UI convenience). */
export const normDeptKey = norm;
