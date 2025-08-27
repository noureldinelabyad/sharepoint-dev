import { Skill } from "../services/models";

type Rule = { rank: number; rx: RegExp };
const RULES: Rule[] = [
  { rank: 5, rx: /\b(expert|experte|expertin|principal|architect)\b/i },
  { rank: 4, rx: /\b(advanced|fortgeschritten|senior|professional|profi|specialist)\b/i },
  { rank: 3, rx: /\b(associate|intermediate|mittel|mittelstufe)\b/i },
  { rank: 2, rx: /\b(foundation|fundamentals|basic|grundkenntnisse)\b/i },
  { rank: 1, rx: /\b(junior|beginner|einsteiger|newbie)\b/i },
];

// quick helpers to recognize level words in free text
const TRAILING_EXPERT = /[\s:\-–—]\s*expert\b/i;
const TRAILING_ADV    = /[\s:\-–—]\s*(advanced|fortgeschritten)\b/i;

function rankFromText(t?: string): number {
  if (!t) return 2;
  if (TRAILING_EXPERT.test(t)) return 5;
  if (TRAILING_ADV.test(t)) return 4;
  for (const r of RULES) if (r.rx.test(t)) return r.rank;
  return 2; // neutral / unknown
}

/** Sorting still uses the best rank we can infer from proficiency or name. */
export function rankForSkill(s: Skill): number {
  const p = rankFromText(s.proficiency);
  return p !== 2 ? p : rankFromText(s.displayName);
}

export function sortSkillsByLevel(skills: Skill[]): Skill[] {
  return [...skills].sort((a, b) => {
    const rb = rankForSkill(b);
    const ra = rankForSkill(a);
    if (rb !== ra) return rb - ra;
    return (a.displayName || "").localeCompare(b.displayName || "", undefined, { sensitivity: "base" });
  });
}

/**
 * Show a trailing label only when the skill name does NOT already include
 * a level word. Use the structured proficiency from Graph if available.
 */
export function effectiveProficiency(s: Skill): string | undefined {
  // If the name already contains "Expert/Advanced/…", don't add anything.
  if (rankFromText(s.displayName) !== 2) return undefined;

  // Otherwise show the normalized proficiency (when Graph provides it).
  const pRank = rankFromText(s.proficiency);
  if (pRank === 5) return "Expert";
  if (pRank === 4) return "Advanced";
  if (pRank === 3) return "Associate";
  return undefined;
}
