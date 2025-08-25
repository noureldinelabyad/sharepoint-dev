import { Skill } from "../services/models";

type Rule = { rank: number; rx: RegExp };
const RULES: Rule[] = [
  { rank: 5, rx: /\b(expert|experte|expertin|principal|architect)\b/i },
  { rank: 4, rx: /\b(advanced|fortgeschritten|senior|professional|profi|specialist)\b/i },
  { rank: 3, rx: /\b(associate|intermediate|mittel|mittelstufe)\b/i },
  { rank: 2, rx: /\b(foundation|fundamentals|basic|grundkenntnisse)\b/i },
  { rank: 1, rx: /\b(junior|beginner|einsteiger|newbie)\b/i },
];
const TRAILING_EXPERT = /[\s:\-–—]\s*expert\b/i;
const TRAILING_ADV    = /[\s:\-–—]\s*(advanced|fortgeschritten)\b/i;

function _rankText(t?: string): number {
  if (!t) return 2;
  if (TRAILING_EXPERT.test(t)) return 5;
  if (TRAILING_ADV.test(t)) return 4;
  for (const r of RULES) if (r.rx.test(t)) return r.rank;
  return 2;
}

export function rankForSkill(s: Skill): number {
  const p = _rankText(s.proficiency);
  return p !== 2 ? p : _rankText(s.displayName);
}

export function sortSkillsByLevel(skills: Skill[]): Skill[] {
  return [...skills].sort((a, b) => {
    const rb = rankForSkill(b);
    const ra = rankForSkill(a);
    if (rb !== ra) return rb - ra;
    return (a.displayName || "").localeCompare(b.displayName || "", undefined, { sensitivity: "base" });
  });
}

export function effectiveProficiency(s: Skill): string | undefined {
  if (s.proficiency) return s.proficiency;
  const r = rankForSkill(s);
  if (r === 5) return "Expert";
  if (r === 4) return "Advanced";
  if (r === 3) return "Associate";
  return undefined;
}
