import * as React from "react";
import { Me, Person, Skill } from "../services/models";

/** lowercase + accent-insensitive */
export const norm = (s?: string) =>
  (s || "")
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "");

/** split the search box text into tokens */
export const tokenize = (q: string): string[] =>
  norm(q).split(/\s+/).filter(Boolean);

/** Does a person match ALL tokens across any of the searchable fields? */
export function matches(p: Person | Me, tokens: string[]): boolean {
  if (!tokens.length) return true;
  const name  = norm(p.displayName);
  const mail  = norm(p.mail || p.userPrincipalName);
  const job   = norm(p.jobTitle);
  const dept  = norm(p.department);
  const skill = norm((p.skills || []).map(s => s.displayName).join(" "));
  return tokens.every(t =>
    name.includes(t) || mail.includes(t) || job.includes(t) || dept.includes(t) || skill.includes(t)
  );
}

/** Does a skill name match ANY token? (used for prioritising chips) */
export function skillMatches(name: string, tokens: string[]): boolean {
  if (!tokens.length) return false;
  const n = norm(name);
  return tokens.some(t => n.includes(t));
}

/** Stable-partition skills: matches first, then the rest (preserve relative order). */
export function prioritiseSkills(skills: Skill[], tokens: string[]): Skill[] {
  if (!tokens.length) return skills;
  const hits: Skill[] = [];
  const rest: Skill[] = [];
  for (const s of skills) (skillMatches(s.displayName, tokens) ? hits : rest).push(s);
  return [...hits, ...rest];
}

/**
 * Highlight ALL token occurrences in `text` and return React nodes.
 * - case/accent-insensitive
 * - merges overlaps
 * - safe (no innerHTML)
 */
export function highlightToNodes(text: string, tokens: string[]): React.ReactNode {
  if (!tokens.length || !text) return text;

  const src = text;
  const normSrc = norm(text);

  type Span = { start: number; end: number };
  const spans: Span[] = [];

  // collect all hit ranges in the normalised string
  for (const t of tokens) {
    const q = norm(t);
    if (!q) continue;
    let i = 0;
    while ((i = normSrc.indexOf(q, i)) !== -1) {
      spans.push({ start: i, end: i + q.length });
      i += Math.max(1, q.length);
    }
  }
  if (!spans.length) return src;

  // merge overlaps
  spans.sort((a, b) => a.start - b.start);
  const merged: Span[] = [];
  for (const s of spans) {
    const last = merged[merged.length - 1];
    if (!last || s.start > last.end) merged.push({ ...s });
    else last.end = Math.max(last.end, s.end);
  }

  // map positions in normalised string back to original string
  const indexMap: number[] = [];
  let orig = 0;
  for (const ch of src) {
    const n = norm(ch) || ch; // at least one step
    for (let k = 0; k < n.length; k++) indexMap.push(orig);
    orig += ch.length;
  }
  const toOrig = (p: number) => Math.min(indexMap[p] ?? src.length, src.length);

  // build nodes
  const out: React.ReactNode[] = [];
  let cursor = 0;
  merged.forEach((m, i) => {
    const s = toOrig(m.start);
    const e = toOrig(m.end);
    if (s > cursor) out.push(src.slice(cursor, s));
    out.push(<mark key={i}>{src.slice(s, e)}</mark>);
    cursor = e;
  });
  if (cursor < src.length) out.push(src.slice(cursor));
  return out;
}
