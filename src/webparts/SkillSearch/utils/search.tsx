import * as React from "react";
import { Me, Person } from "../services/models";

export const norm = (s?: string) =>
  (s || "").toString().toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");

export const tokenize = (q: string) => norm(q).split(/\s+/).filter(Boolean);

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

/** convert text + tokens into React nodes (no dangerouslySetInnerHTML) */
export function highlightToNodes(text: string, tokens: string[]): (string | JSX.Element)[] {
  if (!tokens.length) return [text];
  // pick first matching token for light-weight highlighting
  const t = tokens.find(tok => tok && norm(text).includes(tok));
  if (!t) return [text];

  const rx = new RegExp(`(${t.replace(/[.*+?^${}()|[\\]\\\\]/g, "\\$&")})`, "ig");
  const parts = text.split(rx);
  return parts.map((part, i) =>
    norm(part) === t ? <mark key={i}>{part}</mark> : <span key={i}>{part}</span>
  );
}
