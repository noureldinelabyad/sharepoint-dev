// Skills enrichment (v1.0 + beta merge; de-dupe; Graph $batch safe)

import { MSGraphClientV3 } from '@microsoft/sp-http';
import { Person, Skill } from './models';
import { BATCH_SIZE } from './constants';
import { chunk } from './utils';

// ---------- Graph response types ----------

interface V1BatchItem {
  id: string;
  status: number;
  body?: { skills?: (string | null | undefined)[] };
}

interface GraphProfileSkill {
  displayName?: string | null;
  proficiency?: string | null;
}

interface BetaBatchItem {
  id: string;
  status: number;
  body?: { value?: GraphProfileSkill[] };
}

type BetaSkill = { displayName: string; proficiency?: string };

// Graph enforces max 20 requests per $batch
const GRAPH_BATCH_LIMIT = BATCH_SIZE ;
const EFFECTIVE_BATCH = Math.min(
  typeof BATCH_SIZE === 'number' && BATCH_SIZE > 0 ? BATCH_SIZE : GRAPH_BATCH_LIMIT,
  GRAPH_BATCH_LIMIT
);

/**
 * Strategy
 * 1) v1.0  /users/{key}?$select=skills               -> string[]
 * 2) beta  /users/{id}/profile/skills?$select=...    -> {displayName, proficiency}
 *    Merge beta into v1 results (never overwrite existing names, only enrich).
 */
export class SkillsService {
  constructor(private client: MSGraphClientV3) {}

  public async enrich(users: Person[]): Promise<void> {
    if (!users?.length) return;

    // ----- Pass 1 — v1.0 user.skills (strings) -----
    const v1Groups = chunk(users, EFFECTIVE_BATCH);
    for (let g = 0; g < v1Groups.length; g++) {
      const group = v1Groups[g];
      const base = g * EFFECTIVE_BATCH;

      const req = {
        requests: group.map((u, i) => {
          const idx = base + i;
          const key = encodeURIComponent(u.userPrincipalName || u.id);
          return { id: String(idx), method: 'GET', url: `/users/${key}?$select=skills` };
        })
      };

      try {
        const res = await this.client.api('/$batch').version('v1.0').post(req);
        const responses = Array.isArray(res?.responses) ? (res.responses as V1BatchItem[]) : [];
        for (const r of responses) {
          const idx = Number.parseInt(String(r.id), 10);
          if (!Number.isFinite(idx) || r.status !== 200) continue;

          const raw = r.body?.skills ?? [];
          const names: string[] = raw
            .map(s => (s ?? '').toString().trim())
            .filter(s => s.length > 0);

          if (names.length) this.mergeV1(users[idx], names);
        }
      } catch {
        // ignore; beta pass may still populate skills
      }
    }

    // ----- Pass 2 — beta profile skills (with proficiency) for ALL users; merge/enrich -----
    const allIdx = Array.from({ length: users.length }, (_, i) => i);
    const betaGroups = chunk(allIdx, EFFECTIVE_BATCH);

    for (const grp of betaGroups) {
      const req = {
        requests: grp.map((idx) => ({
          id: String(idx),
          method: 'GET',
          url: `/users/${users[idx].id}/profile/skills?$select=displayName,proficiency`
        }))
      };

      try {
        const res = await this.client.api('/$batch').version('beta').post(req);
        const responses = Array.isArray(res?.responses) ? (res.responses as BetaBatchItem[]) : [];
        for (const r of responses) {
          const idx = Number.parseInt(String(r.id), 10);
          if (!Number.isFinite(idx) || r.status !== 200) continue;

          const raw = r.body?.value ?? [];
          if (!raw.length) continue;

          const items: BetaSkill[] = raw
            .map((b: GraphProfileSkill): BetaSkill => ({
              displayName: (b.displayName ?? '').toString().trim(),
              proficiency: b.proficiency ?? undefined
            }))
            .filter((b: BetaSkill) => b.displayName.length > 0);

          if (items.length) this.mergeBeta(users[idx], items);
        }
      } catch {
        // some tenants/users won’t have profile skills — that’s fine
      }
    }
  }

  // ---------- merge helpers ----------

  /** Merge plain string skills (v1.0) without duplicates. */
  private mergeV1(user: Person, names: string[]): void {
    user.skills ||= [];
    const map = this.mapByName(user.skills);

    for (const name of names) {
      const key = this.norm(name);
      if (!key) continue;
      if (!map.has(key)) {
        const s: Skill = { displayName: name };
        user.skills.push(s);
        map.set(key, s);
      }
    }
  }

  /** Merge beta skills (adds proficiency and missing skill rows). */
  private mergeBeta(user: Person, beta: BetaSkill[]): void {
    user.skills ||= [];
    const map = this.mapByName(user.skills);

    for (const b of beta) {
      const key = this.norm(b.displayName);
      if (!key) continue;

      const existing = map.get(key);
      if (existing) {
        // enrich, don’t overwrite displayName
        if (b.proficiency) existing.proficiency = b.proficiency;
      } else {
        const s: Skill = { displayName: b.displayName, proficiency: b.proficiency };
        user.skills.push(s);
        map.set(key, s);
      }
    }
  }

  private mapByName(skills: Skill[]): Map<string, Skill> {
    const m = new Map<string, Skill>();
    for (const s of skills) {
      const key = this.norm(s.displayName);
      if (key && !m.has(key)) m.set(key, s);
    }
    return m;
  }

  private norm(s?: string): string {
    return (s ?? '')
      .toString()
      .trim()
      .toLowerCase()
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '');
  }
}
