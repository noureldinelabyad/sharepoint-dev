// Skills enrichment (v1.0 + beta fallback)

import { MSGraphClientV3 } from '@microsoft/sp-http';
import { BATCH_SIZE } from './constants';
import { chunk } from './utils';
import { Person } from './models';

/**
 * Enriches Person[].skills using a hybrid strategy:
 * 1) v1.0 /users/{key}?$select=skills (string[])
 * 2) beta /users/{id}/profile/skills (displayName, proficiency) for empties
 */
export class SkillsService {
  constructor(private client: MSGraphClientV3) {}

  async enrich(users: Person[]): Promise<void> {
    if (!users.length) return;

    // Phase 1 — v1.0 user.skills (strings)
    const groups = chunk(users, BATCH_SIZE);
    const indexesNeedingBeta: number[] = [];

    for (let g = 0; g < groups.length; g++) {
      const group = groups[g];
      const req = {
        requests: group.map((u, i) => {
          const globalIdx = g * BATCH_SIZE + i;
          const key = encodeURIComponent(u.userPrincipalName || u.id);
          return { id: String(globalIdx), method: 'GET', url: `/users/${key}?$select=skills` };
        })
      };

      try {
        const res = await this.client.api('/$batch').version('v1.0').post(req);
        if (res?.responses) {
          res.responses.forEach((r: any) => {
            const idx = parseInt(r.id, 10);
            if (r.status === 200 && Array.isArray(r.body?.skills) && r.body.skills.length) {
              users[idx].skills = r.body.skills.map((s: string) => ({ displayName: s }));
            } else {
              indexesNeedingBeta.push(idx);
            }
          });
        }
      } catch {
        // if the batch failed, mark entire group for beta
        for (let i = 0; i < group.length; i++) indexesNeedingBeta.push(g * BATCH_SIZE + i);
      }
    }

    const stillEmpty = indexesNeedingBeta.filter(i => !users[i].skills?.length);
    if (!stillEmpty.length) return;

    // Phase 2 — beta profile skills (with proficiency)
    const betaGroups = chunk(stillEmpty, BATCH_SIZE);
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
        if (res?.responses) {
          res.responses.forEach((r: any) => {
            const idx = parseInt(r.id, 10);
            if (r.status === 200 && Array.isArray(r.body?.value) && r.body.value.length) {
              users[idx].skills = r.body.value.map((s: any) => ({
                displayName: s.displayName, proficiency: s.proficiency
              }));
            }
          });
        }
      } catch {
        // ignore — beacouse there are users without skills
      }
    }
  }
}
