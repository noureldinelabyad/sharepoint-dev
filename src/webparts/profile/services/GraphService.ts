import { MSGraphClientV3 } from '@microsoft/sp-http';
import { ResponseType } from '@microsoft/microsoft-graph-client';

/** ===== Types ===== */
export interface PeopleResult { items: Person[]; nextLink?: string; }
export interface Skill { displayName: string; proficiency?: string; }

export interface Person {
  id: string;
  displayName: string;
  jobTitle?: string;
  department?: string;
  mail?: string;
  userPrincipalName: string;
  photoUrl?: string;
  skills: Skill[];
}

export interface Me extends Person {
  aboutMe?: string;
  responsibilities?: string[]; // "Ask me about"
}

/** ===== Service ===== */
export class GraphService {
  constructor(private client: MSGraphClientV3) {}

  /** Me (photo + about + responsibilities + skills) */
  public async getMe(): Promise<Me> {
    const [meBasic, meProf] = await Promise.all([
      this.client.api('/me')
        .select('id,displayName,jobTitle,department,mail,userPrincipalName,aboutMe,responsibilities')
        .get(),
      this.client.api('/me/profile')
        .version('beta')
        .expand('skills($select=displayName,proficiency)')
        .get()
        .catch(() => ({ skills: [] }))
    ]);

    const photoUrl =
      (await this.getPhotoDataUrl('/me/photo/$value')) ||
      (await this.getPhotoDataUrl('/me/photos/96x96/$value')) ||
      this.makeInitialsAvatar(meBasic.displayName, 96);

    return {
      id: meBasic.id,
      displayName: meBasic.displayName,
      jobTitle: meBasic.jobTitle,
      department: meBasic.department,
      mail: meBasic.mail,
      userPrincipalName: meBasic.userPrincipalName,
      aboutMe: meBasic.aboutMe,
      responsibilities: meBasic.responsibilities,
      skills: (meProf?.skills ?? []).map((s: any) => ({
        displayName: s.displayName,
        proficiency: s.proficiency
      })),
      photoUrl
    };
  }

  /**
   * Page over ACTIVE human users:
   * - accountEnabled = true
   * - userType in ('Member','Guest')  (keeps managers; excludes app/service principals anyway)
   * - assignedLicenses/$count > 0      (filters most service/system accounts)
   * Then we still apply a small client-side “service-like” name filter.
   */
  public async getPeoplePage(pageSize: number = 20, next?: string): Promise<PeopleResult> {
    if (next) {
      const path = next.replace('https://graph.microsoft.com/v1.0', '');
      const res = await this.client.api(path).get();
      let users = (res.value as any[]).map(this._mapUser);
      users = this._filterServiceLike(users);
      await this._enrichPhotos(users);
      await this._enrichSkillsHybrid(users);
      return { items: users, nextLink: res['@odata.nextLink'] };
    }

    const res = await this.client.api('/users')
      .version('v1.0')
      .header('ConsistencyLevel', 'eventual')
      .count(true)
      .select('id,displayName,jobTitle,department,mail,otherMails,userPrincipalName,accountEnabled,userType,assignedLicenses')
      .filter("accountEnabled eq true and (userType eq 'Member' or userType eq 'Guest') and assignedLicenses/$count gt 0")
      .orderby('displayName')
      .top(pageSize)
      .get();

    let users = (res.value as any[]).map(this._mapUser);
    users = this._filterServiceLike(users);
    await this._enrichPhotos(users);
    await this._enrichSkillsHybrid(users);

    return { items: users, nextLink: res['@odata.nextLink'] };
  }

  /** Fallback when directory perms aren’t approved: /me/people (relevance list) */
  public async getPeopleFallback(pageSize: number = 20): Promise<PeopleResult> {
    const res = await this.client.api('/me/people')
      .select('id,displayName,userPrincipalName,jobTitle,department')
      .top(pageSize)
      .get();

    let people: Person[] = (res.value as any[]).map((p) => ({
      id: p.id,
      displayName: p.displayName,
      jobTitle: p.jobTitle,
      department: p.department,
      mail: p.mail || p.userPrincipalName,
      userPrincipalName: p.userPrincipalName || '',
      photoUrl: undefined,
      skills: []
    }));

    people = this._filterServiceLike(people);
    await this._enrichPhotos(people);
    await this._enrichSkillsHybrid(people);

    return { items: people, nextLink: undefined }; // /me/people doesn’t give us a nextLink
  }

  /** ===== Helpers ===== */

  private _mapUser = (u: any): Person => ({
    id: u.id,
    displayName: u.displayName,
    jobTitle: u.jobTitle,
    department: u.department,
    mail: u.mail || (Array.isArray(u.otherMails) ? u.otherMails[0] : undefined) || u.userPrincipalName,
    userPrincipalName: u.userPrincipalName,
    skills: []
  });

  /** Soft filter to exclude obvious service accounts that slipped through */
  private _filterServiceLike(users: Person[]): Person[] {
    const deny = /(svc|service|automation|bot|daemon|system|noreply|no-reply|do-not-reply|admin)/i;
    return users.filter(u =>
      u.displayName && !deny.test(u.displayName) &&
      u.userPrincipalName && !deny.test(u.userPrincipalName)
    );
  }

  /** Prefer /photo/$value; fall back to sized variant; then initials avatar */
  private async _enrichPhotos(users: Person[]): Promise<void> {
    await Promise.all(users.map(async (u) => {
      const key = encodeURIComponent(u.userPrincipalName || u.id);
      u.photoUrl =
        (await this.getPhotoDataUrl(`/users/${key}/photo/$value`)) ||
        (await this.getPhotoDataUrl(`/users/${key}/photos/96x96/$value`)) ||
        this.makeInitialsAvatar(u.displayName, 96);
    }));
  }

  /**
   * Hybrid skills enrichment:
   * 1) v1.0 /users/{key}?$select=skills (string[]) via $batch (20/request)
   * 2) For those still empty → beta /users/{id}/profile/skills (proficiency)
   */
  private async _enrichSkillsHybrid(users: Person[]): Promise<void> {
    if (!users.length) return;

    // Phase 1: v1.0 user.skills
    const groups = chunk(users, 20);
    const unresolvedIdx: number[] = [];

    for (let g = 0; g < groups.length; g++) {
      const group = groups[g];

      // IMPORTANT: request id = GLOBAL user index (not 0..19)
      const req = {
       requests: group.map((u, i) => {
        const globalIdx = g * 20 + i;
        const key = encodeURIComponent(u.userPrincipalName || u.id);
        return { id: String(globalIdx), method: 'GET', url: `/users/${key}?$select=skills` };
      })
      };

      try {
        const res = await this.client.api('/$batch').version('v1.0').post(req);
       if (res?.responses) {
        res.responses.forEach((r: any) => {
          const idxGlobal = parseInt(r.id, 10);           // <- map by id, not position
          if (r.status === 200 && Array.isArray(r.body?.skills) && r.body.skills.length) {
            users[idxGlobal].skills = r.body.skills.map((s: string) => ({ displayName: s }));
          }
        });
      }
    } catch (e) {
        // Whole batch failed → mark this chunk unresolved
        for (let i = 0; i < group.length; i++) unresolvedIdx.push(g * 20 + i);
        console.warn('v1.0 skills batch failed:', e);
      }
    }

    // Figure out who still has no skills
    const pending: number[] = [];
    users.forEach((u, i) => { if (!u.skills || u.skills.length === 0) pending.push(i); });
    if (!pending.length) return;


    // Phase 2: beta profile skills (proficiency)
    const betaGroups = chunk(pending, 20);
    for (const grp of betaGroups) {
    const req = {
      requests: grp.map((globalIdx) => ({
        id: String(globalIdx),                                // <- global idx
        method: 'GET',
        url: `/users/${users[globalIdx].id}/profile/skills?$select=displayName,proficiency`
      }))
    };
      try {
      const res = await this.client.api('/$batch').version('beta').post(req);
      if (res?.responses) {
        res.responses.forEach((r: any) => {
          const idx = parseInt(r.id, 10);                     // <- map by id, not position
          if (r.status === 200 && Array.isArray(r.body?.value) && r.body.value.length) {
            users[idx].skills = r.body.value.map((s: any) => ({
              displayName: s.displayName,
              proficiency: s.proficiency
              }));
            }
          });
        }
      } catch (e) {
        console.warn('beta profile skills batch failed (ok to ignore):', e);
      }
    }
  }

  /** Blob -> data URL (browser-friendly) */
  private async getPhotoDataUrl(path: string): Promise<string | undefined> {
    try {
      const blob: Blob = await this.client.api(path)
        .responseType(ResponseType.BLOB)
        .get();

      const dataUrl = await new Promise<string>((resolve, reject) => {
        const reader = new FileReader();
        reader.onloadend = () => resolve(reader.result as string);
        reader.onerror = reject;
        reader.readAsDataURL(blob);
      });

      return dataUrl;
    } catch {
      return undefined;
    }
  }

  /** Teams-like initials avatar (SVG -> Data URL) */
  private makeInitialsAvatar(name: string, size: number = 96): string {
    const initials = this._initials(name);
    const bg = this._colorFromName(name);
    const svg =
      `<svg xmlns='http://www.w3.org/2000/svg' width='${size}' height='${size}'>` +
      `<rect width='100%' height='100%' rx='${size/2}' ry='${size/2}' fill='${bg}'/>` +
      `<text x='50%' y='54%' dominant-baseline='middle' text-anchor='middle' ` +
      `font-family='Segoe UI, Arial, sans-serif' font-size='${Math.round(size*0.42)}' fill='#fff' font-weight='600'>` +
      `${this._escapeXml(initials)}</text></svg>`;
    return `data:image/svg+xml;utf8,${encodeURIComponent(svg)}`;
  }

  private _initials(name: string): string {
    if (!name) return '?';
    const parts = name.trim().split(/\s+/);
    const first = parts[0]?.[0] ?? '';
    const second = parts.length > 1 ? parts[1][0] : (parts[0]?.[1] ?? '');
    return (first + second).toUpperCase();
  }

  private _colorFromName(name: string): string {
    // Simple hash -> 12 nice hues
    const palette = ['#3867D6','#20BF6B','#F7B731','#EB3B5A','#8854D0','#0FB9B1','#4B7BEC','#26DE81','#FED330','#FA8231','#A55EEA','#2D98DA'];
    let h = 0; for (let i = 0; i < name.length; i++) h = (h * 31 + name.charCodeAt(i)) >>> 0;
    return palette[h % palette.length];
  }

  private _escapeXml(s: string): string {
    return s.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&apos;');
  }
}

/** Utility: chunk an array into size-N groups */
function chunk<T>(arr: T[], size: number): T[][] {
  const out: T[][] = [];
  for (let i = 0; i < arr.length; i += size) out.push(arr.slice(i, i + size));
  return out;
}
