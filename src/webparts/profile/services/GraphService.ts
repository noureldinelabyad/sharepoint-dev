import { MSGraphClientV3 } from '@microsoft/sp-http';
import { ResponseType } from '@microsoft/microsoft-graph-client';


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
  responsibilities?: string[]; // "ask me about"
}

export class GraphService {
  constructor(private client: MSGraphClientV3) {}

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

    const photoUrl = await this.getPhotoDataUrl('/me/photos/96x96/$value');

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
        displayName: s.displayName, proficiency: s.proficiency
      })),
      photoUrl
    };
  }

public async getPeopleTopN(n: number = 50): Promise<Person[]> {
  // try direct directory read first
  try {
    const usersRes = await this.client.api('/users')
      .select('id,displayName,jobTitle,department,mail,userPrincipalName')
      .orderby('displayName')
      .top(n)
      .get();

    const users: Person[] = usersRes.value.map((u: any) => ({
      id: u.id,
      displayName: u.displayName,
      jobTitle: u.jobTitle,
      department: u.department,
      mail: u.mail,
      userPrincipalName: u.userPrincipalName,
      skills: []
    }));

    // Try to enrich with skills (beta). If this 403s, we just keep empty skills.
    try {
      const batchReq = {
        requests: users.map((u, i) => ({
          id: String(i),
          method: 'GET',
          url: `/users/${u.id}/profile/skills?$select=displayName,proficiency`
        }))
      };
      const batch = await this.client.api('/$batch').version('beta').post(batchReq);
      if (batch?.responses) {
        batch.responses.forEach((r: any) => {
          const idx = parseInt(r.id, 10);
          if (r.status === 200 && r.body?.value) {
            users[idx].skills = r.body.value.map((s: any) => ({
              displayName: s.displayName, proficiency: s.proficiency
            }));
          }
        });
      }
    } catch (e) {
      console.warn('Skill batch failed (likely permissions). Continuing without skills.', e);
    }

    // Photos via the endpoint you tested
    await Promise.all(users.map(async (u) => {
      const upn = encodeURIComponent(u.userPrincipalName);
      u.photoUrl = await this.getPhotoDataUrl(`/users/${upn}/photo/$value`);
    }));

    return users;
  } catch (e) {
    console.warn('Directory /users call failed; falling back to /me/people', e);
  }

  // Fallback: "People you work with" relevance list
  try {
    const pplRes = await this.client.api('/me/people')
      .select('id,displayName,userPrincipalName,jobTitle,department')
      .top(n)
      .get();

    const people: Person[] = pplRes.value.map((p: any) => ({
      id: p.id,
      displayName: p.displayName,
      jobTitle: p.jobTitle,
      department: p.department,
      mail: undefined,
      userPrincipalName: p.userPrincipalName || '',
      skills: []
    }));

    await Promise.all(people.map(async (p) => {
      const idOrUpn = encodeURIComponent(p.userPrincipalName || p.id);
      p.photoUrl = await this.getPhotoDataUrl(`/users/${idOrUpn}/photo/$value`);
    }));

    // Skills may require directory perms; try and ignore errors
    try {
      const batchReq = {
        requests: people.map((p, i) => ({
          id: String(i),
          method: 'GET',
          url: `/users/${p.id}/profile/skills?$select=displayName,proficiency`
        }))
      };
      const batch = await this.client.api('/$batch').version('beta').post(batchReq);
      if (batch?.responses) {
        batch.responses.forEach((r: any) => {
          const idx = parseInt(r.id, 10);
          if (r.status === 200 && r.body?.value) {
            people[idx].skills = r.body.value.map((s: any) => ({
              displayName: s.displayName, proficiency: s.proficiency
            }));
          }
        });
      }
    } catch (e) {
      console.warn('Skill batch (fallback) failed. Continuing without skills.', e);
    }

    return people;
  } catch (e2) {
    console.error('Both /users and /me/people failed.', e2);
    return [];
  }
}

  private async getPhotoDataUrl(path: string): Promise<string | undefined> {
        try {
            const blob: Blob = await this.client.api(path)
            .responseType(ResponseType.BLOB)
            .get();

            // Convert Blob -> data URL (base64) in the browser
            const dataUrl = await new Promise<string>((resolve, reject) => {
            const reader = new FileReader();
            reader.onloadend = () => resolve(reader.result as string);
            reader.onerror = reject;
            reader.readAsDataURL(blob);
            });

            return dataUrl;
        } 
        catch {
            return undefined; // fall back to silhouette in UI
        }
    }
}
