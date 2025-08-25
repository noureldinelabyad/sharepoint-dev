// “/me” profile + skills

import { MSGraphClientV3 } from '@microsoft/sp-http';
import { Me } from './models';
import { PhotosService } from './PhotoService';

export class MeService {
  constructor(private client: MSGraphClientV3) {}

  /**
   * Load the current user with about/responsibilities + skills,
   * and attach a photo (Graph or initials).
   */
  async getMe(): Promise<Me> {
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

    const photoSvc = new PhotosService(this.client);
    const me: Me = {
      id: meBasic.id,
      displayName: meBasic.displayName,
      jobTitle: meBasic.jobTitle,
      department: meBasic.department,
      mail: meBasic.mail,
      userPrincipalName: meBasic.userPrincipalName,
      aboutMe: meBasic.aboutMe,
      responsibilities: meBasic.responsibilities,
      photoUrl: undefined,
      skills: (meProf?.skills ?? []).map((s: any) => ({
        displayName: s.displayName,
        proficiency: s.proficiency
      }))
    };

    await photoSvc.enrich([me]);
    return me;
  }
}
