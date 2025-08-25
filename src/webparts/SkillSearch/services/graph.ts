 // Facade that composes all of services

import { MSGraphClientV3 } from '@microsoft/sp-http';
import { PeopleResult, Me } from './models';
import { UsersRepository } from './users';
import { SkillsService } from './skillService';
import { PhotosService } from './PhotoService';
import { MeService } from './meService';

/**
 * Facade that composes repositories/services to expose a simple
 * API to the UI: getMe, getPeoplePage, getPeopleFallback.
 *
 * Keeps the same method names as the original GraphService
 * so Profile.tsx (skilssearch.tsx) needs minimal changes.
 */
export class GraphFacade {
  private usersRepo: UsersRepository;
  private photos: PhotosService;
  private skills: SkillsService;
  private meSvc: MeService;

  constructor( client: MSGraphClientV3) {
    this.usersRepo = new UsersRepository(client);
    this.photos = new PhotosService(client);
    this.skills = new SkillsService(client);
    this.meSvc = new MeService(client);
  }

  /** Load the signed-in user w/ photo, about & skills. */
  public async getMe(): Promise<Me> {
    return this.meSvc.getMe();
  }

  /** Page over active human users; enrich with photos & skills. */
  public async getPeoplePage(pageSize = 20, next?: string): Promise<PeopleResult> {
    const res = await this.usersRepo.getActiveUsersPage(pageSize, next);
    await this.photos.enrich(res.items);
    await this.skills.enrich(res.items);
    return res;
  }

  /** Fallback list when directory read isn’t consented. */
  public async getPeopleFallback(pageSize = 20): Promise<PeopleResult> {
    const res = await this.usersRepo.getRelevantPeople(pageSize);
    await this.photos.enrich(res.items);
    await this.skills.enrich(res.items);
    return res;
  }
}
