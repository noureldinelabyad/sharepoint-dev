 // Facade that composes all of services

import { MSGraphClientV3 } from '@microsoft/sp-http';
import { PeopleResult, Me } from './models';
import { UsersRepository } from './users';
import { SkillsService } from './skillService';
import { MeService } from './meService';

import { Client } from '@microsoft/microsoft-graph-client';
import { WebPartContext } from '@microsoft/sp-webpart-base';


let _client: Client | null = null;

export async function getGraphClient(ctx: WebPartContext): Promise<Client> {
  if (_client) return _client;
  // SPFx wires MSAL token, caching, scopes
  _client = await ctx.msGraphClientFactory.getClient('3') as unknown as Client;
  return _client;
}

/**
 * Facade that composes repositories/services to expose a simple
 * API to the UI: getMe, getPeoplePage, getPeopleFallback.
 *
 * Keeps the same method names as the original GraphService
 * so Profile.tsx (skilssearch.tsx) needs minimal changes.
 */
export class GraphFacade {
  private usersRepo: UsersRepository;
  private skills: SkillsService;
  private meSvc: MeService;

  constructor( client: MSGraphClientV3) {
    this.usersRepo = new UsersRepository(client);
    this.skills = new SkillsService(client);
    this.meSvc = new MeService(client);
  }

  /** Load the signed-in user w/ photo, about & skills. */
  public async getMe(): Promise<Me> {
    return this.meSvc.getMe();
  }

  /** Page over active human users; enrich with photos & skills. */
  public async getPeoplePage(pageSize = 200, next?: string): Promise<PeopleResult> {
    const res = await this.usersRepo.getActiveUsersPage(pageSize, next);
    await this.skills.enrich(res.items);
    return res;
  }

  /** Fallback list when directory read isn’t consented. */
  public async getPeopleFallback(pageSize = 100): Promise<PeopleResult> {
    const res = await this.usersRepo.getRelevantPeople(pageSize);
    await this.skills.enrich(res.items);
    return res;
  }
}
