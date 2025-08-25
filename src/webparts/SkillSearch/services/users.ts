// Directory/fallback reads + mapping/filtering

import { MSGraphClientV3 } from '@microsoft/sp-http';
import { PeopleResult, Person } from './models';
import { SERVICE_LIKE_DENY } from './constants';

/**
 * Repository for directory user reads.
 * Maps raw Graph users to our Person model and filters out service-like accounts.
 */
export class UsersRepository {
  constructor(private client: MSGraphClientV3) {}

  /**
   * Get a page of active human users (Members/Guests) with licenses.
   * Returns a PeopleResult with items + nextLink for paging.
   */
  async getActiveUsersPage(pageSize = 20, next?: string): Promise<PeopleResult> {
    if (next) {
      const path = next.replace('https://graph.microsoft.com/v1.0', '');
      const res = await this.client.api(path).get();
      let users = (res.value as any[]).map(this.mapUser);
      users = this.filterServiceLike(users);
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

    let users = (res.value as any[]).map(this.mapUser);
    users = this.filterServiceLike(users);
    return { items: users, nextLink: res['@odata.nextLink'] };
  }

  /**
   * Fallback when directory-list permissions aren’t approved:
   * Returns a relevant “people” list for the current user.
   */
  async getRelevantPeople(pageSize = 20): Promise<PeopleResult> {
    const res = await this.client.api('/me/people')
      .select('id,displayName,userPrincipalName,jobTitle,department,scoredEmailAddresses')
      .top(pageSize)
      .get();

    const primaryEmail = (p: any): string | undefined =>
    p?.scoredEmailAddresses?.[0]?.address
    ?? p?.emailAddresses?.[0]?.address
    ?? p?.userPrincipalName;

    let people: Person[] = (res.value as any[]).map((p) => ({
      id: p.id,
      displayName: p.displayName,
      jobTitle: p.jobTitle,
      department: p.department,
      mail: primaryEmail(p),
      userPrincipalName: p.userPrincipalName || '',
      photoUrl: undefined,
      skills: []
    }));

    people = this.filterServiceLike(people);
    return { items: people };
  }

  /** Map raw Graph user -> Person (includes mail fallbacks). */
  private mapUser = (u: any): Person => ({
    id: u.id,
    displayName: u.displayName,
    jobTitle: u.jobTitle,
    department: u.department,
    mail: u.mail || (Array.isArray(u.otherMails) ? u.otherMails[0] : undefined) || u.userPrincipalName,
    userPrincipalName: u.userPrincipalName,
    photoUrl: undefined,
    skills: []
  });

  /** Exclude obvious service/system accounts. */
  private filterServiceLike(users: Person[]): Person[] {
    return users.filter(u =>
      u.displayName && !SERVICE_LIKE_DENY.test(u.displayName) &&
      u.userPrincipalName && !SERVICE_LIKE_DENY.test(u.userPrincipalName)
    );
  }
}
