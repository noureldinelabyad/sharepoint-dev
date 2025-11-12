// Directory/fallback reads + mapping/filtering

import { MSGraphClientV3 } from '@microsoft/sp-http';
import { PeopleResult, Person } from './models';
import { SERVICE_LIKE_DENY, ALLOWED_EMAIL_RX, HAS_NO_ROLE } from './constants';

/**
 * Repository for directory user reads.
 * Maps raw Graph users to our Person model and filters out service-like accounts.
 */
export class UsersRepository {
  constructor(private client: MSGraphClientV3) {}

  /**
   * Get a page of active human users (Members/Guests) with licenses.
   * Returns a PeopleResult with items + nextLink for paging.
   *
   * NOTE: The assignedLicenses/$count filter uses "ne 0" (or "gt 0" depending on tenant).
   * We use "ne 0" to avoid the "Operator: 'Greater' is not supported" error seen on some tenants.
   */
  async getActiveUsersPage(pageSize = 500, next?: string): Promise<PeopleResult> {
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
      .filter("accountEnabled eq true and (userType eq 'Member') and assignedLicenses/$count ne 0"+
       "and (endswith(mail,'@thinformatics.com') or endswith(userPrincipalName,'@thinformatics.com')) " +
        "and (jobTitle ne null or department ne null)" )
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
  async getRelevantPeople(pageSize = 100): Promise<PeopleResult> {
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
    mail: this.pickPreferredEmail(u),
    userPrincipalName: u.userPrincipalName,
    photoUrl: undefined,
    skills: []
  });

  /** Exclude obvious service/system accounts. */
  private filterServiceLike(users: Person[]): Person[] {
    return users.filter(u =>
      u.displayName && !SERVICE_LIKE_DENY.test(u.displayName) &&
      u.userPrincipalName && !SERVICE_LIKE_DENY.test(u.userPrincipalName) &&
      this.hasAllowedDomain(u) &&
      !HAS_NO_ROLE(u.jobTitle, u.department)
    );
  }

  private pickPreferredEmail(u: any): string | undefined {
  const candidates: string[] = [
    u?.mail,
    ...(Array.isArray(u?.otherMails) ? u.otherMails : []),
    u?.userPrincipalName
  ].filter(Boolean);
  // Prefer company domain; otherwise fall back to first available
  const preferred = candidates.find(e => ALLOWED_EMAIL_RX.test(e));
  return preferred ?? candidates[0];
}

private hasAllowedDomain(p: { mail?: string; userPrincipalName: string }): boolean {
  const e = p.mail ?? p.userPrincipalName;
  return !!e && ALLOWED_EMAIL_RX.test(e);
}


}
