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

  /** Public surface expected by the hook/facade. */
  async getPeoplePage(pageSize = 100, next?: string, signal?: AbortSignal): Promise<PeopleResult> {
    return this.getActiveUsersPage(pageSize, next, signal);
  }

  /**
   * Page active human users with licenses.
   * Returns { items, nextLink } for paging.
   */
  async getActiveUsersPage(pageSize = 100, next?: string, signal?: AbortSignal): Promise<PeopleResult> {
    if (next) {
      // Graph client expects a relative URL; strip host if needed.
      const rel = next.startsWith('https://')
        ? next.replace(/^https:\/\/graph\.microsoft\.com\/v1\.0/i, '')
        : next;

      const req = this.client.api(rel);
      // pass AbortSignal if your SPFx supports it
      (req as any).config = { fetchOptions: { signal } };

      const res = await req.get();
      let users = (res.value as any[]).map(this.mapUser);
      users = this.filterServiceLike(users);
      return { items: users, nextLink: res['@odata.nextLink'] };
    }

    // Build filter (Members + Guests optional)
    const includeGuests = true; // flip if you want only Members
    const userTypeFilter = includeGuests
      ? "(userType eq 'Member' or userType eq 'Guest')"
      : "(userType eq 'Member')";

    const filter = [
      "accountEnabled eq true",
      userTypeFilter,
      "assignedLicenses/$count ne 0", // safer across tenants
      // restrict to company domain on the server (fewer rows)
      "(endswith(mail,'@thinformatics.com') or endswith(userPrincipalName,'@thinformatics.com'))",
      // only show people with at least some profile signal
      "(jobTitle ne null or department ne null)"
    ].join(' and ');

    const req = this.client.api('/users')
      .version('v1.0')
      .header('ConsistencyLevel', 'eventual') // needed for $count
      .count(true)
      .select('id,displayName,jobTitle,department,mail,otherMails,userPrincipalName,accountEnabled,userType,assignedLicenses')
      .filter(filter)
      .orderby('displayName')
      .top(pageSize);

    (req as any).config = { fetchOptions: { signal } };

    const res = await req.get();

    let users = (res.value as any[]).map(this.mapUser);
    users = this.filterServiceLike(users);
    return { items: users, nextLink: res['@odata.nextLink'] };
  }

  /** Fallback when directory-list permissions aren’t approved. */
  async getRelevantPeople(pageSize = 100, signal?: AbortSignal): Promise<PeopleResult> {
    const req = this.client.api('/me/people')
      .select('id,displayName,userPrincipalName,jobTitle,department,scoredEmailAddresses')
      .top(pageSize);

    (req as any).config = { fetchOptions: { signal } };

    const res = await req.get();

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
    ].filter(Boolean) as string[];
  // Prefer company domain; otherwise fall back to first available
    const preferred = candidates.find(e => ALLOWED_EMAIL_RX.test(e));
    return preferred ?? candidates[0];
  }

  private hasAllowedDomain(p: { mail?: string; userPrincipalName: string }): boolean {
    const e = p.mail ?? p.userPrincipalName;
    return !!e && ALLOWED_EMAIL_RX.test(e);
  }
}
