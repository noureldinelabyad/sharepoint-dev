# 5 Implementierungsphase

In der Implementierungsphase wurde die zuvor entworfene Architektur schrittweise umgesetzt. Ziel war es, die definierten Anforderungen in eine funktionierende Anwendung zu überführen.

Die Umsetzung erfolgte auf Basis des SharePoint Frameworks (SPFx) unter Verwendung von TypeScript und React. Dabei wurde besonderer Wert auf eine modulare Struktur sowie eine klare Trennung zwischen Benutzeroberfläche, Geschäftslogik und Datenzugriff gelegt.

Die Implementierung erfolgte iterativ. Einzelne Funktionen wurden schrittweise entwickelt und direkt getestet, um Fehler frühzeitig zu erkennen und zu beheben.

## 5.1 Implementierung der Datenstrukturen

Die Anwendung verwendet keine eigene Datenbank. Stattdessen werden die benötigten Daten zur Laufzeit aus der Microsoft Graph API geladen und in interne Datenstrukturen überführt.

Zur Abbildung der Daten wurden eigene TypeScript-Interfaces und Klassen definiert.

### Beispiel: Datenmodell „Person“

Aus der Datei `src/webparts/SkillSearch/services/models.ts`:

```typescript
export interface Skill {
  displayName: string;
  proficiency?: string;
}

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
```

Diese Struktur bildet die Grundlage für die Darstellung der Personenkarten innerhalb der Benutzeroberfläche.

### Datenmapping

Die Rohdaten aus der Graph API werden zunächst transformiert und in das interne Datenmodell überführt.

Aus der Datei `src/webparts/SkillSearch/services/users.ts`:

```typescript
private mapUser = (u: any): Person => ({
  id: u.id,
  displayName: u.displayName,
  jobTitle: u.jobTitle,
  department: u.department,
  mail: this.pickPreferredEmail(u),  // with fallback logic
  userPrincipalName: u.userPrincipalName,
  photoUrl: undefined,
  skills: []
});

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
```

Durch diese Abstraktion wird sichergestellt, dass Änderungen an der API keinen direkten Einfluss auf die UI-Komponenten haben.

## 5.2 Implementierung der Geschäftslogik

Die Geschäftslogik wurde in Form von Services und einer zentralen Facade umgesetzt.

### GraphFacade – zentrale Steuerung

Die Klasse GraphFacade dient als zentrale Schnittstelle zwischen UI und Datenzugriff.

Aus der Datei `src/webparts/SkillSearch/services/graph.ts`:

```typescript
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

  /** Fallback list when directory read isn't consented. */
  public async getPeopleFallback(pageSize = 100): Promise<PeopleResult> {
    const res = await this.usersRepo.getRelevantPeople(pageSize);
    await this.skills.enrich(res.items);
    return res;
  }
}
```

Diese Methoden zeigen den typischen Ablauf: Laden der Benutzer, Anreicherung mit Skills, und Bereitstellung für die UI mit Paging-Support.

### Filterlogik

Die Filterung erfolgt clientseitig basierend auf den gewählten Kriterien.

Aus der Datei `src/webparts/SkillSearch/utils/filters.ts`:

```typescript
/** true if a person has any skill in one of the selected levels */
export function personMatchesLevels(p: Person, selected: Set<SkillLevel>): boolean {
  if (!selected.size) return true; // no level filter -> allow
  for (const s of (p.skills || [])) {
    const lvl = labelForRank(rankForSkill(s));
    if (selected.has(lvl)) return true;
  }
  return false;
}

/** true if a person's department is among selected (case/diacritics-insensitive) */
export function personMatchesDepartments(p: Person, selected: Set<string>): boolean {
  if (!selected.size) return true; // no dept filter -> allow
  const d = norm(p.department);
  return d ? selected.has(d) : false;
}

export function applyFilters(people: Person[], state: FilterState): Person[] {
  return people.filter(p => personMatchesDepartments(p, state.depts) &&
                            personMatchesLevels(p, state.levels));
}
```

### Ausschluss unerwünschter Accounts

Ein wichtiger Bestandteil war die Bereinigung der Daten.

Aus der Datei `src/webparts/SkillSearch/services/constants.ts`:

```typescript
/** Exclude obvious service/system accounts by name/UPN. */
const SERVICE_LIKE_DENY_SRC = '(thinformatics |svc|service|automation|bot|daemon|system|noreply|no-reply|do-not-reply|admin)';
export const SERVICE_LIKE_DENY = new RegExp(SERVICE_LIKE_DENY_SRC, 'i');

/** Roll Based Access accounts by jobtitle. */
export const RBA_ALLOW_SRC = 'head\\s*of|hr|sales|ceo|trainee';
export const RBA_ALLOW = new RegExp(`\\b(?:${RBA_ALLOW_SRC})\\b`, 'i');

/** Keep only users whose email/UPN ends with this domain. */
export const ALLOWED_DOMAIN = 'thinformatics.com';
export const ALLOWED_EMAIL_RX = new RegExp(`@${ALLOWED_DOMAIN.replace(/\./g, '\\.')}$`, 'i');
```

Und die Anwendung in `src/webparts/SkillSearch/services/users.ts`:

```typescript
private filterServiceLike(users: Person[]): Person[] {
  return users.filter(u =>
    u.displayName && !SERVICE_LIKE_DENY.test(u.displayName) &&
    u.userPrincipalName && !SERVICE_LIKE_DENY.test(u.userPrincipalName) &&
    this.hasAllowedDomain(u) &&
    !HAS_NO_ROLE(u.jobTitle, u.department)
  );
}
```

Damit werden externe Konten und Systemkonten entfernt.

### RBAC-Logik

Die rollenbasierte Zugriffskontrolle steuert die Sichtbarkeit des ProfileOrdner-Buttons.

Aus der Datei `src/webparts/SkillSearch/services/constants.ts`:

```typescript
/** Roll Based Access accounts by jobtitle. */
export const RBA_ALLOW_SRC = 'head\\s*of|hr|sales|ceo|trainee';
export const RBA_ALLOW = new RegExp(`\\b(?:${RBA_ALLOW_SRC})\\b`, 'i');

export const HAS_NO_ROLE = (job?: string, dept?: string): boolean =>
  isBlank(job) && isBlank(dept);
```

Diese Logik wird in der UI verwendet, um den Zugriff zu steuern.

### ProfileOrdner Integration

Die Navigation zum Beraterprofil erfolgt über einen generierten SharePoint-Link.

Aus der Datei `src/webparts/SkillSearch/services/constants.ts`:

```typescript
export  const profilesUrl = "https://thinformatics.sharepoint.com/sites/Beraterprofile/Freigegebene%20Dokumente/Forms/AllItems.aspx?as=json";
```

Und die Verwendung in `src/webparts/SkillSearch/ui/components/PersonCard.tsx`:

```typescript
type Props = {
  person: Person;
  tokens: string[];
  onOpenSkills: (name: string, skills: Skill[]) => void;
  outlookUrl: (p: Person) => string;
  teamsUrl: (p: Person) => string;
  profilesUrl: string;
  spHttpClient: SPHttpClient;
  absWebUrl: string;
  serverRelWebUrl: string;
  msGraphClientFactory: any;
};
```

### CV-Export (Teilimplementierung)

Während der Implementierung wurde eine zusätzliche Funktion zur Erstellung von Profilen in einer externen Vorlage begonnen.

Aus der Datei `src/webparts/SkillSearch/services/cvGenerate/index.ts`:

```typescript
export async function tryGenerateDataportDoc(
  spHttp: SPHttpClient,
  maybeTemplateUrlOrServerRel: string | null,
  data: import('./types').ProfileData
): Promise<Blob> {
  if (maybeTemplateUrlOrServerRel) {
    try {
      return await fillDataportTemplate(spHttp, maybeTemplateUrlOrServerRel, data);
    } catch (err) {
      console.warn('Falling back to generated Dataport layout because template rendering failed', err);
    }
  }
  return buildDataportDocx(data);
}

export function defaultTemplateUrl(): string {
  //return `${BERATERPROFIL_SITE}/${LIB_INTERNAL_NAME}/Dataport CV Vorlage - TAGGED.docx`;
  return `${BERATERPROFIL_SITE}/Dataport CV Vorlage - TAGGED.docx`;
}
```

Diese Funktion ist aktuell nur teilweise umgesetzt und wird in zukünftigen Versionen erweitert.