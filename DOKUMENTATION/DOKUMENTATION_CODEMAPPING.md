# Dokumentation – Code-Block Mapping

Dieses Dokument ordnet die Code-Beispiele aus der Implementierungsphase den realen Dateien im Projekt zu.

---

## 5.1 Implementierung der Datenstrukturen

### Beispiel: Datenmodell „Person"

**Dokumentation – Code Block:**
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
  skills: Skill[];
}
```

**Datei:** [src/webparts/SkillSearch/services/models.ts](src/webparts/SkillSearch/services/models.ts#L1-L11)

---

### Datenmapping – mapUser Funktion

**Dokumentation – Code Block:**
```typescript
private mapUser(raw: any): Person {
  return {
    id: raw.id,
    displayName: raw.displayName,
    jobTitle: raw.jobTitle,
    department: raw.department,
    mail: raw.mail,
    userPrincipalName: raw.userPrincipalName,
    skills: []
  };
}
```

**Datei:** [src/webparts/SkillSearch/services/users.ts](src/webparts/SkillSearch/services/users.ts#L119-L128)

**Note:** Die reale Implementierung ist erweitert mit E-Mail-Fallback-Logik:
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
  const preferred = candidates.find(e => ALLOWED_EMAIL_RX.test(e));
  return preferred ?? candidates[0];
}
```

---

## 5.2 Implementierung der Geschäftslogik

### GraphFacade – zentrale Steuerung

**Dokumentation – Code Block:**
```typescript
public async getPeople(searchText: string): Promise<Person[]> {
  const users = await this.usersRepository.getRelevantPeople();
  
  await this.skillsService.enrich(users);
  
  return users.filter(user =>
    user.displayName.toLowerCase().includes(searchText.toLowerCase())
  );
}
```

**Datei:** [src/webparts/SkillSearch/services/graph.ts](src/webparts/SkillSearch/services/graph.ts#L28-L59)

**Note:** Die reale Implementierung hat zwei separate Methoden mit Paging-Support:
```typescript
public async getPeoplePage(pageSize = 200, next?: string): Promise<PeopleResult> {
  const res = await this.usersRepo.getActiveUsersPage(pageSize, next);
  await this.skills.enrich(res.items);
  return res;
}

public async getPeopleFallback(pageSize = 100): Promise<PeopleResult> {
  const res = await this.usersRepo.getRelevantPeople(pageSize);
  await this.skills.enrich(res.items);
  return res;
}
```

---

### Filterlogik – filterBySkill Funktion

**Dokumentation – Code Block:**
```typescript
function filterBySkill(users: Person[], skill: string): Person[] {
  return users.filter(user =>
    user.skills.some(s =>
      s.displayName.toLowerCase().includes(skill.toLowerCase())
    )
  );
}
```

**Datei:** [src/webparts/SkillSearch/utils/filters.ts](src/webparts/SkillSearch/utils/filters.ts#L45-L53)

**Note:** Die reale Implementierung verwendetzu eine verfeinerte Filterung nach Skill-Level (Expert, Advanced, Associate, Foundation, Beginner):

```typescript
export function personMatchesLevels(p: Person, selected: Set<SkillLevel>): boolean {
  if (!selected.size) return true;
  for (const s of (p.skills || [])) {
    const lvl = labelForRank(rankForSkill(s));
    if (selected.has(lvl)) return true;
  }
  return false;
}

export function applyFilters(people: Person[], state: FilterState): Person[] {
  return people.filter(p => personMatchesDepartments(p, state.depts) &&
                            personMatchesLevels(p, state.levels));
}
```

---

### Ausschluss unerwünschter Accounts

**Dokumentation – Code Block:**
```typescript
function filterSystemAccounts(users: Person[]): Person[] {
  return users.filter(user =>
    !user.userPrincipalName.includes("extern") &&
    !user.userPrincipalName.includes("service")
  );
}
```

**Datei:** [src/webparts/SkillSearch/services/users.ts](src/webparts/SkillSearch/services/users.ts#L129-L139) und [src/webparts/SkillSearch/services/constants.ts](src/webparts/SkillSearch/services/constants.ts#L9)

**Note:** Die reale Implementierung verwendet erweiterte Regex-Patterns:

```typescript
// From constants.ts
const SERVICE_LIKE_DENY_SRC = '(thinformatics |svc|service|automation|bot|daemon|system|noreply|no-reply|do-not-reply|admin)';
export const SERVICE_LIKE_DENY = new RegExp(SERVICE_LIKE_DENY_SRC, 'i');

// From users.ts
private filterServiceLike(users: Person[]): Person[] {
  return users.filter(u =>
    u.displayName && !SERVICE_LIKE_DENY.test(u.displayName) &&
    u.userPrincipalName && !SERVICE_LIKE_DENY.test(u.userPrincipalName) &&
    this.hasAllowedDomain(u) &&
    !HAS_NO_ROLE(u.jobTitle, u.department)
  );
}
```

---

### RBAC-Logik

**Dokumentation – Code Block:**
```typescript
function canViewProfile(currentUserRole: string): boolean {
  return ["Sales", "HR", "Head"].includes(currentUserRole);
}
```

**Datei:** [src/webparts/SkillSearch/services/constants.ts](src/webparts/SkillSearch/services/constants.ts#L12-L13)

**Note:** Die reale Implementierung nutzen einen flexibleren Regex-Ansatz:

```typescript
export const RBA_ALLOW_SRC = 'head\\s*of|hr|sales|ceo|trainee';
export const RBA_ALLOW = new RegExp(`\\b(?:${RBA_ALLOW_SRC})\\b`, 'i');

export const HAS_NO_ROLE = (job?: string, dept?: string): boolean =>
  isBlank(job) && isBlank(dept);
```

Diese werden in [src/webparts/SkillSearch/services/users.ts](src/webparts/SkillSearch/services/users.ts#L129-L139) verwendet:

```typescript
private filterServiceLike(users: Person[]): Person[] {
  return users.filter(u =>
    // ... other checks ...
    !HAS_NO_ROLE(u.jobTitle, u.department)
  );
}
```

---

### ProfileOrdner Integration

**Dokumentation – Code Block:**
```typescript
function getProfileLink(user: Person): string {
  return /sites/profiles/${user.userPrincipalName};
}
```

**Datei:** [src/webparts/SkillSearch/services/constants.ts](src/webparts/SkillSearch/services/constants.ts#L6) und [src/webparts/SkillSearch/ui/components/PersonCard.tsx](src/webparts/SkillSearch/ui/components/PersonCard.tsx#L21-L31)

**Note:** Die reale Implementierung nutzen eine zentrale Konstante:

```typescript
// From constants.ts
export const profilesUrl = "https://thinformatics.sharepoint.com/sites/Beraterprofile/Freigegebene%20Dokumente/Forms/AllItems.aspx?as=json";

// From PersonCard.tsx - wird als Prop übergeben und verwendet
interface PersonCardProps {
  person: Person;
  profilesUrl: string;
  // ... other props ...
}

// Usage – opens the profile library when user clicks the button
<a href={profilesUrl} target="_blank" rel="noopener noreferrer">
  Profil öffnen
</a>
```

---

### CV-Export (Teilimplementierung)

**Dokumentation – Code Block:**
```typescript
public generateCV(profile: ProfileData): Blob {
  return this.templateService.fillTemplate(profile);
}
```

**Datei:** [src/webparts/SkillSearch/services/cvGenerate/index.ts](src/webparts/SkillSearch/services/cvGenerate/index.ts#L13-L25)

**Note:** Die reale Implementierung hat zwei Modi: Template-basiert und Fallback-basiert:

```typescript
export async function tryGenerateDataportDoc(
  spHttp: SPHttpClient,
  maybeTemplateUrlOrServerRel: string | null,
  data: ProfileData
): Promise<Blob> {
  if (maybeTemplateUrlOrServerRel) {
    try {
      return await fillDataportTemplate(spHttp, maybeTemplateUrlOrServerRel, data);
    } catch (err) {
      console.warn('Falling back to generated Dataport layout...', err);
    }
  }
  return buildDataportDocx(data);
}

export function defaultTemplateUrl(): string {
  return `${BERATERPROFIL_SITE}/Dataport CV Vorlage - TAGGED.docx`;
}
```

---

## Zusammenfassung der Dateien

| Komponente | Haupt-Datei | Typ |
|---|---|---|
| **Datenmodelle** | `services/models.ts` | Interfaces/Types |
| **User Mapping & Filtering** | `services/users.ts` | Repository |
| **Graph API Facade** | `services/graph.ts` | Service/Facade |
| **Filter-Logik** | `utils/filters.ts` | Utilities |
| **RBAC & Konstanten** | `services/constants.ts` | Constants |
| **ProfileOrdner** | `ui/components/PersonCard.tsx` | Component |
| **CV-Generierung** | `services/cvGenerate/index.ts` | Service |

---

## Notizen zur Dokumentation

1. **Models** – vollständig wie dokumentiert
2. **User Services** – komplexer mit Email-Fallback-Logik
3. **Facades** – nutzen Paging-Support für große Datenmengen
4. **Filter** – erweitert mit Skill-Level-Klassifizierung
5. **RBAC** – verwendet reguläre Ausdrücke für mehr Flexibilität
6. **CV-Export** – hat Template + Fallback-Mechanismus

