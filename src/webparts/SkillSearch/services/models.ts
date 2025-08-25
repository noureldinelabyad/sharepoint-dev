// Types used across the feature

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

export interface Me extends Person {
  /** “About me” text (Delve/MyAccount). */
  aboutMe?: string;
  /** “Ask me about” keywords. */
  responsibilities?: string[];
}

export interface PeopleResult {
  items: Person[];
  nextLink?: string;
}
