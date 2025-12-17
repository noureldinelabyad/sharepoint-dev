export interface ProjectItem {
  from?: string;
  to?: string;
  period?: string;

  // structured (used by the template)
  company?: string;
  headline?: string;
  description?: string;
  responsibilitiesTitle?: string;
  bullets?: string[];

  // legacy fields (kept for compatibility with older templates)
  title?: string;
  rolle: string;
  kunde: string;
  taetigkeiten: string[];
  technologien: string[];
}

export interface SkillGroup { category: string; items: string[]; }

export interface ProfileData {
  name: string;
  role: string;
  team: string;
  email?: string;
  skills: string[];
  summary: string;
  projects: ProjectItem[];
  skillGroups?: SkillGroup[];

  firstName?: string;
  lastName?: string;
  birthYear?: string;
  availableFrom?: string;
  einsatzAls?: string;
  einsatzIn?: string;
  languages?: string[];
  branchen?: string[];
  qualifikationen?: string[];
  berufserfahrung?: string;
  education?: string;
  profilnummer?: string;

  // photo from source docx (optional)
  photoBytes?: Uint8Array;
  photoExt?: string; // "png" | "jpg" | "jpeg"
}
