// Tunables (batch size, regex, etc.)

/** 20 is the Max number of requests per Graph $batch. */
export const BATCH_SIZE = 20;

export  const profilesUrl = "https://thinformatics.sharepoint.com/sites/Beraterprofile/Freigegebene%20Dokumente/Forms/AllItems.aspx?as=json";

/** Exclude obvious service/system accounts by name/UPN. */
const SERVICE_LIKE_DENY_SRC = '(thinformatics |svc|service|automation|bot|daemon|system|noreply|no-reply|do-not-reply|admin)';
export const SERVICE_LIKE_DENY = new RegExp(SERVICE_LIKE_DENY_SRC, 'i');

/** Roll Based Access accounts by jobtitle. */
export const RBA_ALLOW_SRC = 'head\\s*of|hr|sales|ceo|trainee';
export const RBA_ALLOW = new RegExp(`\\b(?:${RBA_ALLOW_SRC})\\b`, 'i');

/** Keep only users whose email/UPN ends with this domain. */
export const ALLOWED_DOMAIN = 'thinformatics.com';
export const ALLOWED_EMAIL_RX = new RegExp(`@${ALLOWED_DOMAIN.replace(/\./g, '\\.')}$`, 'i');

export const isBlank = (s?: string): boolean => !s || !String(s).trim();

/** True when BOTH job title and department are empty */
export const HAS_NO_ROLE = (job?: string, dept?: string): boolean =>
  isBlank(job) && isBlank(dept);

// Graph supports: 48, 64, 96, 120, 240, 360, 432, 504, 648
export const AVATAR = {
  list: 64,   // result cards
  hero: 96    // top "me" card
} as const;
export type AvatarKind = keyof typeof AVATAR;
