// Tunables (batch size, regex, etc.)

/** 20 is the Max number of requests per Graph $batch. */
export const BATCH_SIZE = 20;

/** Exclude obvious service/system accounts by name/UPN. */
export const SERVICE_LIKE_DENY = /(svc|service|automation|bot|daemon|system|noreply|no-reply|do-not-reply|Mailbox|admin)/i;

/** Keep only users whose email/UPN ends with this domain. */
export const ALLOWED_DOMAIN = 'thinformatics.com';
export const ALLOWED_EMAIL_RX = new RegExp(`@${ALLOWED_DOMAIN.replace(/\./g, '\\.')}$`, 'i');