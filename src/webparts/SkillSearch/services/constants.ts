// Tunables (batch size, regex, etc.)

/** Max number of requests per Graph $batch. */
export const BATCH_SIZE = 20;

/** Exclude obvious service/system accounts by name/UPN. */
export const SERVICE_LIKE_DENY = /(svc|service|automation|bot|daemon|system|noreply|no-reply|do-not-reply|admin)/i;
