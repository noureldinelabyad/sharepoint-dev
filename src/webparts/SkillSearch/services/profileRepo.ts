import { SPHttpClient, ISPHttpClientOptions } from '@microsoft/sp-http';
import { BERATERPROFIL_SITE } from './profileConstants';

export interface ProfileFile {
  fileUrl: string;     // server-relative URL to *.docx
  folderUrl: string;   // server-relative URL to the person's folder
  libRootUrl: string;  // server-relative URL of the document library root
}

/* ---------------------------- Module-level caches --------------------------- */

type SiteKey = string;          // e.g. "https://tenant.sharepoint.com/sites/Beraterprofile"
type PersonKey = string;        // `${siteKey}|${normalizedDisplayName}`

const libRootCache = new Map<SiteKey, Promise<string | null>>();           // de-dup in flight
const personFolderCache = new Map<PersonKey, Promise<{ folderUrl: string; libRootUrl: string } | null>>();
const existsCache = new Map<string, Promise<boolean>>();                   // key = absolute API GET URL

/* ---------------------------------- Utils ---------------------------------- */

function getOrigin(absWebUrl: string): string {
  return new URL(absWebUrl).origin; // robust, no string replace
}
function siteKey(origin: string, site: string): SiteKey {
  // origin: https://tenant.sharepoint.com, site: /sites/Beraterprofile
  return `${origin}${site}`;
}
function normName(displayName: string): string {
  return displayName.trim().replace(/\s+/g, ' ').toLowerCase();
}

/* ------------------------ Library root: resolve ONCE ------------------------ */

/** Find the real document library root for /sites/Beraterprofile (cached). */
async function getDocLibRoot(spHttp: SPHttpClient, absWebUrl: string): Promise<string | null> {
  const origin = getOrigin(absWebUrl);
  const key = siteKey(origin, BERATERPROFIL_SITE);

  if (!libRootCache.has(key)) {
    const url =
      `${key}/_api/web/lists` +
      `?$filter=BaseTemplate eq 101 and Hidden eq false` +
      `&$select=Title,RootFolder/ServerRelativeUrl` +
      `&$expand=RootFolder`;

    const promise = (async () => {
      const resp = await spHttp.get(url, SPHttpClient.configurations.v1);
      if (!resp.ok) return null;

      const data = await resp.json();
      const libs: any[] = data?.value ?? [];
      if (!libs.length) return null;

      // Prefer localized defaults, else first visible lib
      const pick =
        libs.find(l => String(l?.RootFolder?.ServerRelativeUrl).includes('Freigegebene Dokumente')) ||
        libs.find(l => String(l?.RootFolder?.ServerRelativeUrl).includes('Shared Documents')) ||
        libs.find(l => /\/(Dokumente|Documents)$/i.test(String(l?.RootFolder?.ServerRelativeUrl))) ||
        libs[0];

      return pick?.RootFolder?.ServerRelativeUrl ?? null;
    })();

    libRootCache.set(key, promise);
  }

  return libRootCache.get(key)!;
}

/* ----------------------------- Cheap existence ----------------------------- */

async function httpGetOk(spHttp: SPHttpClient, url: string, opts?: ISPHttpClientOptions): Promise<boolean> {
  if (!existsCache.has(url)) {
    existsCache.set(url, (async () => {
      const res = await spHttp.get(url, SPHttpClient.configurations.v1, opts);
      return res.ok;
    })());
  }
  return existsCache.get(url)!;
}

async function folderExists(spHttp: SPHttpClient, absWebUrl: string, folderServerRel: string): Promise<boolean> {
  const origin = getOrigin(absWebUrl);
  const api =
    `${origin}${BERATERPROFIL_SITE}/_api/web/` +
    `GetFolderByServerRelativePath(DecodedUrl='${encodeURIComponent(folderServerRel)}')` +
    `?$select=ServerRelativeUrl`;
  return httpGetOk(spHttp, api);
}

async function fileExists(spHttp: SPHttpClient, absWebUrl: string, serverRelFile: string): Promise<boolean> {
  const origin = getOrigin(absWebUrl);
  // Resolve which web the path belongs to; fallback to configured site
  const m = serverRelFile.match(/^\/(sites|teams)\/[^/]+/i);
  const webPath = m ? m[0] : BERATERPROFIL_SITE;
  const api =
    `${origin}${webPath}/_api/web/GetFileByServerRelativePath(DecodedUrl='${encodeURIComponent(serverRelFile)}')` +
    `?$select=ServerRelativeUrl`;
  return httpGetOk(spHttp, api);
}

/* --------------------- Person folder: resolve + cache it -------------------- */

/** Resolve the person's folder and the library root (cached). */
async function resolvePersonFolder(
  spHttp: SPHttpClient,
  absWebUrl: string,
  displayName: string,
  libRootUrl?: string
): Promise<{ folderUrl: string; libRootUrl: string } | null> {
  const origin = getOrigin(absWebUrl);
  const site = siteKey(origin, BERATERPROFIL_SITE);
  const key: PersonKey = `${site}|${normName(displayName)}`;

  if (!personFolderCache.has(key)) {
    const promise = (async () => {
      const libRoot = libRootUrl ?? (await getDocLibRoot(spHttp, absWebUrl));
      if (!libRoot) return null;

      // Try "First Last" then "Last First"
      const dn = displayName.trim();
      const parts = dn.split(/\s+/);
      const candidates = parts.length === 2 ? [dn, `${parts[1]} ${parts[0]}`] : [dn];

      for (const name of candidates) {
        const folderUrl = `${libRoot}/${name}`;
        if (await folderExists(spHttp, absWebUrl, folderUrl)) {
          return { folderUrl, libRootUrl: libRoot };
        }
      }
      return null;
    })();

    personFolderCache.set(key, promise);
  }

  return personFolderCache.get(key)!;
}

/* ------------------------------ Public helpers ----------------------------- */

/** Build a working AllItems.aspx link using the site's preferred "id=" syntax. */
export async function buildFolderViewUrlAsync(
  spHttp: SPHttpClient,
  absWebUrl: string,
  serverRelWebUrl: string,   // kept for backward compatibility
  displayName: string,
  libRootUrl?: string        // pass-in to avoid recomputing
): Promise<string | null> {
  const origin = getOrigin(absWebUrl);
  const resolved = await resolvePersonFolder(spHttp, absWebUrl, displayName, libRootUrl);
  if (!resolved) return null;
  const id = encodeURIComponent(resolved.folderUrl);
  return `${origin}${resolved.libRootUrl}/Forms/AllItems.aspx?id=${id}`;
}

/** Find newest Beraterprofil*.docx in the person's folder. */
export async function findLatestProfileDocx(
  spHttp: SPHttpClient,
  absWebUrl: string,
  displayName: string,
  libRootUrl?: string
): Promise<ProfileFile | null> {
  const origin = getOrigin(absWebUrl);
  const resolved = await resolvePersonFolder(spHttp, absWebUrl, displayName, libRootUrl);
  if (!resolved) return null;

  const filesApi =
    `${origin}${BERATERPROFIL_SITE}/_api/web` +
    `/GetFolderByServerRelativePath(DecodedUrl='${encodeURIComponent(resolved.folderUrl)}')` +
    `/Files?$select=Name,TimeLastModified,ServerRelativeUrl`;

  const res = await spHttp.get(filesApi, SPHttpClient.configurations.v1);
  if (!res.ok) return null;

  const files = await res.json();
  if (!files?.value?.length) return null;

  const pick = files.value
    .filter((f: any) => String(f.Name).toLowerCase().endsWith('.docx'))
    .sort((a: any, b: any) => new Date(b.TimeLastModified).getTime() - new Date(a.TimeLastModified).getTime())
    .find((f: any) => /beraterprofil/i.test(String(f.Name))) || files.value[0];

  return {
    fileUrl: pick.ServerRelativeUrl,
    folderUrl: resolved.folderUrl,
    libRootUrl: resolved.libRootUrl
  };
}

/** Resolve "Dataport CV Vorlage*.docx" with no Templates folder (cached existence checks). */
export async function resolveDataportTemplateUrl(
  spHttp: SPHttpClient,
  absWebUrl: string,
  libRootUrl: string,
  personFolderUrl: string
): Promise<string | null> {
  const candidates = [
    // Prefer the person-scoped tagged template first so we keep per-profile formatting.
    `${personFolderUrl}/Dataport CV Vorlage - TAGGED.docx`,
    `${personFolderUrl}/Dataport CV Vorlage.docx`,
    `${personFolderUrl}/Dataport_CV_Vorlage.docx`,

    // Fall back to library-level tagged/default variants.
    `${libRootUrl}/Dataport CV Vorlage - TAGGED.docx`,
    `${libRootUrl}/Dataport CV Vorlage.docx`,
    `${libRootUrl}/Dataport_CV_Vorlage.docx`,
    `${libRootUrl}/Dataport-CV-Vorlage.docx`,
  ];

  // Check sequentially (keeps request count low); uses existsCache internally.
  for (const c of candidates) {
    if (await fileExists(spHttp, absWebUrl, c)) return c;
  }
  return null;
}

/* ------------------------------- Convenience ------------------------------- */

/** Optional: call once at app start to warm up lib root cache (no-op if called again). */
export async function ensureBeraterprofileLibRoot(spHttp: SPHttpClient, absWebUrl: string): Promise<string | null> {
  return getDocLibRoot(spHttp, absWebUrl);
}
