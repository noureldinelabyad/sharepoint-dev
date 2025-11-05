import { SPHttpClient } from '@microsoft/sp-http';
import { BERATERPROFIL_SITE } from './profileConstants';

export interface ProfileFile {
  fileUrl: string;     // server-relative URL to *.docx
  folderUrl: string;   // server-relative URL to the person's folder
  libRootUrl: string;  // server-relative URL of the document library root
}

/** Find the real document library root for /sites/Beraterprofile */
async function getDocLibRoot(spHttp: SPHttpClient, absWebUrl: string): Promise<string | null> {
  const origin = absWebUrl.replace(/\/sites\/.*/i, ''); // https://tenant.sharepoint.com
  const url =
    `${origin}${BERATERPROFIL_SITE}/_api/web/lists` +
    `?$filter=BaseTemplate eq 101 and Hidden eq false` +
    `&$select=Title,RootFolder/ServerRelativeUrl` +
    `&$expand=RootFolder`;

  const resp = await spHttp.get(url, SPHttpClient.configurations.v1);
  if (!resp.ok) return null;

  const data = await resp.json();
  const libs: any[] = data?.value ?? [];
  if (!libs.length) return null;

  // Prefer localized names, else *Dokumente/*Documents, else first visible library
  const pick =
    libs.find(l => String(l?.RootFolder?.ServerRelativeUrl).includes('Freigegebene Dokumente')) ||
    libs.find(l => String(l?.RootFolder?.ServerRelativeUrl).includes('Shared Documents')) ||
    libs.find(l => /\/(Dokumente|Documents)$/i.test(String(l?.RootFolder?.ServerRelativeUrl))) ||
    libs[0];

  return pick?.RootFolder?.ServerRelativeUrl ?? null;
}

async function folderExists(spHttp: SPHttpClient, absWebUrl: string, folderServerRel: string): Promise<boolean> {
  const origin = absWebUrl.replace(/\/sites\/.*/i, '');
  const api =
    `${origin}${BERATERPROFIL_SITE}/_api/web/` +
    `GetFolderByServerRelativePath(DecodedUrl='${encodeURIComponent(folderServerRel)}')` +
    `?$select=ServerRelativeUrl`;
  const res = await spHttp.get(api, SPHttpClient.configurations.v1);
  return res.ok;
}

/** Resolve the person's folder and the library root (tries "First Last", then "Last First"). */
async function resolvePersonFolder(
  spHttp: SPHttpClient,
  absWebUrl: string,
  displayName: string
): Promise<{ folderUrl: string; libRootUrl: string } | null> {
  const libRootUrl = await getDocLibRoot(spHttp, absWebUrl);
  if (!libRootUrl) return null;

  const candidates = [displayName.trim()];
  const parts = displayName.trim().split(/\s+/);
  if (parts.length === 2) candidates.push(`${parts[1]} ${parts[0]}`);

  for (const name of candidates) {
    const folderUrl = `${libRootUrl}/${name}`;
    if (await folderExists(spHttp, absWebUrl, folderUrl)) {
      return { folderUrl, libRootUrl };
    }
  }
  return null;
}

/** Build a working AllItems.aspx link using the site's preferred "id=" syntax. */
export async function buildFolderViewUrlAsync(
  spHttp: SPHttpClient,
  absWebUrl: string,
  serverRelWebUrl: string,
  displayName: string
): Promise<string | null> {
  const origin = absWebUrl.replace(serverRelWebUrl, ''); // https://tenant.sharepoint.com
  const resolved = await resolvePersonFolder(spHttp, absWebUrl, displayName);
  if (!resolved) return null;
  const id = encodeURIComponent(resolved.folderUrl);
  return `${origin}${resolved.libRootUrl}/Forms/AllItems.aspx?id=${id}`; // Profilordner URL
}

/** Find newest Beraterprofil*.docx in the person's folder. */
export async function findLatestProfileDocx(
  spHttp: SPHttpClient,
  absWebUrl: string,
  displayName: string
): Promise<ProfileFile | null> {
  const origin = absWebUrl.replace(/\/sites\/.*/i, '');
  const resolved = await resolvePersonFolder(spHttp, absWebUrl, displayName);
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

// ---------- Template resolution ----------

async function fileExists(spHttp: SPHttpClient, absWebUrl: string, serverRelFile: string): Promise<boolean> {
  const origin = absWebUrl.replace(/\/sites\/.*/i, '');
  const m = serverRelFile.match(/^\/(sites|teams)\/[^/]+/i);
  const webPath = m ? m[0] : BERATERPROFIL_SITE;
  const api =
    `${origin}${webPath}/_api/web/GetFileByServerRelativePath(DecodedUrl='${encodeURIComponent(serverRelFile)}')` +
    `?$select=ServerRelativeUrl`;
  const res = await spHttp.get(api, SPHttpClient.configurations.v1);
  return res.ok;
}

/** Resolve "Dataport CV Vorlage.docx" with no Templates folder:
 *  Order now is: <Lib>/Dataport…  → <Lib>/Dataport_CV_Vorlage…  → <PersonFolder>/Dataport…
 */
export async function resolveDataportTemplateUrl(
  spHttp: SPHttpClient,
  absWebUrl: string,
  libRootUrl: string,
  personFolderUrl: string
): Promise<string | null> {
   const candidates = [
    `${libRootUrl}/Dataport CV Vorlage.docx`,
    `${libRootUrl}/Dataport CV Vorlage - TAGGED.docx`, // 
    `${libRootUrl}/Dataport_CV_Vorlage.docx`,
    `${libRootUrl}/Dataport-CV-Vorlage.docx`,
    `${personFolderUrl}/Dataport CV Vorlage.docx`,
    `${personFolderUrl}/Dataport CV Vorlage - TAGGED.docx`, // 
    `${personFolderUrl}/Dataport_CV_Vorlage.docx`,
  ];

  for (const c of candidates) {
    if (await fileExists(spHttp, absWebUrl, c)) return c;
  }
  return null;
}
