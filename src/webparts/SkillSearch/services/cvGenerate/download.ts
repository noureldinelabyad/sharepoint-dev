import { SPHttpClient } from '@microsoft/sp-http';

export async function downloadArrayBuffer(spHttp: SPHttpClient, urlOrServerRel: string): Promise<ArrayBuffer> {
  if (/^https?:\/\//i.test(urlOrServerRel)) {
    const r = await fetch(urlOrServerRel);
    if (!r.ok) throw new Error(`HTTP ${r.status} on ${urlOrServerRel}`);
    return r.arrayBuffer();
  }

  const origin = window.location.origin;
  const m = urlOrServerRel.match(/^\/(sites|teams)\/[^/]+/i);
  const webPath = m ? m[0] : '';
  const api = `${origin}${webPath}/_api/web/GetFileByServerRelativePath(DecodedUrl='${encodeURIComponent(urlOrServerRel)}')/$value`;

  const res = await spHttp.get(api, SPHttpClient.configurations.v1, {
    headers: { 'binaryStringRequestBody': 'true' } as any
  } as any);

  if (!res.ok) {
    const text = await res.text().catch(() => '');
    throw new Error(`HTTP ${res.status} downloading ${urlOrServerRel}.\n${text.slice(0, 200)}`);
  }
  return res.arrayBuffer();
}
