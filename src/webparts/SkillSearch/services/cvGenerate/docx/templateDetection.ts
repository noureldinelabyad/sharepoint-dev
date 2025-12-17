import PizZip from 'pizzip';
import { SPHttpClient } from '@microsoft/sp-http';
import { downloadArrayBuffer } from '../download';

export async function templateHasDocxtemplaterTags(
  spHttp: SPHttpClient,
  templateUrlOrServerRel: string
): Promise<boolean> {
  try {
    const buf = await downloadArrayBuffer(spHttp, templateUrlOrServerRel);
    const zip = new PizZip(buf);
    const xml = zip.file('word/document.xml')?.asText() || '';
    return /\{\{[^}]+\}\}/.test(xml) || /\{#.+?\}/.test(xml) || /\{\/.+?\}/.test(xml);
  } catch {
    return false;
  }
}
