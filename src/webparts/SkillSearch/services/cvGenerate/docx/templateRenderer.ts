import { SPHttpClient } from '@microsoft/sp-http';
import Docxtemplater from 'docxtemplater';
import PizZip from 'pizzip';
import { MISSING_TOKEN } from '../../constants';
import { ProfileData } from '../types';
import { computeBerufserfahrungFromProjects } from '../parsers/experienceParser';
import { downloadArrayBuffer } from '../download';
import { replaceFirstBodyImageWithSourcePhoto } from './photoHandling';

export async function fillDataportTemplate(
  spHttp: SPHttpClient,
  templateUrlOrServerRel: string,
  profile: ProfileData
): Promise<Blob> {
  const buf = await downloadArrayBuffer(spHttp, templateUrlOrServerRel);
  const zip = new PizZip(buf);

  sanitizeDocxtemplaterZip(zip);

  const doc = new Docxtemplater(zip, {
    paragraphLoop: true,
    linebreaks: true,
    nullGetter: () => MISSING_TOKEN
  });

  const stringValue = (value?: string) => (value && value.trim() ? value.trim() : MISSING_TOKEN);
  const arrayValue = (items?: string[]) => (items && items.length ? items : [MISSING_TOKEN]);

  const beruf = profile.berufserfahrung && profile.berufserfahrung.trim()
    ? profile.berufserfahrung.trim()
    : computeBerufserfahrungFromProjects(profile.projects || []);

  const projects = (profile.projects || []).length
    ? profile.projects.map(project => ({
        period: stringValue(project.period),
        company: stringValue(project.company),
        headline: stringValue(project.headline),
        description: stringValue(project.description),
        responsibilitiesTitle: stringValue(project.responsibilitiesTitle || 'Verantwortlichkeiten:'),
        bullets: (project.bullets && project.bullets.length)
          ? project.bullets
              .map(bullet => (bullet ?? '').toString().trim())
              .filter(bullet => bullet.length > 0)
          : [MISSING_TOKEN]
      }))
    : [{
        period: MISSING_TOKEN,
        company: MISSING_TOKEN,
        headline: MISSING_TOKEN,
        description: MISSING_TOKEN,
        responsibilitiesTitle: 'Verantwortlichkeiten:',
        bullets: [MISSING_TOKEN]
      }];

  const model: any = {
    profilnummer: stringValue(profile.profilnummer),
    photo: '',

    firstName: stringValue(profile.firstName || (profile.name?.split(' ').slice(0, -1).join(' ') || profile.name || '')),
    lastName: stringValue(profile.lastName || (profile.name?.split(' ').slice(-1)[0] || '')),
    birthYear: stringValue(profile.birthYear),
    availableFrom: stringValue(profile.availableFrom),
    einsatzAls: stringValue(profile.einsatzAls || profile.role || ''),
    einsatzIn: stringValue(profile.einsatzIn),

    languages: arrayValue(profile.languages),
    languagesText: arrayValue(profile.languages).join('\n'),

    branchen: arrayValue(profile.branchen),
    branchenText: arrayValue(profile.branchen).join(', '),

    qualifikationen: arrayValue(profile.qualifikationen),
    qualifikationenText: arrayValue(profile.qualifikationen).join(', '),

    education: stringValue(profile.education),
    berufserfahrung: stringValue(beruf),

    name: stringValue(profile.name),
    role: stringValue(profile.role),
    team: stringValue(profile.team),
    email: stringValue(profile.email),
    summary: stringValue(profile.summary),

    skills: arrayValue(profile.skills),
    skillGroups: (profile.skillGroups && profile.skillGroups.length) ? profile.skillGroups : [{ category: MISSING_TOKEN, items: [MISSING_TOKEN] }],

    projects
  };

  doc.setData(model);

  try {
    doc.render();
  } catch (err: any) {
    console.error('Docxtemplater failed', err);
    console.error('Details:', err?.properties?.errors ?? err?.properties);
    throw err;
  }

  // Ensure both generated tokens and literal "Bitte anpassen!" are yellow
  highlightAndReplaceInZip(doc.getZip(), MISSING_TOKEN, 'Bitte anpassen!');
  highlightLiteralInZip(doc.getZip(), 'Bitte anpassen!');

  // Copy photo into the first embedded image in the body
  if (profile.photoBytes && profile.photoBytes.length) {
    try {
      await replaceFirstBodyImageWithSourcePhoto(doc.getZip(), profile.photoBytes, profile.photoExt || '');
    } catch (e) {
      console.warn('Photo copy failed (ignored):', e);
    }
  }

  return doc.getZip().generate({ type: 'blob' });
}

export function sanitizeDocxtemplaterZip(zip: any) {
  const xmlParts = Object.keys(zip.files).filter(p =>
    /^word\/(document|header\d+|footer\d+)\.xml$/i.test(p)
  );

  for (const part of xmlParts) {
    const f = zip.file(part);
    if (!f) continue;

    let xml = f.asText();

    xml = xml.replace(/<w:proofErr\b[^\/]*\/>/g, '');
    xml = xml.replace(/\{%\s*photo\s*\}/g, '{photo}');
    xml = xml.replace(/\{\{\s*([^}]+?)\s*\}\}/g, '{$1}');

    // fix Word-created wrong placeholder in loops: "{$.}" -> "{.}"
    xml = xml.replace(/\{\s*\$\s*\.\s*\}/g, '{.}');

    xml = repairSplitTagsInTextRuns(xml);

    zip.file(part, xml);
  }
}

function repairSplitTagsInTextRuns(xml: string): string {
  const re = /<w:t[^>]*>[\s\S]*?<\/w:t>/g;

  const nodes: Array<{ start: number; end: number; whole: string; text: string }> = [];
  let match: RegExpExecArray | null;

  while ((match = re.exec(xml)) !== null) {
    const whole = match[0];
    const start = match.index || 0;
    const end = start + whole.length;
    const text = whole.replace(/^<w:t[^>]*>/, '').replace(/<\/w:t>$/, '');
    nodes.push({ start, end, whole, text });
  }

  if (!nodes.length) return xml;

  const texts = nodes.map(n => n.text);

  for (let i = 0; i < texts.length; i++) {
    const text = texts[i];
    const openIdx = text.indexOf('{');
    if (openIdx < 0) continue;
    if (text.indexOf('}', openIdx + 1) >= 0) continue;

    let merged = text;
    for (let j = i + 1; j < texts.length; j++) {
      const following = texts[j];
      const closeIdx = following.indexOf('}');
      if (closeIdx < 0) {
        merged += following;
        texts[j] = '';
        continue;
      }
      merged += following.slice(0, closeIdx + 1);
      texts[j] = following.slice(closeIdx + 1);
      break;
    }
    texts[i] = merged;
  }

  let output = '';
  let cursor = 0;
  for (let i = 0; i < nodes.length; i++) {
    const node = nodes[i];
    output += xml.slice(cursor, node.start);
    output += node.whole.replace(node.text, texts[i]);
    cursor = node.end;
  }
  output += xml.slice(cursor);
  return output;
}

function highlightAndReplaceInZip(zip: any, token: string, replacement: string) {
  const xmlParts = Object.keys(zip.files).filter(p =>
    /^word\/(document|header\d+|footer\d+)\.xml$/i.test(p)
  );
  const tokenRe = new RegExp(token.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'g');

  for (const part of xmlParts) {
    const f = zip.file(part);
    if (!f) continue;

    let xml = f.asText();
    if (xml.indexOf(token) < 0) continue;

    xml = xml.replace(/<w:r[\s\S]*?<\/w:r>/g, (runXml: string) => {
      if (runXml.indexOf(token) < 0) return runXml;

      let updated = runXml.replace(tokenRe, replacement);
      if (/<w:highlight\b/i.test(updated)) return updated;

      if (/<w:rPr[\s>]/i.test(updated)) {
        return updated.replace(/<w:rPr[^>]*>/i, (m0: string) => `${m0}<w:highlight w:val="yellow"/>`);
      }
      return updated.replace(/<w:r([^>]*)>/i, (_m0: string, attrs: string) =>
        `<w:r${attrs}><w:rPr><w:highlight w:val="yellow"/></w:rPr>`
      );
    });

    zip.file(part, xml);
  }
}

function highlightLiteralInZip(zip: any, word: string) {
  const xmlParts = Object.keys(zip.files).filter(p =>
    /^word\/(document|header\d+|footer\d+)\.xml$/i.test(p)
  );

  for (const part of xmlParts) {
    const f = zip.file(part);
    if (!f) continue;

    let xml = f.asText();
    if (xml.toLowerCase().indexOf(word.toLowerCase()) < 0) continue;

    xml = xml.replace(/<w:r[\s\S]*?<\/w:r>/g, (runXml: string) => {
      if (runXml.toLowerCase().indexOf(word.toLowerCase()) < 0) return runXml;
      if (/<w:highlight\b/i.test(runXml)) return runXml;

      if (/<w:rPr[\s>]/i.test(runXml)) {
        return runXml.replace(/<w:rPr[^>]*>/i, (m0: string) => `${m0}<w:highlight w:val="yellow"/>`);
      }
      return runXml.replace(/<w:r([^>]*)>/i, (_m0: string, attrs: string) =>
        `<w:r${attrs}><w:rPr><w:highlight w:val="yellow"/></w:rPr>`
      );
    });

    zip.file(part, xml);
  }
}
