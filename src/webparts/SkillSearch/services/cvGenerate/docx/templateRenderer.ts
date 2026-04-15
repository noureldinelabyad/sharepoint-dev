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
  
  // Copy photo into the first embedded image in the body
  if (profile.photoBytes && profile.photoBytes.length) {
    try {
      await replaceFirstBodyImageWithSourcePhoto(doc.getZip(), profile.photoBytes, profile.photoExt || '');
    } catch (e) {
      console.warn('Photo copy failed (ignored):', e);
    }
  }
  
  finalHighlightBitteAnpassen(doc.getZip());

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

function finalHighlightBitteAnpassen(zip: any) {
  const TOKEN = (MISSING_TOKEN || '').trim();
  if (!TOKEN) return;

  const xmlParts = Object.keys(zip.files).filter(p =>
    /^word\/(document|header\d+|footer\d+)\.xml$/i.test(p)
  );

  for (const part of xmlParts) {
    const file = zip.file(part);
    if (!file) continue;

    let xml = file.asText();

    // Work per paragraph, but DO NOT rebuild the paragraph wrapper.
    xml = xml.replace(/<w:p[\s\S]*?<\/w:p>/g, (pXml: string) => {
      const runRe = /<w:r[\s\S]*?<\/w:r>/g;

      // Collect runs with positions
      const runs: Array<{ start: number; end: number; xml: string; text: string }> = [];
      let m: RegExpExecArray | null;

      while ((m = runRe.exec(pXml)) !== null) {
        const rXml = m[0];
        const start = m.index;
        const end = start + rXml.length;

        const texts = [...rXml.matchAll(/<w:t[^>]*>([\s\S]*?)<\/w:t>/g)].map(x => x[1] ?? '');
        const text = texts.join('');
        runs.push({ start, end, xml: rXml, text });
      }

      if (!runs.length) return pXml;

      const paragraphText = runs.map(r => r.text).join('');
      if (!paragraphText.includes(TOKEN)) {
        // Might still be split — but if it's not even in the joined paragraph text, skip
        return pXml;
      }

      // Find all occurrences in paragraphText
      const occurrences: Array<{ start: number; end: number }> = [];
      let from = 0;
      while (true) {
        const idx = paragraphText.indexOf(TOKEN, from);
        if (idx < 0) break;
        occurrences.push({ start: idx, end: idx + TOKEN.length });
        from = idx + 1;
      }
      if (!occurrences.length) return pXml;

      // Rebuild the paragraph by slicing around each run (preserves non-run nodes!)
      let out = '';
      let cursor = 0;

      // Track where each run starts in the paragraphText (global offsets)
      let globalPos = 0;

      for (const r of runs) {
        out += pXml.slice(cursor, r.start);

        const rStart = globalPos;
        const rEnd = globalPos + r.text.length;

        // Runs with no text -> keep as-is
        if (!r.text.length) {
          out += r.xml;
          cursor = r.end;
          globalPos = rEnd;
          continue;
        }

        // Determine highlight segments for THIS run based on overlaps
        // We'll split this run's text into chunks and clone the run for each chunk.
        const segments: Array<{ text: string; highlight: boolean }> = [];

        let localFrom = 0;
        let localGlobal = rStart;

        // Compute all overlaps with occurrences
        const overlaps: Array<{ a: number; b: number }> = [];
        for (const occ of occurrences) {
          if (occ.end <= rStart || occ.start >= rEnd) continue;
          overlaps.push({
            a: Math.max(occ.start, rStart),
            b: Math.min(occ.end, rEnd)
          });
        }

        // No overlap -> keep run unchanged
        if (!overlaps.length) {
          out += r.xml;
          cursor = r.end;
          globalPos = rEnd;
          continue;
        }

        // Sort overlaps and build segments
        overlaps.sort((x, y) => x.a - y.a);

        for (const ov of overlaps) {
          const beforeLen = ov.a - localGlobal;
          const matchLen = ov.b - ov.a;

          if (beforeLen > 0) {
            segments.push({ text: r.text.slice(localFrom, localFrom + beforeLen), highlight: false });
            localFrom += beforeLen;
            localGlobal += beforeLen;
          }

          if (matchLen > 0) {
            segments.push({ text: r.text.slice(localFrom, localFrom + matchLen), highlight: true });
            localFrom += matchLen;
            localGlobal += matchLen;
          }
        }

        // Tail
        if (localFrom < r.text.length) {
          segments.push({ text: r.text.slice(localFrom), highlight: false });
        }

        // Emit cloned runs for segments
        for (const seg of segments) {
          out += cloneRunWithTextSafe(r.xml, seg.text, seg.highlight);
        }

        cursor = r.end;
        globalPos = rEnd;
      }

      out += pXml.slice(cursor);
      return out;
    });

    zip.file(part, xml);
  }
}

function cloneRunWithTextSafe(originalRunXml: string, newText: string, highlight: boolean): string {
  // Only touch "plain text runs" (same idea as you had)
  const isPureTextRun =
    /<w:t[\s>]/i.test(originalRunXml) &&
    !/<w:numPr|<w:fldChar|<w:instrText|<w:sym|<w:object|<w:drawing/i.test(originalRunXml);

  if (!isPureTextRun) return originalRunXml;

  let runXml = originalRunXml;

  // Remove all existing w:t nodes
  runXml = runXml.replace(/<w:t[^>]*>[\s\S]*?<\/w:t>/gi, '');

  // Ensure rPr exists
  if (!/<w:rPr[\s>]/i.test(runXml)) {
    runXml = runXml.replace(/<w:r([^>]*)>/i, `<w:r$1><w:rPr></w:rPr>`);
  }

  // Remove existing highlight tags (both self-closing and expanded forms)
  runXml = runXml.replace(/<w:highlight\b[^>]*\/>/gi, '');
  runXml = runXml.replace(/<w:highlight\b[^>]*>[\s\S]*?<\/w:highlight>/gi, '');

  const normalizedText = (newText ?? '').replace(/\u00A0/g, ' ');
  const escaped = escapeXml(normalizedText);

  // Keep preserve if original had it
  const hasPreserve = /<w:t[^>]*xml:space="preserve"/i.test(originalRunXml) || /^\s|\s$/.test(normalizedText);
  const tOpen = hasPreserve ? `<w:t xml:space="preserve">` : `<w:t>`;

  // Add highlight if requested
  if (highlight) {
    runXml = runXml.replace(/<w:rPr[^>]*>/i, m => `${m}<w:highlight w:val="yellow"/>`);
  }

  // Insert the new w:t before the FIRST closing </w:r> (robust vs whitespace)
  runXml = runXml.replace(/<\/w:r\s*>/i, `${tOpen}${escaped}</w:t></w:r>`);

  return runXml;
}

function escapeXml(text: string): string {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}
