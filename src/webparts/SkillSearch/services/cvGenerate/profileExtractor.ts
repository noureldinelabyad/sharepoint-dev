/* eslint-disable @typescript-eslint/no-explicit-any */
import { extractFirstPhotoFromDocx } from './docx/photoHandling';
import { computeBerufserfahrungFromProjects } from './parsers/experienceParser';
import { extractLanguages } from './parsers/languageParser';
import { parseProjectsCopyPaste } from './parsers/projectParser';
import { extractSkills } from './parsers/skillParser';
import { grabBetween } from './textExtraction';
import { ProfileData } from './types';

export async function extractProfileDataFromDocx(buffer: ArrayBuffer): Promise<ProfileData> {
  // Extract photo bytes from the *source* docx (if any)
  const { photoBytes, photoExt } = extractFirstPhotoFromDocx(buffer);

  const mammoth: any = await import('mammoth/mammoth.browser');
  const { value: html } = await mammoth.convertToHtml({ arrayBuffer: buffer });
  const htmlDoc = new DOMParser().parseFromString(html, 'text/html');

  const paras = Array.from(htmlDoc.getElementsByTagName('p')) as HTMLParagraphElement[];
  const getText = (el?: Element | null) => (el && el.textContent ? el.textContent : '').trim();
  const textLines = paras.map(p => getText(p)).filter(Boolean);

  // 0) collect tables into key-value pairs (top metadata)
  const tablePairs: Record<string, string> = {};
  const allTables = Array.from(htmlDoc.getElementsByTagName('table')) as HTMLTableElement[];
  for (const t of allTables) {
    for (const tr of Array.from(t.rows) as HTMLTableRowElement[]) {
      const rawCells = Array.from(tr.cells)
        .map(td => getText(td).replace(/\s+/g, ' ').trim())
        .filter(c => c !== '');
      if (rawCells.length < 2) continue;

      const uniq: string[] = [];
      for (const c of rawCells) if (!uniq.includes(c)) uniq.push(c);

      const key = (uniq[0] || '').replace(/:$/, '').toLowerCase();
      const value = uniq[uniq.length - 1] || '';
      if (key && value && key !== value) tablePairs[key] = value;
    }
  }

  const firstName = tablePairs['vorname'] || '';
  const lastName  = tablePairs['name'] || '';
  const birthYear = tablePairs['geburtsjahr'] || '';
  const availableFrom =
    tablePairs['verfügbar ab'] ||
    tablePairs['verfuegbar ab'] ||
    tablePairs['verfugbar ab'] ||
    tablePairs['verfÇ¬gbar ab'] ||
    '';
  const education = tablePairs['ausbildung'] || '';

  const name = firstName && lastName ? `${firstName} ${lastName}` : findAfterLabel(/Name/i);
  const role = findAfterLabel(/Rolle|Position/i);
  const team = findAfterLabel(/Team|Abteilung/i);
  const email = (() => {
    const a = Array.from(htmlDoc.getElementsByTagName('a')) as HTMLAnchorElement[];
    const href = a.map(x => x.getAttribute('href') || '').find(h => /^mailto:/i.test(h));
    return href ? href.replace(/^mailto:/i, '') : '';
  })();

  // summary
  const rawSummary = grabBetween(
    /Zu meiner Person|Kurzprofil|Profil/i,
    /Kompetenzen|Skills|Kenntnisse|Zeitraum|Projekte|Tätigkeitsbeschreibung/i,
    paras
  );
  const summary = rawSummary
    .split(/\n+/)
    .filter(l => !/^(name|vorname|geburtsjahr|nationalität|ausbildung|verfügbar ab)\s*:?\s*$/i.test(l.trim()))
    .join('\n')
    .trim();

  const { flatSkills, skillGroups } = extractSkills(htmlDoc);

  // IMPORTANT: projects parsed from 2-column table, right column split into company/headline/desc/bullets
  const projects = parseProjectsCopyPaste(htmlDoc);

  const languages = extractLanguages(allTables, textLines);
  const berufserfahrung = computeBerufserfahrungFromProjects(projects);

  return {
    name: name || '',
    role: role || '',
    team: team || '',
    email,
    skills: flatSkills,
    skillGroups,
    summary,
    projects,
    firstName, lastName, birthYear, availableFrom,
    einsatzAls: role || '',
    einsatzIn: '',
    languages,
    branchen: [],
    qualifikationen: [],
    berufserfahrung,
    education,
    profilnummer: '',
    photoBytes,
    photoExt
  };

  function findAfterLabel(re: RegExp): string {
    const p = paras.find(x => re.test(getText(x)));
    const next = p ? (p.nextElementSibling as HTMLElement | null) : null;
    return getText(next) || getText(p);
  }
}
