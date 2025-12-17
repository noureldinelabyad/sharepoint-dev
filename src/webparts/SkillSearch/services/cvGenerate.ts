/* eslint-disable @typescript-eslint/no-explicit-any */
import { SPHttpClient } from '@microsoft/sp-http';
import PizZip from 'pizzip';
import Docxtemplater from 'docxtemplater';
import * as DOCX from 'docx';
import { BERATERPROFIL_SITE, LIB_INTERNAL_NAME } from './profileConstants';

// ---------- Types ----------
export interface ProjectItem {
  from?: string;
  to?: string;
  period?: string;

  // structured (used by the template)
  company?: string;
  headline?: string;
  description?: string;
  responsibilitiesTitle?: string;
  bullets?: string[];

  // legacy fields (kept for compatibility with older templates)
  title?: string;
  rolle: string;
  kunde: string;
  taetigkeiten: string[];
  technologien: string[];
}

export interface SkillGroup { category: string; items: string[]; }

export interface ProfileData {
  name: string;
  role: string;
  team: string;
  email?: string;
  skills: string[];
  summary: string;
  projects: ProjectItem[];
  skillGroups?: SkillGroup[];

  firstName?: string;
  lastName?: string;
  birthYear?: string;
  availableFrom?: string;
  einsatzAls?: string;
  einsatzIn?: string;
  languages?: string[];
  branchen?: string[];
  qualifikationen?: string[];
  berufserfahrung?: string;
  education?: string;
  profilnummer?: string;

  // photo from source docx (optional)
  photoBytes?: Uint8Array;
  photoExt?: string; // "png" | "jpg" | "jpeg"
}

// ---------- Constants ----------
const MISSING_TOKEN = '__ANPASSEN__';

// ---------- Binary download ----------
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

// ---------- Parse Beraterprofil ----------
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

function grabBetween(start: RegExp, end: RegExp, paras: HTMLParagraphElement[]): string {
  const texts = paras.map(p => (p.textContent || '').trim());
  const s = texts.findIndex(t => start.test(t));
  if (s < 0) return '';
  const eRel = texts.slice(s + 1).findIndex(t => end.test(t));
  const slice = texts.slice(s + 1, eRel > -1 ? s + 1 + eRel : undefined);
  return slice.join('\n').trim();
}

// ---------- Skills ----------
function extractSkills(htmlDoc: Document): { flatSkills: string[]; skillGroups: SkillGroup[] } {
  const tables = Array.from(htmlDoc.getElementsByTagName('table')) as HTMLTableElement[];

  for (const t of tables) {
    const header = Array.from(t.rows && t.rows[0] ? t.rows[0].cells : [])
      .map(c => (c.textContent || '').trim().toLowerCase());

    if (!header.some(h => h.includes('beschreibung'))) continue;

    const groups: SkillGroup[] = [];
    for (const tr of Array.from(t.rows).slice(1)) {
      const cells = Array.from(tr.cells)
        .map(td => (td.textContent || '').replace(/\s+/g, ' ').trim())
        .filter(Boolean);

      if (cells.length < 2) continue;

      const category = cells[0];
      const items = cells[1].split(/,\s*/).map(x => x.trim()).filter(Boolean);
      if (category && items.length) groups.push({ category, items });
    }

    const flat = groups.reduce((acc: string[], g: SkillGroup) => {
      acc.push(g.category);
      for (const it of g.items) acc.push(it);
      return acc;
    }, []);

    return { flatSkills: flat, skillGroups: groups };
  }

  return { flatSkills: [], skillGroups: [] };
}

// ---------- Languages ----------
function extractLanguages(tables: HTMLTableElement[], lines: string[]): string[] {
  for (const t of tables) {
    const headerCells = Array.from(t.rows && t.rows[0] ? t.rows[0].cells : [])
      .map(c => (c.textContent || '').trim().toLowerCase());

    const isLangTable = headerCells.some(h => h.includes('sprachen')) && headerCells.some(h => h.includes('niveau'));
    if (!isLangTable) continue;

    const res: string[] = [];
    for (const tr of Array.from(t.rows).slice(1)) {
      const lang = (tr.cells && tr.cells[0] ? tr.cells[0].textContent : '') || '';
      const level = (tr.cells && tr.cells[1] ? tr.cells[1].textContent : '') || '';
      const l = lang.trim();
      const lv = level.trim();
      if (l) res.push(lv ? `${l} (${lv})` : l);
    }
    return uniqueKeepOrder(res);
  }

  const idx = lines.findIndex(l => /^sprachen\b/i.test(l) || /\bsprachen\b/i.test(l));
  if (idx >= 0) {
    const out: string[] = [];
    const firstLine = lines[idx];
    const after = firstLine.replace(/^.*?\bsprachen\b\s*:?\s*/i, '').trim();
    if (after) out.push(...splitToItems(after));

    const stopRe = /(kompetenzen|skills|kenntnisse|projekte|projektreferenzen|tätigkeitsbeschreibung|ausbildung|zertifikat|branchen|qualifikationen)/i;
    for (let i = idx + 1; i < lines.length; i++) {
      const l = (lines[i] || '').trim();
      if (!l) break;
      if (stopRe.test(l)) break;
      out.push(...splitToItems(l));
    }
    return uniqueKeepOrder(out).filter(Boolean);
  }

  return [];
}

function splitToItems(s: string): string[] {
  return s.split(/[,;•\n]/).map(x => x.trim()).filter(Boolean);
}

function uniqueKeepOrder(arr: string[]): string[] {
  const seen = new Set<string>();
  const out: string[] = [];
  for (const v of arr) {
    const k = v.trim();
    if (!k) continue;
    if (seen.has(k)) continue;
    seen.add(k);
    out.push(k);
  }
  return out;
}

// ---------- Projects: 2-column copy/paste parsing ----------
function parseProjectsCopyPaste(htmlDoc: Document): ProjectItem[] {
  const tables = Array.from(htmlDoc.getElementsByTagName('table')) as HTMLTableElement[];

  const periodRe =
    /^\s*(seit\s*)?(\d{1,2}\s*[./&-]\s*\d{4}|\d{4})(\s*[–-]\s*(heute|\d{1,2}\s*[./&-]\s*\d{4}|\d{4}))?/i;

  for (const t of tables) {
    const rows = Array.from(t.rows) as HTMLTableRowElement[];
    if (!rows.length) continue;

    const out: ProjectItem[] = [];

    for (const tr of rows) {
      const cells = Array.from(tr.cells);
      if (cells.length < 2) continue;

      const left = normalizeWhitespace(extractCellTextPreservingOrder(cells[0]));
      const rightRaw = normalizeWhitespace(extractCellTextPreservingOrder(cells[cells.length - 1]));

      // skip obvious header row
      if (/zeitraum/i.test(left) && /tätigkeitsbeschreibung|projekthistorie|projekte/i.test(rightRaw)) continue;
      if (!left || !rightRaw) continue;
      if (!periodRe.test(left)) continue;

      const parsed = splitProjectRightColumn(rightRaw);

      out.push({
        period: left,

        company: parsed.company,
        headline: parsed.headline,
        description: parsed.description,
        responsibilitiesTitle: parsed.responsibilitiesTitle,
        bullets: parsed.bullets,

        // keep compat fields (empty, but required by your interface)
        rolle: '',
        kunde: '',
        taetigkeiten: [],
        technologien: []
      });
    }

    if (out.length) return out;
  }

  return [];
}

function splitProjectRightColumn(right: string): {
  company: string;
  headline: string;
  description: string;
  responsibilitiesTitle: string;
  bullets: string[];
} {
  const lines = right.split('\n').map(l => l.trim()).filter(Boolean);

  let company = lines[0] || '';
  let headline = lines[1] || '';

  // Find “Verantwortlichkeiten:” even if it’s like “Verantwortlichkeiten :”
  const respIdx = findIndexRegex(lines, /^verantwortlichkeiten\b\s*:?\s*$/i);

  // Find first bullet line index (even without responsibilities title)
  const firstBulletIdx = findIndexRegex(lines, /^[•\-\u2022]\s+/);

  // Decide where description stops
  let descEnd = lines.length;
  let responsibilitiesTitle = 'Verantwortlichkeiten:';

  if (respIdx >= 0) {
    responsibilitiesTitle = (lines[respIdx] || 'Verantwortlichkeiten:').replace(/\s*:\s*$/, ':');
    descEnd = respIdx;
  } else if (firstBulletIdx >= 0) {
    // no explicit label, but bullets exist: description ends before first bullet
    descEnd = firstBulletIdx;
  }

  // Description: between headline and responsibilities/bullets
  const descLines: string[] = [];
  for (let i = 2; i < descEnd; i++) descLines.push(lines[i]);

  // Bullets: after responsibilities label, OR from first bullet onward
  const bullets: string[] = [];
  let i = respIdx >= 0 ? respIdx + 1 : (firstBulletIdx >= 0 ? firstBulletIdx : lines.length);

  for (; i < lines.length; i++) {
    const l = lines[i];

    const m = l.match(/^[•\-\u2022]\s*(.+)$/);
    if (m) {
      bullets.push(m[1].trim());
      continue;
    }

    // continuation line: append to previous bullet
    if (/^[a-zäöü]/i.test(l) && bullets.length) {
      bullets[bullets.length - 1] = `${bullets[bullets.length - 1]} ${l}`.trim();
    }
  }

  // If company/headline are missing because the source had fewer lines, degrade gracefully
  if (!headline && descLines.length) {
    headline = descLines.shift() || '';
  }

  return {
    company: company.trim(),
    headline: headline.trim(),
    description: descLines.join('\n').trim(),
    responsibilitiesTitle,
    bullets: bullets.filter(Boolean)
  };
}

function findIndexRegex(arr: string[], re: RegExp): number {
  for (let i = 0; i < arr.length; i++) if (re.test(arr[i])) return i;
  return -1;
}

// IMPORTANT: preserve DOM order (so description + bullets don’t get scrambled)
function extractCellTextPreservingOrder(cell: HTMLTableCellElement): string {
  const out: string[] = [];

  const push = (txt: string) => {
    const t = normalizeWhitespace(txt);
    if (!t) return;
    if (out.length && out[out.length - 1] === t) return;
    out.push(t);
  };

  const walk = (node: Node) => {
    if (node.nodeType === Node.TEXT_NODE) {
      push(node.textContent || '');
      return;
    }

    if (node.nodeType !== Node.ELEMENT_NODE) return;
    const el = node as HTMLElement;
    const tag = (el.tagName || '').toLowerCase();

    if (tag === 'br') {
      push('\n');
      return;
    }

    if (tag === 'li') {
      const txt = (el.textContent || '').trim();
      if (txt) push(`• ${txt}`);
      return;
    }

    if (tag === 'ul' || tag === 'ol') {
      // walk only direct li children in order
      const lis = Array.from(el.children).filter(c => (c as HTMLElement).tagName?.toLowerCase() === 'li');
      for (const li of lis) walk(li);
      return;
    }

    if (tag === 'p' || tag === 'div') {
      // walk children then force a line break boundary
      const kids = Array.from(el.childNodes);
      for (const k of kids) walk(k);
      // ensure paragraph boundary
      push('\n');
      return;
    }

    // generic: keep order
    const kids = Array.from(el.childNodes);
    for (const k of kids) walk(k);
  };

  const kids = Array.from(cell.childNodes);
  if (kids.length) {
    for (const k of kids) walk(k);
  } else {
    push(cell.textContent || '');
  }

  return out
    .join('\n')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}

function normalizeWhitespace(s: string): string {
  return (s || '')
    .replace(/\u00A0/g, ' ')
    .replace(/[ \t]+\n/g, '\n')
    .replace(/\n[ \t]+/g, '\n')
    .replace(/[ \t]{2,}/g, ' ')
    .trim();
}

// ---------- Berufserfahrung ----------
function computeBerufserfahrungFromProjects(projects: ProjectItem[]): string {
  const starts: Date[] = [];
  for (const p of projects) {
    const dt = parseStartDateFromPeriod(p.period || '');
    if (dt) starts.push(dt);
  }
  if (!starts.length) return '';

  starts.sort((a, b) => a.getTime() - b.getTime());
  const start = starts[0];
  const now = new Date();
  const months = (now.getFullYear() - start.getFullYear()) * 12 + (now.getMonth() - start.getMonth());
  const years = Math.max(0, Math.floor(months / 12));

  if (years <= 0) return '0 Jahre';
  if (years === 1) return '1 Jahr';
  return `${years} Jahre`;
}

function parseStartDateFromPeriod(period: string): Date | null {
  const p = (period || '').replace(/&/g, '/').trim().toLowerCase();
  if (!p) return null;

  const seit = p.match(/seit\s*(\d{1,2}\s*[./-]\s*\d{4}|\d{4})/i);
  if (seit) return parseMonthYearOrYear(seit[1]);

  const mmyyyy = p.match(/(\d{1,2})\s*[./-]\s*(\d{4})/);
  if (mmyyyy) {
    const mm = clampInt(parseInt(mmyyyy[1], 10), 1, 12);
    const yyyy = parseInt(mmyyyy[2], 10);
    if (yyyy >= 1900 && yyyy <= 2100) return new Date(yyyy, mm - 1, 1);
  }

  const yyyyOnly = p.match(/(\d{4})/);
  if (yyyyOnly) {
    const yyyy = parseInt(yyyyOnly[1], 10);
    if (yyyy >= 1900 && yyyy <= 2100) return new Date(yyyy, 0, 1);
  }

  return null;
}

function parseMonthYearOrYear(s: string): Date | null {
  const t = (s || '').replace(/&/g, '/').trim();
  const m = t.match(/(\d{1,2})\s*[./-]\s*(\d{4})/);
  if (m) {
    const mm = clampInt(parseInt(m[1], 10), 1, 12);
    const yyyy = parseInt(m[2], 10);
    if (yyyy >= 1900 && yyyy <= 2100) return new Date(yyyy, mm - 1, 1);
  }
  const y = t.match(/(\d{4})/);
  if (y) {
    const yyyy = parseInt(y[1], 10);
    if (yyyy >= 1900 && yyyy <= 2100) return new Date(yyyy, 0, 1);
  }
  return null;
}

function clampInt(v: number, min: number, max: number): number {
  return Math.max(min, Math.min(max, v));
}

// ---------- Photo extraction from source docx ----------
function extractFirstPhotoFromDocx(buffer: ArrayBuffer): { photoBytes?: Uint8Array; photoExt?: string } {
  try {
    const zip = new PizZip(buffer as any);
    const files = Object.keys(zip.files || {}).filter(p => /^word\/media\//i.test(p) && !zip.files[p].dir);
    if (!files.length) return {};

    const ordered = files.filter(f => /\.(png|jpe?g)$/i.test(f)).sort((a, b) => a.localeCompare(b));
    const pick = ordered[0] || files[0];

    const u8 = zip.file(pick)?.asUint8Array?.();
    if (!u8 || !u8.length) return {};

    const extMatch = pick.match(/\.(png|jpe?g)$/i);
    const photoExt = extMatch ? extMatch[1].toLowerCase() : 'png';
    return { photoBytes: u8, photoExt };
  } catch {
    return {};
  }
}

// ---------- Docxtemplater path ----------
export async function fillDataportTemplate(
  spHttp: SPHttpClient,
  templateUrlOrServerRel: string,
  d: ProfileData
): Promise<Blob> {
  const buf = await downloadArrayBuffer(spHttp, templateUrlOrServerRel);
  const zip = new PizZip(buf);

  sanitizeDocxtemplaterZip(zip);

  const doc = new Docxtemplater(zip, {
    paragraphLoop: true,
    linebreaks: true,
    nullGetter: () => MISSING_TOKEN
  });

  const v = (s?: string) => (s && s.trim() ? s.trim() : MISSING_TOKEN);
  const arr = (a?: string[]) => (a && a.length ? a : [MISSING_TOKEN]);

  const beruf = d.berufserfahrung && d.berufserfahrung.trim()
    ? d.berufserfahrung.trim()
    : computeBerufserfahrungFromProjects(d.projects || []);

  const projects = (d.projects || []).length
    ? d.projects.map(p => ({
        period: v(p.period),
        company: v(p.company),
        headline: v(p.headline),
        description: v(p.description),
        responsibilitiesTitle: v(p.responsibilitiesTitle || 'Verantwortlichkeiten:'),
        bullets: (p.bullets && p.bullets.length)
          ? p.bullets
              .map(b => (b ?? '').toString().trim())
              .filter(b => b.length > 0)
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
    profilnummer: v(d.profilnummer),
    photo: '',

    firstName: v(d.firstName || (d.name?.split(' ').slice(0, -1).join(' ') || d.name || '')),
    lastName: v(d.lastName || (d.name?.split(' ').slice(-1)[0] || '')),
    birthYear: v(d.birthYear),
    availableFrom: v(d.availableFrom),
    einsatzAls: v(d.einsatzAls || d.role || ''),
    einsatzIn: v(d.einsatzIn),

    languages: arr(d.languages),
    languagesText: arr(d.languages).join('\n'),

    branchen: arr(d.branchen),
    branchenText: arr(d.branchen).join(', '),

    qualifikationen: arr(d.qualifikationen),
    qualifikationenText: arr(d.qualifikationen).join(', '),

    education: v(d.education),
    berufserfahrung: v(beruf),

    name: v(d.name),
    role: v(d.role),
    team: v(d.team),
    email: v(d.email),
    summary: v(d.summary),

    skills: arr(d.skills),
    skillGroups: (d.skillGroups && d.skillGroups.length) ? d.skillGroups : [{ category: MISSING_TOKEN, items: [MISSING_TOKEN] }],

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

  // Ensure both generated tokens and literal "anpassen" are yellow
  highlightAndReplaceInZip(doc.getZip(), MISSING_TOKEN, 'anpassen');
  highlightLiteralInZip(doc.getZip(), 'anpassen');

  // Copy photo into the first embedded image in the body
  if (d.photoBytes && d.photoBytes.length) {
    try {
      await replaceFirstBodyImageWithSourcePhoto(doc.getZip(), d.photoBytes, d.photoExt || '');
    } catch (e) {
      console.warn('Photo copy failed (ignored):', e);
    }
  }

  return doc.getZip().generate({ type: 'blob' });
}

function sanitizeDocxtemplaterZip(zip: any) {
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
  let m: RegExpExecArray | null;

  while ((m = re.exec(xml)) !== null) {
    const whole = m[0];
    const start = m.index || 0;
    const end = start + whole.length;
    const text = whole.replace(/^<w:t[^>]*>/, '').replace(/<\/w:t>$/, '');
    nodes.push({ start, end, whole, text });
  }

  if (!nodes.length) return xml;

  const texts = nodes.map(n => n.text);

  for (let i = 0; i < texts.length; i++) {
    const t = texts[i];
    const openIdx = t.indexOf('{');
    if (openIdx < 0) continue;
    if (t.indexOf('}', openIdx + 1) >= 0) continue;

    let merged = t;
    for (let j = i + 1; j < texts.length; j++) {
      const tj = texts[j];
      const closeIdx = tj.indexOf('}');
      if (closeIdx < 0) {
        merged += tj;
        texts[j] = '';
        continue;
      }
      merged += tj.slice(0, closeIdx + 1);
      texts[j] = tj.slice(closeIdx + 1);
      break;
    }
    texts[i] = merged;
  }

  let out = '';
  let cursor = 0;
  for (let i = 0; i < nodes.length; i++) {
    const n = nodes[i];
    out += xml.slice(cursor, n.start);
    out += n.whole.replace(n.text, texts[i]);
    cursor = n.end;
  }
  out += xml.slice(cursor);
  return out;
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

async function replaceFirstBodyImageWithSourcePhoto(zip: any, photoBytes: Uint8Array, photoExt: string) {
  const docXmlFile = zip.file('word/document.xml');
  const relsFile = zip.file('word/_rels/document.xml.rels');
  if (!docXmlFile || !relsFile) return;

  const docXml = docXmlFile.asText();
  const relsXml = relsFile.asText();

  const ridMatch = docXml.match(/r:embed="(rId\d+)"/);
  if (!ridMatch) return;

  const rid = ridMatch[1];
  const relRe = new RegExp(`<Relationship[^>]+Id="${rid}"[^>]+Target="([^"]+)"[^>]*/>`, 'i');
  const relMatch = relsXml.match(relRe);
  if (!relMatch) return;

  const target = relMatch[1]; // e.g. "media/image2.png"
  const targetPath = target.startsWith('word/') ? target : `word/${target.replace(/^\.?\//, '')}`;
  const oldExtMatch = targetPath.match(/\.(png|jpe?g)$/i);
  const oldExt = oldExtMatch ? oldExtMatch[1].toLowerCase() : 'png';

  let finalBytes = photoBytes;
  if (oldExt !== (photoExt || '').toLowerCase()) {
    const want = (oldExt === 'jpg' || oldExt === 'jpeg') ? 'image/jpeg' : 'image/png';
    finalBytes = await convertImageBytes(photoBytes, want);
  }

  zip.file(targetPath, finalBytes);
}

async function convertImageBytes(bytes: Uint8Array, mime: string): Promise<Uint8Array> {
  const blob = new Blob([bytes], { type: 'application/octet-stream' });
  const url = URL.createObjectURL(blob);

  try {
    const img = await loadImage(url);
    const canvas = document.createElement('canvas');
    canvas.width = img.width || 300;
    canvas.height = img.height || 300;

    const ctx = canvas.getContext('2d');
    if (!ctx) return bytes;

    ctx.drawImage(img, 0, 0);

    const outBlob: Blob = await new Promise((resolve) => {
      canvas.toBlob((b) => resolve(b || blob), mime, 0.92);
    });

    const ab = await outBlob.arrayBuffer();
    return new Uint8Array(ab);
  } finally {
    URL.revokeObjectURL(url);
  }
}

function loadImage(url: string): Promise<HTMLImageElement> {
  return new Promise((resolve, reject) => {
    const img = new Image();
    img.onload = () => resolve(img);
    img.onerror = (e) => reject(e);
    img.src = url;
  });
}

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

// ---------- Fallback ----------
export async function buildDataportDocx(data: ProfileData): Promise<Blob> {
  const heading = new DOCX.Paragraph({
    text: 'Datenblatt und Einschätzung zur Erbringung der Arbeitnehmerüberlassung',
    heading: DOCX.HeadingLevel.HEADING_1,
    spacing: { after: 200 }
  });

  const summaryTable = new DOCX.Table({
    width: { size: 100, type: DOCX.WidthType.PERCENTAGE },
    rows: [
      row2('Name', data.name || ''),
      row2('Rolle/Position', data.role || ''),
      row2('Team/Abteilung', data.team || ''),
      row2('E-Mail', data.email || '')
    ]
  });

  const doc = new DOCX.Document({ sections: [{ children: [heading, summaryTable] }] });
  return DOCX.Packer.toBlob(doc);

  function row2(label: string, value: string): DOCX.TableRow {
    return new DOCX.TableRow({
      children: [
        new DOCX.TableCell({ width: { size: 30, type: DOCX.WidthType.PERCENTAGE }, children: [paraBold(label)] }),
        new DOCX.TableCell({ width: { size: 70, type: DOCX.WidthType.PERCENTAGE }, children: [new DOCX.Paragraph({ text: value || 'anpassen' })] })
      ]
    });
  }
  function paraBold(text: string) { return new DOCX.Paragraph({ children: [new DOCX.TextRun({ text, bold: true })] }); }
}

// ---------- Smart wrapper ----------
export async function tryGenerateDataportDoc(
  spHttp: SPHttpClient,
  maybeTemplateUrlOrServerRel: string | null,
  data: ProfileData
): Promise<Blob> {
  if (maybeTemplateUrlOrServerRel) {
    try {
      return await fillDataportTemplate(spHttp, maybeTemplateUrlOrServerRel, data);
    } catch (err) {
      console.warn('Falling back to generated Dataport layout because template rendering failed', err);
    }
  }
  return buildDataportDocx(data);
}

export function defaultTemplateUrl(): string {
  return `${BERATERPROFIL_SITE}/${LIB_INTERNAL_NAME}/Dataport CV Vorlage - TAGGED.docx`;
}
