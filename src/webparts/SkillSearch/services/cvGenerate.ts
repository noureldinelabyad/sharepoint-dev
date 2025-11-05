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
  title?: string;
  rolle: string;
  kunde: string;
  taetigkeiten: string[];
  technologien: string[];
}

export interface ProfileData {
  name: string;
  role: string;
  team: string;
  email?: string;
  skills: string[];
  summary: string;
  projects: ProjectItem[];

  // Optional extras your template may use
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
}

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
  const mammoth: any = await import('mammoth/mammoth.browser');
  const { value: html } = await mammoth.convertToHtml({ arrayBuffer: buffer });
  const htmlDoc = new DOMParser().parseFromString(html, 'text/html');

  const paras = Array.from(htmlDoc.getElementsByTagName('p')) as HTMLParagraphElement[];
  const getText = (el?: Element | null) => (el && el.textContent ? el.textContent : '').trim();
  const textLines = paras.map(p => getText(p)).filter(Boolean);

  // 0) collect tables into key-value pairs for top metadata
  const tablePairs: Record<string, string> = {};
  const allTables = Array.from(htmlDoc.getElementsByTagName('table')) as HTMLTableElement[];
  for (const t of allTables) {
    for (const tr of Array.from(t.rows) as HTMLTableRowElement[]) {
      const cells = Array.from(tr.cells) as HTMLTableCellElement[];
      if (cells.length >= 2) {
        const k = getText(cells[0]).toLowerCase();
        const v = getText(cells[1]);
        if (k && v) tablePairs[k] = v;
      }
    }
  }

  const firstName = tablePairs['vorname'] || '';
  const lastName  = tablePairs['name'] || '';
  const birthYear = tablePairs['geburtsjahr'] || '';
  const availableFrom = tablePairs['verfügbar ab'] || tablePairs['verfugbar ab'] || '';
  const education = tablePairs['ausbildung'] || '';

  const name = firstName && lastName ? `${firstName} ${lastName}` : findAfterLabel(/Name/i);
  const role = findAfterLabel(/Rolle|Position/i);
  const team = findAfterLabel(/Team|Abteilung/i);
  const email = (() => {
    const a = Array.from(htmlDoc.getElementsByTagName('a')) as HTMLAnchorElement[];
    const href = a.map(x => x.getAttribute('href') || '').find(h => /^mailto:/i.test(h));
    return href ? href.replace(/^mailto:/i, '') : '';
  })();

  // 1) summary (strip obvious label rows if they leaked in)
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

  // 2) skills: UL after heading; else table after heading; else paragraph block
  const skills = extractSkills(htmlDoc, textLines);

  // 3) projects: table flavour or free-text flavour
  const projects = parseProjects(htmlDoc, textLines);

  return {
    name: name || '',
    role: role || '',
    team: team || '',
    email,
    skills,
    summary,
    projects,
    firstName, lastName, birthYear, availableFrom,
    einsatzAls: role || '',
    einsatzIn: '',
    languages: extractLanguages(allTables),
    branchen: [],
    qualifikationen: [],
    berufserfahrung: '',
    education,
    profilnummer: ''
  };

  function findAfterLabel(re: RegExp): string {
    const p = paras.find(x => re.test(getText(x)));
    const next = p ? (p.nextElementSibling as HTMLElement | null) : null;
    return getText(next) || getText(p);
  }
}

function extractLanguages(tables: HTMLTableElement[]): string[] {
  const res: string[] = [];
  for (const t of tables) {
    const txt = (t.textContent || '');
    if (/Sprachen/i.test(txt) && /Niveau/i.test(txt)) {
      const rows = Array.from(t.rows).slice(1);
      for (const tr of rows) {
        const cells = Array.from(tr.cells);
        if (cells[0]) {
          const lang = (cells[0].textContent || '').trim();
          if (lang) res.push(lang);
        }
      }
    }
  }
  return res;
}

function grabBetween(start: RegExp, end: RegExp, paras: HTMLParagraphElement[]): string {
  const texts = paras.map(p => (p.textContent || '').trim());
  const s = texts.findIndex(t => start.test(t));
  if (s < 0) return '';
  const eRel = texts.slice(s + 1).findIndex(t => end.test(t));
  const slice = texts.slice(s + 1, eRel > -1 ? s + 1 + eRel : undefined);
  return slice.join('\n').trim();
}

function extractSkills(htmlDoc: Document, allLines: string[]): string[] {
  // UL after heading
  const elems = Array.from(htmlDoc.querySelectorAll('p,strong,h1,h2,h3')) as HTMLElement[];
  const probe = elems.find(e => /Kompetenzen|Skills|Kenntnisse/i.test((e.textContent || '').trim()));
  const next = probe ? (probe.nextElementSibling as HTMLElement | null) : null;
  const ul = next && next.tagName === 'UL'
    ? next
    : (probe && probe.parentElement ? (probe.parentElement.querySelector('ul') as HTMLElement | null) : null);
  if (ul) {
    return Array.from(ul.querySelectorAll('li')).map(li => (li.textContent || '').trim()).filter(Boolean);
  }

  // Table after heading -> take all cell values as flat list
  const followingTable = probe
    ? (probe.nextElementSibling?.tagName === 'TABLE'
        ? (probe.nextElementSibling as HTMLTableElement)
        : (probe.parentElement?.querySelector('table') as HTMLTableElement | null))
    : null;
  if (followingTable) {
    const set = new Set<string>();
    for (const tr of Array.from(followingTable.rows)) {
      for (const td of Array.from(tr.cells)) {
        const t = (td.textContent || '').trim();
        if (t) set.add(t);
      }
    }
    return Array.from(set);
  }

  // Paragraph block fallback
  const idx = allLines.findIndex(l => /Kompetenzen|Skills|Kenntnisse/i.test(l));
  if (idx >= 0) {
    const rest = allLines.slice(idx + 1);
    const stop = rest.findIndex(l => /Projekte|Tätigkeitsbeschreibung|Zeitraum|Sprachen/i.test(l));
    const block = stop >= 0 ? rest.slice(0, stop) : rest;
    return block
      .join('\n')
      .split(/[,;\n•]/)
      .map(s => s.trim())
      .filter(Boolean);
  }
  return [];
}

function parseProjects(htmlDoc: Document, lines: string[]): ProjectItem[] {
  // 1) Table flavour
  const tables = Array.from(htmlDoc.getElementsByTagName('table')) as HTMLTableElement[];
  for (const t of tables) {
    const txt = t.textContent || '';
    if (/Kunde|Branche/i.test(txt) && /Tätigkeiten|Taetigkeiten/i.test(txt)) {
      const rows = Array.from(t.rows).slice(1) as HTMLTableRowElement[];
      const arr = rows.map(tr => {
        const tds = Array.from(tr.cells).map(td => (td.textContent || '').trim());
        return {
          kunde: tds[0] || '',
          rolle: tds[1] || '',
          taetigkeiten: splitList(tds[2] || ''),
          technologien: splitList(tds[3] || '')
        } as ProjectItem;
      });
      if (arr.length) return arr;
    }
  }

  // 2) Free-text flavour — capture date + inline title on same line
  const out: ProjectItem[] = [];
  const startIdx = lines.findIndex(l => /Projektreferenzen|Tätigkeitsbeschreibung|Projekthistorie/i.test(l));
  if (startIdx < 0) return out;

  let i = startIdx + 1;
  const dateRe = /^(seit\s*)?(\d{2}[./-]\d{4}|\d{4})(?:\s*[–-]\s*(heute|\d{2}[./-]\d{4}|\d{4}))?(.*)$/i;

  while (i < lines.length) {
    const line = lines[i].trim();
    const m = line.match(dateRe);
    if (!m) { i++; continue; }

    const from = m[2] || '';
    let to = (m[3] || '').toLowerCase();
    if (!to && /^seit/i.test(line)) to = 'heute';

    const inlineTitle = (m[4] || '').trim(); // e.g. "Migration Microsoft 365"
    const period = [from, to].filter(Boolean).join(' – ') || from;

    // next non-empty -> kunde/branch or role/title if formatted that way
    i++;
    while (i < lines.length && !lines[i].trim()) i++;
    let kunde = (lines[i] || '').trim();

    // If the next line is like "Rolle: X" or "Kunde/Branche: Y", capture properly
    let rolle = '';
    let title = inlineTitle;
    const details: string[] = [];
    i++;

    while (i < lines.length) {
      const l = lines[i].trim();
      if (dateRe.test(l)) break; // next project

      if (/^Rolle\s*:/.test(l)) { rolle = l.replace(/^Rolle\s*:/, '').trim(); i++; continue; }
      if (/^Kunde\/?Branche\s*:/.test(l)) { kunde = l.replace(/^Kunde\/?Branche\s*:/, '').trim(); i++; continue; }

      details.push(l);
      i++;
    }

    // Tätigkeiten block detection
    const tasks: string[] = [];
    const idxV = details.findIndex(l => /Verantwortlichkeiten|Tätigkeiten/i.test(l));
    if (idxV > -1) {
      for (const l of details.slice(idxV + 1)) {
        if (!l.trim()) break;
        tasks.push(l.replace(/^[•\-\u2022]\s*/, '').trim());
      }
    }

    // Technologies guess
    const techs: string[] = [];
    const techLine = details.slice().reverse().find(l => /,/.test(l) || /(mit|via|unter Nutzung von)/i.test(l));
    if (techLine) techs.push(...splitList(techLine));

    out.push({ from, to, period, title, rolle, kunde, taetigkeiten: tasks, technologien: techs });
  }
  return out;
}

function splitList(s: string): string[] {
  return s.split(/[,;•\n]/).map(x => x.trim()).filter(Boolean);
}

// ---------- Docxtemplater path ----------
export async function fillDataportTemplate(
  spHttp: SPHttpClient,
  templateUrlOrServerRel: string,
  d: ProfileData
): Promise<Blob> {
  const buf = await downloadArrayBuffer(spHttp, templateUrlOrServerRel);
  const zip = new PizZip(buf);
  const doc = new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true, nullGetter: () => '' });

  const model = {
    profilnummer: d.profilnummer || '',
    lastName: d.lastName || (d.name?.split(' ').slice(-1)[0] || ''),
    firstName: d.firstName || (d.name?.split(' ').slice(0, -1).join(' ') || d.name || ''),
    birthYear: d.birthYear || '',
    availableFrom: d.availableFrom || '',
    einsatzAls: d.einsatzAls || d.role || '',
    einsatzIn: d.einsatzIn || '',
    languages: d.languages || [],
    branchen: d.branchen || [],
    qualifikationen: d.qualifikationen || [],
    berufserfahrung: d.berufserfahrung || '',
    education: d.education || '',

    name: d.name || '',
    role: d.role || '',
    team: d.team || '',
    email: d.email || '',
    summary: d.summary || '',
    skills: d.skills || [],

    projects: (d.projects || []).map(p => ({
      period: p.period || [p.from, p.to].filter(Boolean).join(' – '),
      from: p.from || '',
      to: p.to || '',
      title: p.title || '',
      rolle: p.rolle || '',
      kunde: p.kunde || '',
      taetigkeiten: p.taetigkeiten || [],
      technologien: p.technologien || []
    }))
  };

  doc.setData(model);
  doc.render();
  return doc.getZip().generate({ type: 'blob' });
}

export async function templateHasDocxtemplaterTags(
  spHttp: SPHttpClient,
  templateUrlOrServerRel: string
): Promise<boolean> {
  try {
    const buf = await downloadArrayBuffer(spHttp, templateUrlOrServerRel);
    const zip = new PizZip(buf);
    const xml = zip.file('word/document.xml')?.asText() || '';
    return /\{\{[^}]+\}\}/.test(xml) || /\{#.+?\}/.test(xml);
  } catch {
    return false;
  }
}

// ---------- Fallback: Dataport-like layout with projects table ----------
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

  // Left column (photo placeholder + facts)
  const left: (DOCX.Paragraph | DOCX.Table)[] = [
    new DOCX.Paragraph({ text: '' }),
    new DOCX.Paragraph({ children: [new DOCX.TextRun({ text: 'Foto', italics: true })] }),
    new DOCX.Paragraph({ text: '' }),
    new DOCX.Paragraph({ text: 'Fakten', heading: DOCX.HeadingLevel.HEADING_3 }),
    ...paraIf('Verfügbar ab', data.availableFrom),
  ];
  if (data.languages?.length) {
    left.push(new DOCX.Paragraph({ children: [new DOCX.TextRun({ text: 'Sprachen', bold: true })] }));
    left.push(...bullets(data.languages));
  }

  // Projects table with header row
  const projectRows: DOCX.TableRow[] = [
    new DOCX.TableRow({
      children: [
        new DOCX.TableCell({
          width: { size: 25, type: DOCX.WidthType.PERCENTAGE },
          children: [paraBold('Zeitraum')],
          margins: margins()
        }),
        new DOCX.TableCell({
          width: { size: 75, type: DOCX.WidthType.PERCENTAGE },
          children: [paraBold('Details')],
          margins: margins()
        })
      ]
    }),
    ...(data.projects?.length ? data.projects.map(p => projectRow(p)) : [projectRow({} as ProjectItem)])
  ];
  const projectsTable = new DOCX.Table({ width: { size: 100, type: DOCX.WidthType.PERCENTAGE }, rows: projectRows });

  // Right column (content)
  const right: (DOCX.Paragraph | DOCX.Table)[] = [
    new DOCX.Paragraph({ text: 'Kurzprofil', heading: DOCX.HeadingLevel.HEADING_2 }),
    ...paraMultiline(data.summary || ''),
    new DOCX.Paragraph({ text: '' }),
    new DOCX.Paragraph({ text: 'Kompetenzen', heading: DOCX.HeadingLevel.HEADING_2 }),
    ...(data.skills?.length ? bullets(data.skills) : [new DOCX.Paragraph({ text: '—' })]),
    new DOCX.Paragraph({ text: '' }),
    new DOCX.Paragraph({ text: 'Projekte / Referenzen', heading: DOCX.HeadingLevel.HEADING_2 }),
    projectsTable
  ];

  const grid = new DOCX.Table({
    width: { size: 100, type: DOCX.WidthType.PERCENTAGE },
    rows: [
      new DOCX.TableRow({
        children: [
          new DOCX.TableCell({ width: { size: 35, type: DOCX.WidthType.PERCENTAGE }, children: left }),
          new DOCX.TableCell({ width: { size: 65, type: DOCX.WidthType.PERCENTAGE }, children: right })
        ]
      })
    ]
  });

  const doc = new DOCX.Document({ sections: [{ children: [heading, summaryTable, new DOCX.Paragraph({ text: '' }), grid] }] });
  return DOCX.Packer.toBlob(doc);

  // ---- helpers ----
  function row2(label: string, value: string): DOCX.TableRow {
    return new DOCX.TableRow({
      children: [
        new DOCX.TableCell({ width: { size: 30, type: DOCX.WidthType.PERCENTAGE }, children: [paraBold(label)], margins: margins() }),
        new DOCX.TableCell({ width: { size: 70, type: DOCX.WidthType.PERCENTAGE }, children: [new DOCX.Paragraph({ children: [new DOCX.TextRun({ text: value })] })], margins: margins() })
      ]
    });
  }
  function paraBold(text: string) { return new DOCX.Paragraph({ children: [new DOCX.TextRun({ text, bold: true })] }); }
  function paraMultiline(s: string) {
    return s ? s.split('\n').map(line => new DOCX.Paragraph({ children: [new DOCX.TextRun({ text: line })] })) : [new DOCX.Paragraph({ text: '—' })];
  }
  function bullets(items: string[]) {
    return items.map(s => new DOCX.Paragraph({ children: [new DOCX.TextRun({ text: s })], bullet: { level: 0 } }));
  }
  function paraIf(label: string, value?: string) {
    if (!value) return [] as DOCX.Paragraph[];
    return [new DOCX.Paragraph({ children: [new DOCX.TextRun({ text: `${label}: `, bold: true }), new DOCX.TextRun({ text: value })] })];
  }
  function margins() { return { left: 120, right: 120, top: 80, bottom: 80 }; }
  function projectRow(p: ProjectItem): DOCX.TableRow {
    const period = p?.period || [p?.from, p?.to].filter(Boolean).join(' – ') || '—';
    const rightParas: DOCX.Paragraph[] = [];
    if (p?.title) rightParas.push(paraBold(p.title));
    rightParas.push(new DOCX.Paragraph({ children: [new DOCX.TextRun({ text: `Rolle: ${p?.rolle || ''}` })] }));
    rightParas.push(new DOCX.Paragraph({ children: [new DOCX.TextRun({ text: `Kunde/Branche: ${p?.kunde || ''}` })] }));
    if (p?.taetigkeiten?.length) {
      rightParas.push(paraBold('Tätigkeiten'));
      rightParas.push(...p.taetigkeiten.map(t => new DOCX.Paragraph({ children: [new DOCX.TextRun({ text: t })], bullet: { level: 0 } })));
    }
    if (p?.technologien?.length) {
      rightParas.push(paraBold('Verwendete Technologien'));
      rightParas.push(...p.technologien.map(t => new DOCX.Paragraph({ children: [new DOCX.TextRun({ text: t })], bullet: { level: 0 } })));
    }
    return new DOCX.TableRow({
      children: [
        new DOCX.TableCell({ width: { size: 25, type: DOCX.WidthType.PERCENTAGE }, children: [new DOCX.Paragraph({ children: [new DOCX.TextRun({ text: period })] })], margins: margins() }),
        new DOCX.TableCell({ width: { size: 75, type: DOCX.WidthType.PERCENTAGE }, children: rightParas.length ? rightParas : [new DOCX.Paragraph({ text: '—' })], margins: margins() })
      ]
    });
  }
}

// ---------- Smart wrapper ----------
export async function tryGenerateDataportDoc(
  spHttp: SPHttpClient,
  maybeTemplateUrlOrServerRel: string | null,
  data: ProfileData
): Promise<Blob> {
  if (maybeTemplateUrlOrServerRel && await templateHasDocxtemplaterTags(spHttp, maybeTemplateUrlOrServerRel)) {
    return fillDataportTemplate(spHttp, maybeTemplateUrlOrServerRel, data);
  }
  return buildDataportDocx(data);
}

export function defaultTemplateUrl(): string {
  return `${BERATERPROFIL_SITE}/${LIB_INTERNAL_NAME}/Dataport CV Vorlage.docx`;
}
