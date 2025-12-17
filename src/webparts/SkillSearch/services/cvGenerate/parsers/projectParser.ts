import { ProjectItem } from '../types';

export function parseProjectsCopyPaste(htmlDoc: Document): ProjectItem[] {
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

export function splitProjectRightColumn(right: string): {
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

export function findIndexRegex(arr: string[], re: RegExp): number {
  for (let i = 0; i < arr.length; i++) if (re.test(arr[i])) return i;
  return -1;
}

// IMPORTANT: preserve DOM order (so description + bullets don’t get scrambled)
export function extractCellTextPreservingOrder(cell: HTMLTableCellElement): string {
  const out: string[] = [];

  const push = (txt: string) => {
    const text = normalizeWhitespace(txt);
    if (!text) return;
    if (out.length && out[out.length - 1] === text) return;
    out.push(text);
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

export function normalizeWhitespace(value: string): string {
  return (value || '')
    .replace(/\u00A0/g, ' ')
    .replace(/[ \t]+\n/g, '\n')
    .replace(/\n[ \t]+/g, '\n')
    .replace(/[ \t]{2,}/g, ' ')
    .trim();
}
