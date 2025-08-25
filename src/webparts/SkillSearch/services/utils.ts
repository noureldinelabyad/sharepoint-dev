// Generic helpers (chunk, initials, colors, xml)

/** Split array into chunks of size N. */
export function chunk<T>(arr: T[], size: number): T[][] {
  const out: T[][] = [];
  for (let i = 0; i < arr.length; i += size) out.push(arr.slice(i, i + size));
  return out;
}

/** “JD” -> initials. Falls back safely. */
export function initials(name?: string): string {
  if (!name) return '?';
  const parts = name.trim().split(/\s+/);
  const first = parts[0]?.[0] ?? '';
  const second = parts.length > 1 ? parts[1][0] : (parts[0]?.[1] ?? '');
  return (first + second).toUpperCase();
}

/** Stable color from name (12-color palette). */
export function colorFromName(name: string): string {
  const palette = ['#3867D6','#20BF6B','#F7B731','#EB3B5A','#8854D0','#0FB9B1','#4B7BEC','#26DE81','#FED330','#FA8231','#A55EEA','#2D98DA'];
  let h = 0; for (let i = 0; i < name.length; i++) h = (h * 31 + name.charCodeAt(i)) >>> 0;
  return palette[h % palette.length];
}

/** Escape minimal XML entities. */
export function escapeXml(s: string): string {
  return s.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&apos;');
}

/** Teams-like initials avatar (SVG → data URL). */
export function makeInitialsAvatar(name: string, size = 96): string {
  const _initials = initials(name);
  const bg = colorFromName(name);
  const svg =
    `<svg xmlns='http://www.w3.org/2000/svg' width='${size}' height='${size}'>` +
    `<rect width='100%' height='100%' rx='${size/2}' ry='${size/2}' fill='${bg}'/>` +
    `<text x='50%' y='54%' dominant-baseline='middle' text-anchor='middle' ` +
    `font-family='Segoe UI, Arial, sans-serif' font-size='${Math.round(size*0.42)}' fill='#fff' font-weight='600'>` +
    `${escapeXml(_initials)}</text></svg>`;
  return `data:image/svg+xml;utf8,${encodeURIComponent(svg)}`;
}
