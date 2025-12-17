export function grabBetween(start: RegExp, end: RegExp, paras: HTMLParagraphElement[]): string {
  const texts = paras.map(p => (p.textContent || '').trim());
  const s = texts.findIndex(t => start.test(t));
  if (s < 0) return '';
  const eRel = texts.slice(s + 1).findIndex(t => end.test(t));
  const slice = texts.slice(s + 1, eRel > -1 ? s + 1 + eRel : undefined);
  return slice.join('\n').trim();
}
