import { ProjectItem } from '../types';

export function computeBerufserfahrungFromProjects(projects: ProjectItem[]): string {
  const starts: Date[] = [];
  for (const project of projects) {
    const startDate = parseStartDateFromPeriod(project.period || '');
    if (startDate) starts.push(startDate);
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

export function parseStartDateFromPeriod(period: string): Date | null {
  const normalizedPeriod = (period || '').replace(/&/g, '/').trim().toLowerCase();
  if (!normalizedPeriod) return null;

  const seit = normalizedPeriod.match(/seit\s*(\d{1,2}\s*[./-]\s*\d{4}|\d{4})/i);
  if (seit) return parseMonthYearOrYear(seit[1]);

  const mmyyyy = normalizedPeriod.match(/(\d{1,2})\s*[./-]\s*(\d{4})/);
  if (mmyyyy) {
    const mm = clampInt(parseInt(mmyyyy[1], 10), 1, 12);
    const yyyy = parseInt(mmyyyy[2], 10);
    if (yyyy >= 1900 && yyyy <= 2100) return new Date(yyyy, mm - 1, 1);
  }

  const yyyyOnly = normalizedPeriod.match(/(\d{4})/);
  if (yyyyOnly) {
    const yyyy = parseInt(yyyyOnly[1], 10);
    if (yyyy >= 1900 && yyyy <= 2100) return new Date(yyyy, 0, 1);
  }

  return null;
}

export function parseMonthYearOrYear(value: string): Date | null {
  const normalized = (value || '').replace(/&/g, '/').trim();
  const monthYear = normalized.match(/(\d{1,2})\s*[./-]\s*(\d{4})/);
  if (monthYear) {
    const mm = clampInt(parseInt(monthYear[1], 10), 1, 12);
    const yyyy = parseInt(monthYear[2], 10);
    if (yyyy >= 1900 && yyyy <= 2100) return new Date(yyyy, mm - 1, 1);
  }
  const yearOnly = normalized.match(/(\d{4})/);
  if (yearOnly) {
    const yyyy = parseInt(yearOnly[1], 10);
    if (yyyy >= 1900 && yyyy <= 2100) return new Date(yyyy, 0, 1);
  }
  return null;
}

export function clampInt(v: number, min: number, max: number): number {
  return Math.max(min, Math.min(max, v));
}
