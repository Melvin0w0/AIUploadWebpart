export const ISSUE_DATE_LABEL: string = 'Issue Date';
export const ISSUE_DATE_FORMAT: string = 'dd/MM/yyyy';
export const ISSUE_DATE_DISPLAY_LABEL: string = `${ISSUE_DATE_LABEL} (${ISSUE_DATE_FORMAT})`;

const MONTHS: string[] = [
  'january', 'february', 'march', 'april', 'may', 'june',
  'july', 'august', 'september', 'october', 'november', 'december'
];

const MONTH_ABBREV: string[] = [
  'jan', 'feb', 'mar', 'apr', 'may', 'jun',
  'jul', 'aug', 'sep', 'oct', 'nov', 'dec'
];

export function isIssueDateField(label: string): boolean {
  return (label || '').trim().toLowerCase() === ISSUE_DATE_LABEL.toLowerCase();
}

export function formatIssueDate(date: Date): string {
  return `${pad2(date.getDate())}/${pad2(date.getMonth() + 1)}/${date.getFullYear()}`;
}

export function parseIssueDate(value: string): Date | undefined {
  const trimmed = (value || '').replace(/\s+/g, ' ').trim();
  if (!trimmed) {
    return undefined;
  }

  const iso = trimmed.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})/);
  if (iso) {
    return validDate(parseInt(iso[1], 10), parseInt(iso[2], 10), parseInt(iso[3], 10));
  }

  const dmy = trimmed.match(/^(\d{1,2})[/.\-](\d{1,2})[/.\-](\d{2,4})$/);
  if (dmy) {
    let year = parseInt(dmy[3], 10);
    if (year < 100) {
      year += 2000;
    }
    return validDate(year, parseInt(dmy[2], 10), parseInt(dmy[1], 10));
  }

  const named = trimmed.match(/^(\d{1,2})(?:st|nd|rd|th)?\s+([A-Za-z]+)\.?\s+(\d{2,4})$/i);
  if (named) {
    const month = monthNumber(named[2]);
    if (month > 0) {
      let year = parseInt(named[3], 10);
      if (year < 100) {
        year += 2000;
      }
      return validDate(year, month, parseInt(named[1], 10));
    }
  }

  const millis = Date.parse(trimmed);
  if (!isNaN(millis)) {
    const parsed = new Date(millis);
    return validDate(parsed.getFullYear(), parsed.getMonth() + 1, parsed.getDate());
  }
  return undefined;
}

export function sanitizeIssueDate(value: string): string {
  const date = parseIssueDate(value);
  return date ? formatIssueDate(date) : (value || '').trim();
}

function monthNumber(name: string): number {
  const key = (name || '').replace(/\./g, '').toLowerCase();
  const full = MONTHS.indexOf(key);
  if (full >= 0) {
    return full + 1;
  }
  const abbrev = MONTH_ABBREV.indexOf(key.substring(0, 3));
  return abbrev >= 0 ? abbrev + 1 : 0;
}

function validDate(year: number, month: number, day: number): Date | undefined {
  if (year < 1900 || year > 2100 || month < 1 || month > 12 || day < 1 || day > 31) {
    return undefined;
  }
  const date = new Date(year, month - 1, day);
  if (date.getFullYear() !== year || date.getMonth() !== month - 1 || date.getDate() !== day) {
    return undefined;
  }
  return date;
}

function pad2(value: number): string {
  return value < 10 ? `0${value}` : String(value);
}
