export const PROJECT_NUMBER_LABEL: string = 'Project Number';
export const PROJECT_NUMBER_MAX_DIGITS: number = 8;

export function isProjectNumberField(label: string): boolean {
  return (label || '').trim().toLowerCase() === PROJECT_NUMBER_LABEL.toLowerCase();
}

export function sanitizeProjectNumber(value: string): string {
  return (value || '').replace(/\D/g, '').substring(0, PROJECT_NUMBER_MAX_DIGITS);
}

export function isValidProjectNumber(value: string): boolean {
  return /^\d{1,8}$/.test((value || '').trim());
}

export function projectNumberFromYourRef(yourRef: string): string {
  const source = (yourRef || '').trim();
  if (!source) {
    return '';
  }
  const slash = source.search(/[/\\\uFF0F]/);
  if (slash < 0) {
    return '';
  }
  const before = source.substring(0, slash);
  const atEnd = before.match(/(\d{8})\s*$/);
  if (atEnd) {
    return atEnd[1];
  }
  const digits = before.replace(/\D/g, '');
  if (digits.length >= 8) {
    return digits.substring(digits.length - 8);
  }
  return '';
}

