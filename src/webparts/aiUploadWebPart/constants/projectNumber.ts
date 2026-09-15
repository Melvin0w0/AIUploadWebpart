export const PROJECT_NUMBER_LABEL: string = 'Project Number';
export const PROJECT_NUMBER_MAX_DIGITS: number = 8;

export function isProjectNumberField(label: string): boolean {
  return (label || '').trim().toLowerCase() === PROJECT_NUMBER_LABEL.toLowerCase();
}

export function isEoiProjectNumber(value: string): boolean {
  return /^.+-EOI\.?$/i.test((value || '').trim());
}

export function eoiProjectNumberFromCode(code: string): string {
  const text = (code || '').trim().replace(/-EOI\.?$/i, '').trim();
  return text ? `${text}-EOI` : '';
}

export function sanitizeProjectNumber(value: string): string {
  const text = (value || '').trim();
  const eoi = text.match(/^(.+)-EOI\.?$/i);
  if (eoi && eoi[1].trim()) {
    return `${eoi[1].trim()}-EOI`;
  }
  return text.replace(/\D/g, '').substring(0, PROJECT_NUMBER_MAX_DIGITS);
}

export function isValidProjectNumber(value: string): boolean {
  const text = (value || '').trim();
  return /^\d{1,8}$/.test(text) || /^[A-Za-z0-9]+-EOI$/i.test(text);
}

export function projectNumberFromRef(refValue: string): string {
  const source = (refValue || '').trim();
  if (!source) {
    return '';
  }
  return eightDigitsBeforeSeparator(source, /[/\\\uFF0F]/)
    || eightDigitsBeforeSeparator(source, /[-–—\uFF0D]/);
}

function eightDigitsBeforeSeparator(source: string, separator: RegExp): string {
  const index = source.search(separator);
  if (index < 0) {
    return '';
  }
  const before = source.substring(0, index);
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

