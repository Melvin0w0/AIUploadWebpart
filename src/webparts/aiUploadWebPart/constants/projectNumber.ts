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
