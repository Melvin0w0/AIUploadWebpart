export const SUB_PROJECT_NUMBER_LABEL: string = 'Sub-Project Number';
export const SUB_PROJECT_NONE: string = 'None';

const options: string[] = [SUB_PROJECT_NONE];
for (let number = 1; number <= 99; number++) {
  options.push(String(number));
}

export const SUB_PROJECT_NUMBER_OPTIONS: string[] = options;

export function isSubProjectNumberField(label: string): boolean {
  return (label || '').trim().toLowerCase() === SUB_PROJECT_NUMBER_LABEL.toLowerCase();
}

export function canonicalSubProjectNumber(value: string): string {
  const trimmed = (value || '').trim();
  if (!trimmed || /^none$/i.test(trimmed)) {
    return SUB_PROJECT_NONE;
  }
  const digits = trimmed.replace(/\D/g, '');
  if (!digits) {
    return SUB_PROJECT_NONE;
  }
  const parsed = parseInt(digits, 10);
  if (parsed >= 1 && parsed <= 99) {
    return String(parsed);
  }
  return SUB_PROJECT_NONE;
}
