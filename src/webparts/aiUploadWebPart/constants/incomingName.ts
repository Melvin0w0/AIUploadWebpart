export const INCOMING_FOLDER_NAME: string = 'InComing';
export const OUTGOING_FOLDER_NAME: string = 'Outgoing';

export function isIncomingName(value: string): boolean {
  const text = nameText(value, 18);
  if (!text) {
    return false;
  }
  if (text.charAt(0).toUpperCase() !== 'I') {
    return false;
  }
  return /^[A-Za-z]$/.test(text.charAt(13));
}

export function isOutgoingName(value: string): boolean {
  const text = nameText(value, 11);
  if (!text) {
    return false;
  }
  return /^[A-Za-z]$/.test(text.charAt(10));
}

function nameText(value: string, length: number): string {
  const trimmed = (value || '').trim();
  const compact = trimmed.replace(/\s+/g, '');
  const text = compact.length === length ? compact : trimmed;
  return text.length === length ? text : '';
}

export function nameFromPdfFile(fileName: string): string {
  const name = (fileName || '').replace(/^.*[\\/]/, '').trim();
  return name.replace(/\.pdf$/i, '');
}
