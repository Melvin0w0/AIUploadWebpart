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

export type CorrespondenceKind = 'incoming' | 'outgoing' | 'unknown';

export function correspondenceKindFromName(value: string): CorrespondenceKind {
  if (isIncomingName(value)) {
    return 'incoming';
  }
  if (isOutgoingName(value)) {
    return 'outgoing';
  }
  return 'unknown';
}

export function correspondenceKindFromFileName(fileName: string): CorrespondenceKind {
  return correspondenceKindFromName(nameFromPdfFile(fileName));
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

export function generateIncomingName(now?: Date): string {
  const date = now || new Date();
  const yy = String(date.getFullYear()).slice(-2);
  const MM = twoDigits(date.getMonth() + 1);
  const dd = twoDigits(date.getDate());
  const HH = twoDigits(date.getHours());
  const mm = twoDigits(date.getMinutes());
  const ss = twoDigits(date.getSeconds());
  return `I${yy}${MM}${dd}${HH}${mm}${ss}Z0001`;
}

function twoDigits(value: number): string {
  return value < 10 ? `0${value}` : String(value);
}
