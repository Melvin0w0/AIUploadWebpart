export const YES_VALUE: string = 'Yes';
export const NO_VALUE: string = 'No';
export const YES_NO_OPTIONS: string[] = [YES_VALUE, NO_VALUE];
export const ATTACHMENT_LABEL: string = 'Attachment';
export const SCAN_LABEL: string = 'Scan';
export const CC_TO_AECOM_LABEL: string = 'cc to AECOM';

export function isAttachmentField(label: string): boolean {
  return (label || '').trim().toLowerCase() === ATTACHMENT_LABEL.toLowerCase();
}

export function isScanField(label: string): boolean {
  return (label || '').trim().toLowerCase() === SCAN_LABEL.toLowerCase();
}

export function isCcToAecomField(label: string): boolean {
  return (label || '').trim().toLowerCase() === CC_TO_AECOM_LABEL.toLowerCase();
}

export function isYesNoChoiceField(label: string): boolean {
  return isAttachmentField(label) || isScanField(label) || isCcToAecomField(label);
}

export function canonicalYesNo(value: string): string {
  const trimmed = (value || '').replace(/\s+/g, ' ').trim();
  if (!trimmed) {
    return YES_VALUE;
  }
  const key = trimmed.toLowerCase();
  if (/^(no|n|false|0|nil|none|n\/a|na|-|無|否)$/.test(key) ||
    /\b(?:no attachment|not attached|without attachment|nil enclosure)\b/.test(key)) {
    return NO_VALUE;
  }
  if (/^(yes|y|true|1|有|是)$/.test(key) ||
    /\b(?:yes|attached|attachment|encl(?:osure)?s?|appendix|scan(?:ned)?)\b/.test(key)) {
    return YES_VALUE;
  }
  return YES_VALUE;
}
