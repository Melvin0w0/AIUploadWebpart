export type LabelType = 'normal' | 'confidential' | 'invoice' | 'site';

export const LABEL_TYPE_NORMAL: LabelType = 'normal';
export const LABEL_TYPE_CONFIDENTIAL: LabelType = 'confidential';
export const LABEL_TYPE_INVOICE: LabelType = 'invoice';
export const LABEL_TYPE_SITE: LabelType = 'site';

export const LABEL_TYPE_OPTIONS: { key: LabelType; label: string }[] = [
  { key: LABEL_TYPE_NORMAL, label: 'Normal' },
  { key: LABEL_TYPE_CONFIDENTIAL, label: 'Confidential' },
  { key: LABEL_TYPE_INVOICE, label: 'Invoice' },
  { key: LABEL_TYPE_SITE, label: 'Site' }
];

const LABEL_FILES: { [key: string]: string } = {
  normal: 'WCONF.jpg',
  confidential: 'CONFD.jpg',
  invoice: 'INVOC.jpg',
  site: 'GOVTS.jpg'
};

export function canonicalLabelType(value: string): LabelType {
  const key = (value || '').trim().toLowerCase();
  if (key === LABEL_TYPE_CONFIDENTIAL) {
    return LABEL_TYPE_CONFIDENTIAL;
  }
  if (key === LABEL_TYPE_INVOICE) {
    return LABEL_TYPE_INVOICE;
  }
  if (key === LABEL_TYPE_SITE) {
    return LABEL_TYPE_SITE;
  }
  return LABEL_TYPE_NORMAL;
}

export function labelImageFileName(labelType: LabelType): string {
  return LABEL_FILES[labelType] || LABEL_FILES.normal;
}
