import { isEoiProjectNumber } from './projectNumber';

export type UploadType = 'normal' | 'confidentialInvoice' | 'confidentialMisc';

export const UPLOAD_TYPE_NORMAL: UploadType = 'normal';
export const UPLOAD_TYPE_CONFIDENTIAL_INVOICE: UploadType = 'confidentialInvoice';
export const UPLOAD_TYPE_CONFIDENTIAL_MISC: UploadType = 'confidentialMisc';

export const CONFIDENTIAL_INVOICE_FOLDER: string = 'Confidential_Invoice';
export const CONFIDENTIAL_MISC_FOLDER: string = 'Confidential_MISC';

export const UPLOAD_TYPE_OPTIONS: { key: UploadType; label: string }[] = [
  { key: UPLOAD_TYPE_NORMAL, label: 'Normal' },
  { key: UPLOAD_TYPE_CONFIDENTIAL_INVOICE, label: 'Confidential Invoice' },
  { key: UPLOAD_TYPE_CONFIDENTIAL_MISC, label: 'Confidential MISC' }
];

export const EOI_UPLOAD_TYPE_OPTIONS: { key: UploadType; label: string }[] = [
  { key: UPLOAD_TYPE_NORMAL, label: 'Normal' },
  { key: UPLOAD_TYPE_CONFIDENTIAL_MISC, label: 'Confidential MISC' }
];

export function uploadTypeOptionsForProjectNumber(projectNumber: string): { key: UploadType; label: string }[] {
  return isEoiProjectNumber(projectNumber) ? EOI_UPLOAD_TYPE_OPTIONS : UPLOAD_TYPE_OPTIONS;
}

export function canonicalUploadType(value: string): UploadType {
  const key = (value || '').trim().toLowerCase().replace(/[\s-]+/g, '');
  if (key === UPLOAD_TYPE_CONFIDENTIAL_INVOICE.toLowerCase() || key === 'confidentialinvoice') {
    return UPLOAD_TYPE_CONFIDENTIAL_INVOICE;
  }
  if (key === UPLOAD_TYPE_CONFIDENTIAL_MISC.toLowerCase() || key === 'confidentialmisc') {
    return UPLOAD_TYPE_CONFIDENTIAL_MISC;
  }
  return UPLOAD_TYPE_NORMAL;
}

export function canonicalUploadTypeForProjectNumber(value: string, projectNumber: string): UploadType {
  const uploadType = canonicalUploadType(value);
  if (isEoiProjectNumber(projectNumber) && uploadType === UPLOAD_TYPE_CONFIDENTIAL_INVOICE) {
    return UPLOAD_TYPE_NORMAL;
  }
  return uploadType;
}

export function uploadTypeFromFolderName(name: string): UploadType | undefined {
  const key = (name || '').trim().toLowerCase().replace(/[\s_-]+/g, '');
  if (key === 'confidentialinvoice') {
    return UPLOAD_TYPE_CONFIDENTIAL_INVOICE;
  }
  if (key === 'confidentialmisc') {
    return UPLOAD_TYPE_CONFIDENTIAL_MISC;
  }
  return undefined;
}

export function constrainUploadType(value: string, available: UploadType[]): UploadType {
  const uploadType = canonicalUploadType(value);
  if (available.indexOf(uploadType) >= 0) {
    return uploadType;
  }
  return available.length > 0 ? available[0] : UPLOAD_TYPE_NORMAL;
}

export function confidentialFolderForUploadType(uploadType: UploadType): string {
  if (uploadType === UPLOAD_TYPE_CONFIDENTIAL_INVOICE) {
    return CONFIDENTIAL_INVOICE_FOLDER;
  }
  if (uploadType === UPLOAD_TYPE_CONFIDENTIAL_MISC) {
    return CONFIDENTIAL_MISC_FOLDER;
  }
  return '';
}
