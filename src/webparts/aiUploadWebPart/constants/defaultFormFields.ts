export const NAME_LABEL: string = 'Name';
export const REGISTRATION_NUMBER_LABEL: string = 'Registration Number';
export const SENDER_LABEL: string = 'Sender';

export const RECEIVER_LABEL: string = 'Receiver';

export const SUBJECT_LABEL: string = 'Subject';
export const REF_NO_LABEL: string = 'Ref No';

export function isNameField(label: string): boolean {
  return (label || '').trim().toLowerCase() === NAME_LABEL.toLowerCase();
}

export function isRegistrationNumberField(label: string): boolean {
  return (label || '').trim().toLowerCase() === REGISTRATION_NUMBER_LABEL.toLowerCase();
}

export function isSenderField(label: string): boolean {
  return (label || '').trim().toLowerCase() === SENDER_LABEL.toLowerCase();
}

export function isReceiverField(label: string): boolean {
  return (label || '').trim().toLowerCase() === RECEIVER_LABEL.toLowerCase();
}

export function isSubjectField(label: string): boolean {
  return (label || '').trim().toLowerCase() === SUBJECT_LABEL.toLowerCase();
}

export function isRefNoField(label: string): boolean {
  return (label || '').trim().toLowerCase() === REF_NO_LABEL.toLowerCase();
}

export const DEFAULT_FORM_FIELDS: string[] = [
  'Name',
  'Registration Number',
  'Leading BL',
  'Project Number',
  'Sub-Project Number',
  'Organization',
  'Sender',
  'Receiver',
  'Subject',
  'File No',
  'Ref No',
  'Issue Date',
  'Attachment',
  'Scan',
  'Remark',
  'Location',
  'cc to AECOM'
];

export const REQUIRED_FORM_FIELDS: string[] = [
  'Name',
  'Registration Number',
  'Leading BL',
  'Project Number',
  'Sub-Project Number',
  'Organization',
  'Sender',
  'Receiver',
  'Subject',
  'Ref No'
];

export function isRequiredField(label: string): boolean {
  const key = (label || '').trim().toLowerCase();
  return REQUIRED_FORM_FIELDS.some((name) => name.toLowerCase() === key);
}

export function missingRequiredFields(fields: { label: string; value: string }[]): string[] {
  return REQUIRED_FORM_FIELDS.filter((label) => {
    const field = fields.filter((item) => item.label.toLowerCase() === label.toLowerCase())[0];
    return !field || !(field.value || '').trim();
  });
}

export const DEFAULT_FORM_FIELDS_TEXT: string = DEFAULT_FORM_FIELDS.join('\n');

