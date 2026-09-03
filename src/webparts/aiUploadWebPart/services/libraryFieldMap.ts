import { INamedValue } from './uploadDestination';
import { isNameField } from '../constants/defaultFormFields';

export interface ILibraryField {
  InternalName: string;
  Title: string;
  TypeAsString: string;
  Hidden?: boolean;
  ReadOnlyField?: boolean;
}

export type FieldPayload = { [internalName: string]: string | number | boolean };

const SKIP_TYPES: string[] = [
  'Lookup',
  'LookupMulti',
  'User',
  'UserMulti',
  'TaxonomyFieldType',
  'TaxonomyFieldTypeMulti',
  'Calculated',
  'Computed',
  'File',
  'Attachments',
  'Guid',
  'Counter',
  'ContentTypeId',
  'ModStat',
  'WorkflowStatus',
  'URL'
];

export const SYSTEM_FIELD_NAMES: string[] = [
  'id',
  'contenttype',
  'contenttypeid',
  'modified',
  'created',
  'author',
  'editor',
  'fileleafref',
  'fileref',
  'filedirref',
  'fsobjtype',
  'uniqueid',
  'owshiddenversion',
  '_uiversionstring',
  'attachments',
  'edit',
  'linkfilename',
  'linkfilename2',
  'docicon',
  'itemchildcount',
  'foldercount',
  'filetype',
  'htmlfilename',
  'encodedabsurl',
  'basename',
  'filesizedisplay',
  'appauthor',
  'appeditor',
  'selecttitle',
  'selectfilename',
  'linktitle',
  'linktitlenomenu',
  'serverurl',
  'metainfo',
  '_level',
  '_iscurrentversion',
  'complianceassetid',
  'checkoutuser',
  'virusstatus',
  '_checkincomment',
  '_moderationstatus',
  '_moderationcomments'
];

const ALIASES: { [key: string]: string[] } = {
  name: ['title', 'name', 'documentname'],
  registrationnumber: ['registrationnumber', 'regno', 'regnumber'],
  leadingbl: ['leadingbl', 'businessline', 'bl'],
  projectnumber: ['projectnumber', 'projectno', 'project'],
  subprojectnumber: ['subprojectnumber', 'subprojectno', 'subproject'],
  organization: ['organization', 'organisation', 'org'],
  sender: ['sender', 'from'],
  receiver: ['receiver', 'recipient', 'to'],
  subject: ['subject'],
  fileno: ['fileno', 'filenumber', 'filenum'],
  refno: ['refno', 'referenceno', 'referencenumber', 'reference'],
  issuedate: ['issuedate', 'documentdate', 'date'],
  attachment: ['attachment'],
  scan: ['scan', 'scanned'],
  remark: ['remark', 'remarks', 'comment', 'comments'],
  location: ['location'],
  cctoaecom: ['cctoaecom', 'cc', 'ccto']
};

export function isWritableLibraryField(field: ILibraryField): boolean {
  const internal = (field.InternalName || '').toLowerCase();
  const type = field.TypeAsString || '';
  return !field.Hidden &&
    !field.ReadOnlyField &&
    SYSTEM_FIELD_NAMES.indexOf(internal) < 0 &&
    SKIP_TYPES.indexOf(type) < 0;
}

export function buildFieldPayload(fieldValues: INamedValue[], columns: ILibraryField[]): FieldPayload {
  const payload: FieldPayload = {};
  const used: { [internal: string]: boolean } = {};
  const writable = columns.filter(isWritableLibraryField);
  const ordered = fieldValues.slice().sort((a, b) => {
    if (isNameField(a.label) === isNameField(b.label)) {
      return 0;
    }
    return isNameField(a.label) ? -1 : 1;
  });

  ordered.forEach((field) => {
    const value = (field.value || '').trim();
    if (!value) {
      return;
    }
    const column = findColumn(field.label, writable, used);
    if (!column) {
      return;
    }
    const converted = convertValue(value, column.TypeAsString);
    if (converted === undefined) {
      return;
    }
    payload[column.InternalName] = converted;
    used[column.InternalName.toLowerCase()] = true;
  });

  return payload;
}

function findColumn(
  label: string,
  columns: ILibraryField[],
  used: { [internal: string]: boolean }
): ILibraryField | undefined {
  let best: ILibraryField | undefined;
  let bestScore = 0;
  columns.forEach((column) => {
    if (used[(column.InternalName || '').toLowerCase()]) {
      return;
    }
    const score = matchScore(label, column);
    if (score > bestScore) {
      bestScore = score;
      best = column;
    }
  });
  return bestScore > 0 ? best : undefined;
}

function matchScore(label: string, column: ILibraryField): number {
  const labelKey = normalizeKey(label);
  const titleKey = normalizeKey(column.Title);
  const internalKey = normalizeKey(column.InternalName);
  const encoded = normalizeKey(encodeSharePointName(label));

  if (equalsIgnoreCase(column.Title, label)) {
    return 100;
  }
  if (equalsIgnoreCase(column.InternalName, label)) {
    return 95;
  }
  if (equalsIgnoreCase(column.InternalName, encodeSharePointName(label))) {
    return 90;
  }
  if (titleKey && titleKey === labelKey) {
    return 80;
  }
  if (internalKey && (internalKey === labelKey || internalKey === encoded)) {
    return 70;
  }

  const aliases = ALIASES[labelKey] || [];
  if (aliases.indexOf(titleKey) >= 0 || aliases.indexOf(internalKey) >= 0) {
    return 50;
  }
  return 0;
}

function convertValue(value: string, type: string): string | number | boolean | undefined {
  if (type === 'DateTime') {
    return toSharePointDate(value) || value;
  }
  if (type === 'Boolean') {
    return toBoolean(value);
  }
  if (type === 'Number' || type === 'Currency' || type === 'Integer') {
    const number = parseFloat(value.replace(/,/g, ''));
    return isNaN(number) ? undefined : number;
  }
  return value;
}

function toBoolean(value: string): boolean | undefined {
  const key = value.trim().toLowerCase();
  if (['yes', 'true', '1', 'y', 'checked'].indexOf(key) >= 0) {
    return true;
  }
  if (['no', 'false', '0', 'n', 'unchecked'].indexOf(key) >= 0) {
    return false;
  }
  return undefined;
}

function toSharePointDate(value: string): string | undefined {
  const trimmed = value.trim();
  if (/^\d{4}-\d{2}-\d{2}/.test(trimmed)) {
    const parsed = new Date(trimmed);
    return isNaN(parsed.getTime()) ? undefined : parsed.toISOString();
  }
  const dmy = trimmed.match(/^(\d{1,2})[/.-](\d{1,2})[/.-](\d{2,4})$/);
  if (dmy) {
    const day = parseInt(dmy[1], 10);
    const month = parseInt(dmy[2], 10);
    let year = parseInt(dmy[3], 10);
    if (year < 100) {
      year += 2000;
    }
    const parsed = new Date(Date.UTC(year, month - 1, day));
    return isNaN(parsed.getTime()) ? undefined : parsed.toISOString();
  }
  const millis = Date.parse(trimmed);
  return isNaN(millis) ? undefined : new Date(millis).toISOString();
}

export function encodeSharePointName(displayName: string): string {
  return (displayName || '').replace(/([^A-Za-z0-9])/g, (ch) => {
    let hex = ch.charCodeAt(0).toString(16);
    while (hex.length < 4) {
      hex = `0${hex}`;
    }
    return `_x${hex}_`;
  });
}

function normalizeKey(value: string): string {
  return decodeSharePointName(value || '')
    .toLowerCase()
    .replace(/[^a-z0-9]/g, '');
}

function decodeSharePointName(value: string): string {
  return value.replace(/_x([0-9a-fA-F]{4})_/g, (_all, hex: string) =>
    String.fromCharCode(parseInt(hex, 16))
  );
}

function equalsIgnoreCase(left: string, right: string): boolean {
  return (left || '').toLowerCase() === (right || '').toLowerCase();
}
