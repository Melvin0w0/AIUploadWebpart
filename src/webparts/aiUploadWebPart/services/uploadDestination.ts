import {
  LEADING_BL_LABEL,
  resolveLeadingBlSite
} from '../constants/blSiteMap';
import {
  isProjectNumberField,
  sanitizeProjectNumber
} from '../constants/projectNumber';
import { INCOMING_FOLDER_NAME, OUTGOING_FOLDER_NAME, isIncomingName, isOutgoingName } from '../constants/incomingName';

export interface INamedValue {
  label: string;
  value: string;
}

export interface IUploadDestinationConfig {
  tenantUrl: string;
  libraryName: string;
  folderPathTemplate: string;
}

export interface IResolvedUploadDestination {
  siteUrl: string;
  libraryName: string;
  folderPath: string;
  leadingBl: string;
  projectNumber: string;
  missingFields: string[];
  unrecognizedLeadingBl: string;
}

export function fieldValueMap(fields: INamedValue[]): { [label: string]: string } {
  const values: { [label: string]: string } = {};
  fields.forEach((field) => {
    const value = (field.value || '').trim();
    values[field.label] = isProjectNumberField(field.label) ? sanitizeProjectNumber(value) : value;
  });
  return values;
}

export function resolveUploadDestination(
  fields: INamedValue[],
  config: IUploadDestinationConfig
): IResolvedUploadDestination {
  const values = fieldValueMap(fields);
  const extras: { [label: string]: string } = {
    tenant: trimSlash(config.tenantUrl)
  };
  const folderTemplate = (config.folderPathTemplate || '').trim() || '{Project Number}';
  const libraryName = (config.libraryName || '').trim() || 'Project Documents';
  const leadingBl = lookup(values, LEADING_BL_LABEL);
  const blSite = resolveLeadingBlSite(leadingBl, config.tenantUrl);
  const missingFields = collectMissingFields(folderTemplate, values, extras);
  if (!leadingBl && missingFields.indexOf(LEADING_BL_LABEL) < 0) {
    missingFields.unshift(LEADING_BL_LABEL);
  }

  const folderPath = collapsePath(replaceTokens(folderTemplate, values, extras, true));
  const nameValue = lookup(values, 'Name');
  return {
    siteUrl: blSite ? blSite.siteUrl : '',
    libraryName,
    folderPath: appendCorrespondenceFolder(folderPath, nameValue),
    leadingBl: blSite ? blSite.name : leadingBl,
    projectNumber: lookup(values, 'Project Number'),
    missingFields,
    unrecognizedLeadingBl: leadingBl && !blSite ? leadingBl : ''
  };
}

export function replaceTokens(
  template: string,
  values: { [label: string]: string },
  extras: { [label: string]: string },
  sanitize: boolean
): string {
  return template.replace(/\{([^}]+)\}/g, (_match, rawName: string) => {
    const name = (rawName || '').trim();
    const extra = lookup(extras, name);
    if (extra) {
      return extra;
    }
    const value = lookup(values, name);
    if (!value) {
      return '';
    }
    return sanitize ? sanitizeSegment(value) : value;
  });
}

function collectMissingFields(
  template: string,
  values: { [label: string]: string },
  extras: { [label: string]: string }
): string[] {
  const missing: string[] = [];
  const tokens = template.match(/\{([^}]+)\}/g) || [];
  tokens.forEach((token) => {
    const name = token.substring(1, token.length - 1).trim();
    if (!name || lookup(extras, name) || lookup(values, name)) {
      return;
    }
    if (missing.indexOf(name) < 0) {
      missing.push(name);
    }
  });
  return missing;
}

function lookup(values: { [label: string]: string }, name: string): string {
  if (values[name]) {
    return values[name];
  }
  const key = Object.keys(values).filter((item) => item.toLowerCase() === name.toLowerCase())[0];
  return key ? values[key] : '';
}

function appendCorrespondenceFolder(folderPath: string, nameValue: string): string {
  if (!folderPath) {
    return folderPath;
  }
  if (isOutgoingName(nameValue)) {
    return appendFolderSegment(folderPath, OUTGOING_FOLDER_NAME);
  }
  if (isIncomingName(nameValue)) {
    return appendFolderSegment(folderPath, INCOMING_FOLDER_NAME);
  }
  return folderPath;
}

function appendFolderSegment(folderPath: string, segment: string): string {
  const parts = folderPath.split('/');
  if (parts[parts.length - 1].toLowerCase() === segment.toLowerCase()) {
    return folderPath;
  }
  return `${folderPath}/${segment}`;
}

export function sanitizeSegment(value: string): string {
  return value
    .replace(/[\\/:*?"<>|#%]/g, '-')
    .replace(/\s+/g, ' ')
    .replace(/\.+$/g, '')
    .trim();
}

export function collapsePath(path: string): string {
  return path
    .split('/')
    .map((segment) => segment.trim())
    .filter((segment) => segment.length > 0)
    .join('/');
}

export function trimSlash(value: string): string {
  return (value || '').replace(/\/+$/, '');
}

export function buildUploadFolderUrl(destination: IResolvedUploadDestination): string {
  return buildUploadFileUrl(destination, '');
}

export function buildUploadFileUrl(destination: IResolvedUploadDestination, uploadFileName: string): string {
  if (!destination.siteUrl) {
    return '';
  }
  const segments = [destination.libraryName]
    .concat(destination.folderPath ? destination.folderPath.split('/') : [])
    .concat(uploadFileName ? [uploadFileName] : [])
    .map((segment) => segment.trim())
    .filter((segment) => segment.length > 0)
    .map((segment) => encodeURIComponent(segment));
  return `${trimSlash(destination.siteUrl)}/${segments.join('/')}`;
}

export function fileNameFromFields(file: File, fields: INamedValue[]): string {
  const values = fieldValueMap(fields);
  const preferred = sanitizeSegment(values['Registration Number'] || values['File No'] || values['Ref No'] || '');
  const original = sanitizeSegment(file.name.replace(/\.pdf$/i, '')) || 'document';
  const base = preferred || original;
  return /\.pdf$/i.test(base) ? base : `${base}.pdf`;
}
