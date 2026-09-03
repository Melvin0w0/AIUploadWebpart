import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { canonicalLeadingBl } from '../constants/blSiteMap';
import { sanitizeProjectNumber } from '../constants/projectNumber';

export const NOTIFICATION_SETUP_LIST_TITLE: string = 'Notification Set-up';
export const LEADING_BUSINESS_LINE_INTERNAL: string = 'Leading_x0020_Business_x0020_Lin';

export interface INotificationSetupLookup {
  leadingBl: string;
  thresholdExceeded?: boolean;
}

interface IListColumns {
  projectNoInternal: string;
  projectNoType: string;
  leadingBlInternal: string;
  leadingBlType: string;
}

interface IFieldMeta {
  InternalName?: string;
  Title?: string;
  TypeAsString?: string;
}

const LIST_TITLES: string[] = [
  'Notification Set-up',
  'Notification Setup',
  'Notification Set up'
];

const PROJECT_FIELD_FALLBACKS: string[] = [
  'Project_x0020_No',
  'Project_x0020_no',
  'Project_x0020_Number',
  'ProjectNo'
];

let cached: { siteUrl: string; listTitle: string; columns: IListColumns } | undefined;

export async function lookupLeadingBlFromNotificationSetup(
  http: SPHttpClient,
  siteUrl: string,
  projectNumber: string
): Promise<INotificationSetupLookup> {
  const projectNo = sanitizeProjectNumber(projectNumber);
  if (!projectNo || !siteUrl || !http) {
    return { leadingBl: '' };
  }

  try {
    const resolved = await resolveList(http, siteUrl);
    if (!resolved) {
      return { leadingBl: '' };
    }

    const candidates = projectNumberCandidates(projectNo);
    for (let i = 0; i < candidates.length; i++) {
      try {
        const item = await queryMatchingItem(
          http,
          resolved.siteUrl,
          resolved.listTitle,
          resolved.columns,
          candidates[i]
        );
        const leadingBl = item
          ? canonicalLeadingBl(readFieldText(item, resolved.columns.leadingBlInternal))
          : '';
        if (leadingBl) {
          return { leadingBl };
        }
        if (item) {
          return { leadingBl: '' };
        }
      } catch (err) {
        const message = err instanceof Error ? err.message : String(err);
        if (isThresholdError(message)) {
          return { leadingBl: '', thresholdExceeded: true };
        }
        if (i === candidates.length - 1) {
          return { leadingBl: '' };
        }
      }
    }
  } catch (err) {
    const message = err instanceof Error ? err.message : String(err);
    if (isThresholdError(message)) {
      return { leadingBl: '', thresholdExceeded: true };
    }
  }

  return { leadingBl: '' };
}

async function resolveList(
  http: SPHttpClient,
  siteUrl: string
): Promise<{ siteUrl: string; listTitle: string; columns: IListColumns } | undefined> {
  const webUrl = trimSlash(siteUrl);
  if (cached && cached.siteUrl === webUrl) {
    return cached;
  }

  for (let i = 0; i < LIST_TITLES.length; i++) {
    const listTitle = LIST_TITLES[i];
    try {
      const columns = await readColumns(http, webUrl, listTitle);
      cached = { siteUrl: webUrl, listTitle, columns };
      return cached;
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      if (isMissingListError(message)) {
        continue;
      }
      throw err;
    }
  }

  return undefined;
}

async function readColumns(
  http: SPHttpClient,
  siteUrl: string,
  listTitle: string
): Promise<IListColumns> {
  const url =
    `${trimSlash(siteUrl)}/_api/web/lists/GetByTitle('${escapeOData(listTitle)}')` +
    `/fields?$select=InternalName,Title,TypeAsString&$top=500`;
  const response = await http.get(url, SPHttpClient.configurations.v1, {
    headers: {
      Accept: 'application/json;odata=nometadata'
    }
  });
  if (response.status === 404) {
    throw new Error(`List "${listTitle}" was not found on ${siteUrl}.`);
  }
  await ensureOk(response, `Could not read columns for "${listTitle}".`);
  const json = await response.json() as { value?: IFieldMeta[] };
  const fields = json.value || [];

  const projectField = pickProjectNoField(fields);
  const leadingField = pickLeadingBlField(fields);
  if (!projectField || !projectField.InternalName) {
    throw new Error(`Could not find Project No on "${listTitle}".`);
  }
  if (!leadingField || !leadingField.InternalName) {
    throw new Error(`Could not find Leading Business Line on "${listTitle}".`);
  }

  return {
    projectNoInternal: String(projectField.InternalName),
    projectNoType: projectField.TypeAsString ? String(projectField.TypeAsString) : 'Text',
    leadingBlInternal: String(leadingField.InternalName),
    leadingBlType: leadingField.TypeAsString ? String(leadingField.TypeAsString) : 'Text'
  };
}

async function queryMatchingItem(
  http: SPHttpClient,
  siteUrl: string,
  listTitle: string,
  columns: IListColumns,
  projectNo: string
): Promise<{ [key: string]: unknown } | undefined> {
  try {
    return await queryByCaml(http, siteUrl, listTitle, columns, projectNo);
  } catch (err) {
    const message = err instanceof Error ? err.message : String(err);
    if (isThresholdError(message)) {
      throw err;
    }
  }

  return queryByRestFilter(http, siteUrl, listTitle, columns, projectNo);
}

async function queryByCaml(
  http: SPHttpClient,
  siteUrl: string,
  listTitle: string,
  columns: IListColumns,
  projectNo: string
): Promise<{ [key: string]: unknown } | undefined> {
  const camlType = isNumericType(columns.projectNoType) ? 'Number' : 'Text';
  const viewXml =
    `<View Scope="RecursiveAll">` +
    `<Query><Where><Eq>` +
    `<FieldRef Name="${columns.projectNoInternal}"/>` +
    `<Value Type="${camlType}">${escapeXml(projectNo)}</Value>` +
    `</Eq></Where></Query>` +
    `<ViewFields>` +
    `<FieldRef Name="ID"/>` +
    `<FieldRef Name="${columns.projectNoInternal}"/>` +
    `<FieldRef Name="${columns.leadingBlInternal}"/>` +
    `</ViewFields>` +
    `<RowLimit>1</RowLimit>` +
    `</View>`;

  const url =
    `${trimSlash(siteUrl)}/_api/web/lists/GetByTitle('${escapeOData(listTitle)}')/GetItems`;
  const response = await http.post(url, SPHttpClient.configurations.v1, {
    headers: {
      Accept: 'application/json;odata=nometadata',
      'Content-Type': 'application/json;odata=verbose',
      'odata-version': '3.0'
    },
    body: JSON.stringify({
      query: {
        __metadata: { type: 'SP.CamlQuery' },
        ViewXml: viewXml
      }
    })
  });
  await ensureOk(response, `Could not query "${listTitle}".`);
  return firstItem(await response.json());
}

async function queryByRestFilter(
  http: SPHttpClient,
  siteUrl: string,
  listTitle: string,
  columns: IListColumns,
  projectNo: string
): Promise<{ [key: string]: unknown } | undefined> {
  const filter = isNumericType(columns.projectNoType)
    ? `${columns.projectNoInternal} eq ${Number(projectNo)}`
    : `${columns.projectNoInternal} eq '${escapeOData(projectNo)}'`;
  const isLookup = isLookupType(columns.leadingBlType);
  const selectParts: string[] = [
    'Id',
    columns.projectNoInternal,
    isLookup ? `${columns.leadingBlInternal}/Title` : columns.leadingBlInternal
  ];
  let url =
    `${trimSlash(siteUrl)}/_api/web/lists/GetByTitle('${escapeOData(listTitle)}')` +
    `/items?$filter=${encodeURIComponent(filter)}` +
    `&$select=${selectParts.join(',')}` +
    `&$top=1`;
  if (isLookup) {
    url += `&$expand=${columns.leadingBlInternal}`;
  }

  const response = await http.get(url, SPHttpClient.configurations.v1, {
    headers: {
      Accept: 'application/json;odata=nometadata'
    }
  });
  await ensureOk(response, `Could not query "${listTitle}".`);
  return firstItem(await response.json());
}

function firstItem(json: {
  value?: { [key: string]: unknown }[];
  d?: { results?: { [key: string]: unknown }[] };
}): { [key: string]: unknown } | undefined {
  const items = json.value || (json.d && json.d.results) || [];
  return items.length > 0 ? items[0] : undefined;
}

function pickProjectNoField(fields: IFieldMeta[]): IFieldMeta | undefined {
  const matches = fields.filter((field) =>
    isProjectNoColumn(String(field.Title || ''), String(field.InternalName || ''))
  );
  if (matches.length === 0) {
    return fields.filter((field) =>
      PROJECT_FIELD_FALLBACKS.indexOf(String(field.InternalName || '')) >= 0
    )[0];
  }
  const exact = matches.filter((field) => {
    const title = normalizeLabel(String(field.Title || ''));
    return title === 'project no' || title === 'project number' || title === 'projectno';
  })[0];
  return exact || matches[0];
}

function pickLeadingBlField(fields: IFieldMeta[]): IFieldMeta | undefined {
  const byInternal = fields.filter((field) => {
    const internal = String(field.InternalName || '').toLowerCase();
    return internal === LEADING_BUSINESS_LINE_INTERNAL.toLowerCase()
      || internal.indexOf('leading_x0020_business') === 0;
  })[0];
  if (byInternal) {
    return byInternal;
  }
  return fields.filter((field) =>
    /leading\s*business\s*lin/.test(String(field.Title || '').toLowerCase())
  )[0];
}

function isProjectNoColumn(title: string, internal: string): boolean {
  const t = normalizeLabel(title);
  const i = (internal || '').toLowerCase();
  if (/sub[\s-]*project/.test(t) || (i.indexOf('sub') >= 0 && i.indexOf('project') >= 0)) {
    return false;
  }
  if (
    t === 'project no' ||
    t === 'project no.' ||
    t === 'project number' ||
    t === 'projectno' ||
    t === 'project #'
  ) {
    return true;
  }
  if (
    i === 'project_x0020_no' ||
    i === 'project_x0020_number' ||
    i === 'projectno'
  ) {
    return true;
  }
  return /^project[\s._-]*(no\.?|number|#)$/.test(t);
}

function readFieldText(item: { [key: string]: unknown }, fieldName: string): string {
  const raw = item[fieldName];
  const fromValue = stringifySharePointValue(raw);
  if (fromValue) {
    return fromValue;
  }
  const lookupId = item[`${fieldName}Id`];
  if (lookupId && typeof lookupId === 'object') {
    return stringifySharePointValue(lookupId);
  }
  return '';
}

function stringifySharePointValue(raw: unknown): string {
  if (raw === null || raw === undefined) {
    return '';
  }
  if (typeof raw === 'string' || typeof raw === 'number' || typeof raw === 'boolean') {
    return String(raw).trim();
  }
  if (Array.isArray(raw)) {
    return raw.map(stringifySharePointValue).filter((part) => !!part).join(', ');
  }
  if (typeof raw === 'object') {
    const record = raw as { [key: string]: unknown };
    const titled = stringifySharePointValue(record.Title || record.Name || record.Value);
    if (titled) {
      return titled;
    }
    if (Array.isArray(record.results)) {
      return stringifySharePointValue(record.results);
    }
  }
  return '';
}

function projectNumberCandidates(projectNo: string): string[] {
  const results: string[] = [];
  const push = (value: string): void => {
    if (value && results.indexOf(value) < 0) {
      results.push(value);
    }
  };
  push(projectNo);
  if (projectNo.length < 8) {
    push(('00000000' + projectNo).slice(-8));
  }
  const unpadded = projectNo.replace(/^0+/, '') || '0';
  push(unpadded);
  return results;
}

function isNumericType(typeAsString: string): boolean {
  return /^(number|integer|counter|currency)$/i.test(typeAsString || '');
}

function isLookupType(typeAsString: string): boolean {
  return /lookup/i.test(typeAsString || '');
}

function isThresholdError(message: string): boolean {
  const text = (message || '').toLowerCase();
  return text.indexOf('list view threshold') >= 0
    || text.indexOf('exceeds the list view threshold') >= 0
    || (text.indexOf('threshold') >= 0 && text.indexOf('5000') >= 0);
}

function isMissingListError(message: string): boolean {
  const text = (message || '').toLowerCase();
  return text.indexOf('not found') >= 0 || text.indexOf('does not exist') >= 0;
}

function normalizeLabel(value: string): string {
  return (value || '').trim().toLowerCase().replace(/\s+/g, ' ');
}

function trimSlash(value: string): string {
  return (value || '').replace(/\/+$/, '');
}

function escapeOData(value: string): string {
  return (value || '').replace(/'/g, "''");
}

function escapeXml(value: string): string {
  return (value || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

async function ensureOk(response: SPHttpClientResponse, fallback: string): Promise<void> {
  if (response.ok) {
    return;
  }
  throw new Error(await readSharePointError(response, fallback));
}

async function readSharePointError(response: SPHttpClientResponse, fallback: string): Promise<string> {
  try {
    const json = await response.json() as {
      error?: { message?: string | { value?: string } };
    };
    const verbose = json.error && typeof json.error.message === 'object' ? json.error.message.value : undefined;
    const simple = json.error && typeof json.error.message === 'string' ? json.error.message : undefined;
    return verbose || simple || `${fallback} (${response.status})`;
  } catch {
    return `${fallback} (${response.status})`;
  }
}
