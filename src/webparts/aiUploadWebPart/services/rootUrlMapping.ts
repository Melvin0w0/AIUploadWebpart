import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

export const ROOT_URL_MAPPING_LIST_TITLE: string = 'Root URL Mapping List';

export interface IRootUrlMappingLookup {
  listWebUrl: string;
  destinationSiteUrl: string;
  leadingBl: string;
  projectNumber: string;
  listTitle?: string;
}

export async function assertActiveRootUrlMapping(
  http: SPHttpClient,
  lookup: IRootUrlMappingLookup
): Promise<void> {
  const listTitle = (lookup.listTitle || ROOT_URL_MAPPING_LIST_TITLE).trim() || ROOT_URL_MAPPING_LIST_TITLE;
  const sitesToTry = uniqueUrls([lookup.listWebUrl, lookup.destinationSiteUrl]);
  let items: { [key: string]: unknown }[] | undefined;
  let lastError: string = '';

  for (const siteUrl of sitesToTry) {
    try {
      items = await readMappingItems(http, siteUrl, listTitle);
      break;
    } catch (err) {
      lastError = err instanceof Error ? err.message : String(err);
      if (!isMissingListError(lastError)) {
        throw err;
      }
    }
  }

  if (!items) {
    throw new Error(
      lastError ||
      `Could not find "${listTitle}". Add this list on the current site or the destination site.`
    );
  }

  const matches = items.filter((item) => itemMatchesDestination(item, lookup));
  if (matches.length === 0) {
    throw new Error(
      `No "${listTitle}" item was found for ${lookup.leadingBl || lookup.destinationSiteUrl}.`
    );
  }

  const active = matches.filter((item) => isActiveYes(readActiveValue(item)));
  if (active.length === 0) {
    throw new Error(
      `"${listTitle}" has a matching item, but IsActive is not Yes. Upload is blocked for this site.`
    );
  }
}

async function readMappingItems(
  http: SPHttpClient,
  siteUrl: string,
  listTitle: string
): Promise<{ [key: string]: unknown }[]> {
  const items: { [key: string]: unknown }[] = [];
  let url =
    `${trimSlash(siteUrl)}/_api/web/lists/GetByTitle('${escapeOData(listTitle)}')` +
    `/items?$top=5000`;

  while (url) {
    const response = await http.get(url, SPHttpClient.configurations.v1, {
      headers: {
        Accept: 'application/json;odata=nometadata'
      }
    });
    if (response.status === 404) {
      throw new Error(`List "${listTitle}" was not found on ${siteUrl}.`);
    }
    await ensureOk(response, `Could not read "${listTitle}".`);
    const json = await response.json() as {
      value?: { [key: string]: unknown }[];
      '@odata.nextLink'?: string;
    };
    (json.value || []).forEach((item) => items.push(item));
    url = json['@odata.nextLink'] || '';
  }

  return items;
}

function itemMatchesDestination(
  item: { [key: string]: unknown },
  lookup: IRootUrlMappingLookup
): boolean {
  const needles = uniqueStrings([
    lookup.destinationSiteUrl,
    sitePath(lookup.destinationSiteUrl),
    lookup.leadingBl,
    blCode(lookup.destinationSiteUrl),
    lookup.projectNumber
  ]);
  if (needles.length === 0) {
    return false;
  }

  const haystacks = collectSearchableValues(item);
  return needles.some((needle) => haystacks.some((haystack) => valuesMatch(haystack, needle)));
}

function collectSearchableValues(item: { [key: string]: unknown }): string[] {
  const values: string[] = [];
  Object.keys(item).forEach((key) => {
    if (isSystemField(key) || isActiveField(key)) {
      return;
    }
    pushSearchable(values, item[key]);
  });
  return uniqueStrings(values);
}

function pushSearchable(values: string[], raw: unknown): void {
  if (raw === null || raw === undefined) {
    return;
  }
  if (typeof raw === 'string' || typeof raw === 'number' || typeof raw === 'boolean') {
    values.push(String(raw));
    return;
  }
  if (typeof raw === 'object') {
    const record = raw as { Url?: string; Description?: string };
    if (record.Url) {
      values.push(record.Url);
    }
    if (record.Description) {
      values.push(record.Description);
    }
  }
}

function readActiveValue(item: { [key: string]: unknown }): unknown {
  const direct = Object.keys(item).filter((key) => isActiveField(key))[0];
  if (direct) {
    return item[direct];
  }
  return undefined;
}

function isActiveField(name: string): boolean {
  const key = normalizeKey(name.replace(/_x0020_/gi, ' '));
  return key === 'isactive' || key === 'is active' || key === 'active' || key.indexOf('is active') >= 0;
}

function isActiveYes(value: unknown): boolean {
  if (value === true || value === 1) {
    return true;
  }
  const text = String(value || '').trim().toLowerCase();
  return text === 'yes' || text === 'y' || text === 'true' || text === '1';
}

function valuesMatch(haystack: string, needle: string): boolean {
  const left = normalizeKey(haystack);
  const right = normalizeKey(needle);
  if (!left || !right) {
    return false;
  }
  if (left === right) {
    return true;
  }
  const leftUrl = normalizeUrl(haystack);
  const rightUrl = normalizeUrl(needle);
  if (leftUrl && rightUrl && leftUrl === rightUrl) {
    return true;
  }
  if (looksLikeUrlOrPath(leftUrl) && looksLikeUrlOrPath(rightUrl) &&
    (leftUrl.indexOf(rightUrl) >= 0 || rightUrl.indexOf(leftUrl) >= 0)) {
    return true;
  }
  return false;
}

function looksLikeUrlOrPath(value: string): boolean {
  return value.indexOf('/') >= 0 || value.indexOf('.') >= 0;
}

function sitePath(siteUrl: string): string {
  const match = trimSlash(siteUrl).match(/\/sites\/[^/]+/i);
  return match ? match[0] : '';
}

function blCode(siteUrl: string): string {
  const path = sitePath(siteUrl);
  const underscore = path.lastIndexOf('_');
  return underscore >= 0 ? path.substring(underscore + 1) : '';
}

function normalizeUrl(value: string): string {
  return (value || '')
    .trim()
    .toLowerCase()
    .replace(/^https?:\/\//, '')
    .replace(/\/+$/, '');
}

function normalizeKey(value: string): string {
  return (value || '')
    .trim()
    .toLowerCase()
    .replace(/&/g, 'and')
    .replace(/[^a-z0-9]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function isSystemField(name: string): boolean {
  return /^(odata\.|@odata|__metadata|id$|guid$|created|modified|authorid|editorid|contenttypeid)/i.test(name);
}

function isMissingListError(message: string): boolean {
  const text = (message || '').toLowerCase();
  return text.indexOf('not found') >= 0 || text.indexOf('does not exist') >= 0;
}

function uniqueUrls(urls: string[]): string[] {
  const seen: { [key: string]: boolean } = {};
  const result: string[] = [];
  urls.forEach((url) => {
    const trimmed = trimSlash(url);
    const key = trimmed.toLowerCase();
    if (!trimmed || seen[key]) {
      return;
    }
    seen[key] = true;
    result.push(trimmed);
  });
  return result;
}

function uniqueStrings(values: string[]): string[] {
  const seen: { [key: string]: boolean } = {};
  const result: string[] = [];
  values.forEach((value) => {
    const trimmed = (value || '').trim();
    const key = trimmed.toLowerCase();
    if (!trimmed || seen[key]) {
      return;
    }
    seen[key] = true;
    result.push(trimmed);
  });
  return result;
}

function trimSlash(value: string): string {
  return (value || '').replace(/\/+$/, '');
}

function escapeOData(value: string): string {
  return value.replace(/'/g, "''");
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
