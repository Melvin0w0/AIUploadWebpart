import { SPHttpClient } from '@microsoft/sp-http';
import {
  INCOMING_FOLDER_NAME,
  OUTGOING_FOLDER_NAME
} from '../constants/incomingName';
import {
  UploadType,
  UPLOAD_TYPE_NORMAL,
  uploadTypeFromFolderName
} from '../constants/uploadType';
import { collapsePath, trimSlash } from './uploadDestination';

export async function lookupProjectUploadTypes(
  http: SPHttpClient,
  siteUrl: string,
  libraryName: string,
  projectFolderPath: string,
  correspondenceFolderName: string
): Promise<UploadType[]> {
  const webUrl = trimSlash(siteUrl);
  const library = (libraryName || '').trim();
  const projectPath = collapsePath(projectFolderPath || '');
  if (!webUrl || !library || !projectPath || !http) {
    return [UPLOAD_TYPE_NORMAL];
  }

  try {
    const libraryRoot = await readLibraryRoot(http, webUrl, library);
    if (!libraryRoot) {
      return [UPLOAD_TYPE_NORMAL];
    }

    const found: { [key: string]: boolean } = {};
    const add = (uploadType: UploadType): void => {
      found[uploadType] = true;
    };

    const projectUrl = joinServerPath(libraryRoot, projectPath);
    const projectFolders = await listFolderNames(http, webUrl, projectUrl);
    if (projectFolders) {
      add(UPLOAD_TYPE_NORMAL);
      projectFolders.forEach((name) => {
        const mapped = uploadTypeFromFolderName(name);
        if (mapped) {
          add(mapped);
        }
      });
    }

    const correspondenceUrl = joinServerPath(projectUrl, correspondenceFolderName);
    const correspondenceFolders = await listFolderNames(http, webUrl, correspondenceUrl);
    if (correspondenceFolders) {
      add(UPLOAD_TYPE_NORMAL);
      correspondenceFolders.forEach((name) => {
        const mapped = uploadTypeFromFolderName(name);
        if (mapped) {
          add(mapped);
        }
      });
    }

    const types: UploadType[] = [UPLOAD_TYPE_NORMAL];
    (['confidentialInvoice', 'confidentialMisc'] as UploadType[]).forEach((uploadType) => {
      if (found[uploadType] && types.indexOf(uploadType) < 0) {
        types.push(uploadType);
      }
    });
    return types;
  } catch {
    return [UPLOAD_TYPE_NORMAL];
  }
}

export function correspondenceFolderName(kind: 'incoming' | 'outgoing' | 'unknown' | string): string {
  if (kind === 'incoming') {
    return INCOMING_FOLDER_NAME;
  }
  if (kind === 'outgoing') {
    return OUTGOING_FOLDER_NAME;
  }
  return '';
}

async function readLibraryRoot(http: SPHttpClient, siteUrl: string, libraryTitle: string): Promise<string> {
  const url =
    `${siteUrl}/_api/web/lists/GetByTitle('${escapeOData(libraryTitle)}')` +
    `?$select=RootFolder/ServerRelativeUrl&$expand=RootFolder`;
  const response = await http.get(url, SPHttpClient.configurations.v1, {
    headers: {
      Accept: 'application/json;odata=nometadata'
    }
  });
  if (!response.ok) {
    return '';
  }
  const json = await response.json() as { RootFolder?: { ServerRelativeUrl?: string } };
  return (json.RootFolder && json.RootFolder.ServerRelativeUrl) || '';
}

async function listFolderNames(
  http: SPHttpClient,
  siteUrl: string,
  serverRelativeUrl: string
): Promise<string[] | undefined> {
  const url =
    `${siteUrl}/_api/web/GetFolderByServerRelativeUrl('${escapeOData(serverRelativeUrl)}')` +
    `/Folders?$select=Name`;
  const response = await http.get(url, SPHttpClient.configurations.v1, {
    headers: {
      Accept: 'application/json;odata=nometadata'
    }
  });
  if (response.status === 404 || !response.ok) {
    return undefined;
  }
  const json = await response.json() as {
    value?: { Name?: string }[];
    d?: { results?: { Name?: string }[] };
  };
  const items = json.value || (json.d && json.d.results) || [];
  return items
    .map((item) => (item.Name || '').trim())
    .filter((name) => name && !isIgnoredFolderName(name));
}

function isIgnoredFolderName(name: string): boolean {
  const text = (name || '').trim();
  if (!text || text.charAt(0) === '_') {
    return true;
  }
  const key = text.toLowerCase();
  return key === 'forms' || key === 'attachments' || key === 'item';
}

function joinServerPath(root: string, child: string): string {
  const left = trimSlash(root || '');
  const right = collapsePath(child || '');
  if (!right) {
    return left;
  }
  return `${left}/${right}`;
}

function escapeOData(value: string): string {
  return (value || '').replace(/'/g, "''");
}
