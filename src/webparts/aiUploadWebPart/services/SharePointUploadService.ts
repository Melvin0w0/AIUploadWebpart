import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { INamedValue, IResolvedUploadDestination } from './uploadDestination';
import { assertActiveRootUrlMapping, ROOT_URL_MAPPING_LIST_TITLE } from './rootUrlMapping';
import {
  buildFieldPayload,
  ILibraryField,
  isWritableLibraryField
} from './libraryFieldMap';

export interface ISharePointUploadResult {
  fileUrl: string;
  folderUrl: string;
  fileName: string;
  metadataError: string;
}

export class SharePointUploadService {
  public constructor(private readonly _http: SPHttpClient) {
  }

  public async uploadPdf(
    destination: IResolvedUploadDestination,
    file: File,
    fileName: string,
    onStatus: (status: string) => void,
    mappingWebUrl: string,
    fieldValues: INamedValue[],
    originalPdfBytes?: Uint8Array
  ): Promise<ISharePointUploadResult> {
    const siteUrl = destination.siteUrl;
    onStatus('Checking Root URL Mapping List…');
    await assertActiveRootUrlMapping(this._http, {
      listWebUrl: mappingWebUrl,
      destinationSiteUrl: siteUrl,
      leadingBl: destination.leadingBl,
      projectNumber: destination.projectNumber,
      listTitle: ROOT_URL_MAPPING_LIST_TITLE
    });

    onStatus('Reading destination site…');
    const digest = await this._getDigest(siteUrl);
    onStatus('Finding Project Documents…');
    const libraryRoot = await this._getLibraryRoot(siteUrl, destination.libraryName, digest);
    const folderServerRelativeUrl = destination.folderPath
      ? `${trimEnd(libraryRoot, '/')}/${destination.folderPath}`
      : libraryRoot;

    if (destination.folderPath) {
      onStatus('Creating folder…');
      await this._ensureFolderPath(siteUrl, libraryRoot, destination.folderPath, digest);
    }

    onStatus('Uploading PDF…');
    const buffer = originalPdfBytes && originalPdfBytes.byteLength > 0
      ? copyArrayBuffer(originalPdfBytes)
      : await file.arrayBuffer();
    const addUrl =
      `${siteUrl}/_api/web/GetFolderByServerRelativeUrl('${escapeOData(folderServerRelativeUrl)}')` +
      `/Files/add(overwrite=true,url='${escapeOData(fileName)}')`;
    const uploaded = await this._http.post(addUrl, SPHttpClient.configurations.v1, {
      headers: {
        Accept: 'application/json;odata=nometadata',
        'Content-Type': 'application/octet-stream',
        'X-RequestDigest': digest
      },
      body: buffer
    });
    await this._ensureOk(uploaded, 'Could not upload the PDF to Project Documents.');

    const payload = await uploaded.json() as {
      ServerRelativeUrl?: string;
      LinkingUri?: string;
      UniqueId?: string;
    };
    const serverRelativeUrl = payload.ServerRelativeUrl || `${folderServerRelativeUrl}/${fileName}`;
    const origin = siteOrigin(siteUrl);

    onStatus('Updating library fields…');
    let metadataError = '';
    try {
      metadataError = await this._updateLibraryFields(
        siteUrl,
        destination.libraryName,
        serverRelativeUrl,
        payload.UniqueId,
        fieldValues,
        digest
      );
    } catch (err) {
      const details = err instanceof Error ? err.message : 'Could not update library fields.';
      metadataError = `The file was uploaded, but library fields were not updated. ${details}`;
    }

    return {
      fileName,
      fileUrl: payload.LinkingUri || `${origin}${serverRelativeUrl}`,
      folderUrl: buildLibraryFolderUrl(origin, libraryRoot, folderServerRelativeUrl),
      metadataError
    };
  }

  private async _getDigest(siteUrl: string): Promise<string> {
    const response = await this._http.post(
      `${siteUrl}/_api/contextinfo`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: 'application/json;odata=nometadata'
        }
      }
    );
    await this._ensureOk(response, 'Could not reach the project site. Check Project Number and your access.');
    const json = await response.json() as { FormDigestValue?: string };
    if (!json.FormDigestValue) {
      throw new Error('Could not get a SharePoint request digest for the project site.');
    }
    return json.FormDigestValue;
  }

  private async _getLibraryRoot(siteUrl: string, libraryTitle: string, digest: string): Promise<string> {
    const url =
      `${siteUrl}/_api/web/lists/GetByTitle('${escapeOData(libraryTitle)}')` +
      `?$select=RootFolder/ServerRelativeUrl&$expand=RootFolder`;
    const response = await this._http.get(url, SPHttpClient.configurations.v1, {
      headers: {
        Accept: 'application/json;odata=nometadata',
        'X-RequestDigest': digest
      }
    });
    await this._ensureOk(response, `Could not find the "${libraryTitle}" library on the project site.`);
    const json = await response.json() as { RootFolder?: { ServerRelativeUrl?: string } };
    const root = json.RootFolder && json.RootFolder.ServerRelativeUrl;
    if (!root) {
      throw new Error(`Could not find the "${libraryTitle}" library on the project site.`);
    }
    return root;
  }

  private async _ensureFolderPath(
    siteUrl: string,
    libraryRoot: string,
    folderPath: string,
    digest: string
  ): Promise<void> {
    const segments = folderPath.split('/').filter((segment) => segment.length > 0);
    let current = libraryRoot;
    for (const segment of segments) {
      const next = `${trimEnd(current, '/')}/${segment}`;
      if (!(await this._folderExists(siteUrl, next, digest))) {
        await this._createFolder(siteUrl, current, segment, digest);
      }
      current = next;
    }
  }

  private async _folderExists(siteUrl: string, serverRelativeUrl: string, digest: string): Promise<boolean> {
    const url =
      `${siteUrl}/_api/web/GetFolderByServerRelativeUrl('${escapeOData(serverRelativeUrl)}')` +
      `?$select=Exists`;
    const response = await this._http.get(url, SPHttpClient.configurations.v1, {
      headers: {
        Accept: 'application/json;odata=nometadata',
        'X-RequestDigest': digest
      }
    });
    if (response.status === 404) {
      return false;
    }
    if (!response.ok) {
      return false;
    }
    const json = await response.json() as { Exists?: boolean };
    return json.Exists !== false;
  }

  private async _createFolder(
    siteUrl: string,
    parentServerRelativeUrl: string,
    name: string,
    digest: string
  ): Promise<void> {
    const url =
      `${siteUrl}/_api/web/GetFolderByServerRelativeUrl('${escapeOData(parentServerRelativeUrl)}')` +
      `/Folders/add('${escapeOData(name)}')`;
    const response = await this._http.post(url, SPHttpClient.configurations.v1, {
      headers: {
        Accept: 'application/json;odata=nometadata',
        'Content-Type': 'application/json;odata=nometadata',
        'X-RequestDigest': digest
      }
    });
    if (response.ok || response.status === 409) {
      return;
    }
    await this._ensureOk(response, `Could not create folder "${name}".`);
  }

  private async _updateLibraryFields(
    siteUrl: string,
    libraryTitle: string,
    serverRelativeUrl: string,
    uniqueId: string | undefined,
    fieldValues: INamedValue[],
    digest: string
  ): Promise<string> {
    const item = await this._getListItem(siteUrl, serverRelativeUrl, uniqueId, digest);
    if (!item.Id) {
      return 'The file was uploaded, but its library item could not be found.';
    }
    const columns = await this._getWritableFields(siteUrl, libraryTitle, digest);
    const payload = buildFieldPayload(fieldValues, columns);
    if (Object.keys(payload).length === 0) {
      return 'The file was uploaded, but no matching library columns were found for the form fields.';
    }

    const validated = await this._validateUpdateListItem(
      siteUrl,
      libraryTitle,
      item.Id,
      payload,
      digest
    );
    if (validated.attempted && validated.failed.length === 0) {
      return '';
    }
    if (validated.attempted && validated.failed.length < Object.keys(payload).length) {
      return `The file was uploaded. Some library fields could not be updated: ${validated.failed.join(', ')}.`;
    }

    const mergeUrl = `${siteUrl}/_api/web/lists/GetByTitle('${escapeOData(libraryTitle)}')/items(${item.Id})`;
    const response = await this._http.post(mergeUrl, SPHttpClient.configurations.v1, {
      headers: {
        Accept: 'application/json;odata=nometadata',
        'Content-Type': 'application/json;odata=nometadata',
        'X-HTTP-Method': 'MERGE',
        'IF-MATCH': '*',
        'X-RequestDigest': digest
      },
      body: JSON.stringify(payload)
    });
    if (response.ok || response.status === 204) {
      return '';
    }

    const details = await readSharePointError(response, 'Could not update library fields.');
    const failed: string[] = [];
    const keys = Object.keys(payload);
    for (const key of keys) {
      const single: { [name: string]: string | number | boolean } = {};
      single[key] = payload[key];
      const one = await this._http.post(mergeUrl, SPHttpClient.configurations.v1, {
        headers: {
          Accept: 'application/json;odata=nometadata',
          'Content-Type': 'application/json;odata=nometadata',
          'X-HTTP-Method': 'MERGE',
          'IF-MATCH': '*',
          'X-RequestDigest': digest
        },
        body: JSON.stringify(single)
      });
      if (!(one.ok || one.status === 204)) {
        failed.push(key);
      }
    }
    if (failed.length === 0) {
      return '';
    }
    if (failed.length < keys.length) {
      return `The file was uploaded. Some library fields could not be updated: ${failed.join(', ')}.`;
    }
    const extra = validated.failed.length > 0 ? ` ${validated.failed.join(', ')}.` : '';
    return `The file was uploaded, but library fields were not updated. ${details}${extra}`;
  }

  private async _validateUpdateListItem(
    siteUrl: string,
    libraryTitle: string,
    itemId: number,
    payload: { [name: string]: string | number | boolean },
    digest: string
  ): Promise<{ attempted: boolean; failed: string[] }> {
    const formValues = Object.keys(payload).map((name) => ({
      FieldName: name,
      FieldValue: toValidateFieldValue(payload[name])
    }));
    const url =
      `${siteUrl}/_api/web/lists/GetByTitle('${escapeOData(libraryTitle)}')` +
      `/items(${itemId})/ValidateUpdateListItem()`;
    const response = await this._http.post(url, SPHttpClient.configurations.v1, {
      headers: {
        Accept: 'application/json;odata=nometadata',
        'Content-Type': 'application/json;odata=nometadata',
        'X-RequestDigest': digest
      },
      body: JSON.stringify({
        formValues,
        bNewDocumentUpdate: true
      })
    });
    if (!response.ok) {
      return { attempted: false, failed: [] };
    }
    const json = await response.json() as {
      value?: { FieldName?: string; ErrorMessage?: string; HasException?: boolean }[];
    };
    const failed = (json.value || [])
      .filter((row) => row.HasException)
      .map((row) => row.FieldName || row.ErrorMessage || 'field');
    return { attempted: true, failed };
  }

  private async _getListItem(
    siteUrl: string,
    serverRelativeUrl: string,
    uniqueId: string | undefined,
    digest: string
  ): Promise<{ Id?: number }> {
    const urls: string[] = [];
    if (uniqueId) {
      const guid = extractGuid(uniqueId);
      if (guid) {
        urls.push(`${siteUrl}/_api/web/GetFileById('${escapeOData(guid)}')/ListItemAllFields?$select=Id`);
      }
    }
    urls.push(
      `${siteUrl}/_api/web/GetFileByServerRelativePath(decodedUrl='${escapeOData(serverRelativeUrl)}')` +
      `/ListItemAllFields?$select=Id`
    );
    urls.push(
      `${siteUrl}/_api/web/GetFileByServerRelativeUrl('${escapeOData(serverRelativeUrl)}')` +
      `/ListItemAllFields?$select=Id`
    );

    let lastError = 'Could not read the uploaded file item.';
    for (let i = 0; i < urls.length; i++) {
      const response = await this._http.get(urls[i], SPHttpClient.configurations.v1, {
        headers: {
          Accept: 'application/json;odata=nometadata',
          'X-RequestDigest': digest
        }
      });
      if (response.ok) {
        return await response.json() as { Id?: number };
      }
      lastError = await readSharePointError(response, lastError);
    }
    throw new Error(lastError);
  }

  private async _getWritableFields(
    siteUrl: string,
    libraryTitle: string,
    digest: string
  ): Promise<ILibraryField[]> {
    const url =
      `${siteUrl}/_api/web/lists/GetByTitle('${escapeOData(libraryTitle)}')/fields` +
      `?$select=InternalName,Title,TypeAsString,Hidden,ReadOnlyField`;
    const response = await this._http.get(url, SPHttpClient.configurations.v1, {
      headers: {
        Accept: 'application/json;odata=nometadata',
        'X-RequestDigest': digest
      }
    });
    await this._ensureOk(response, 'Could not read library columns.');
    const json = await response.json() as { value?: ILibraryField[] };
    return (json.value || []).filter(isWritableLibraryField);
  }

  private async _ensureOk(response: SPHttpClientResponse, fallback: string): Promise<void> {
    if (response.ok) {
      return;
    }
    throw new Error(await readSharePointError(response, fallback));
  }
}

function copyArrayBuffer(bytes: Uint8Array): ArrayBuffer {
  const copy = new Uint8Array(bytes.byteLength);
  copy.set(bytes);
  return copy.buffer;
}

function extractGuid(value: string): string {
  const match = (value || '').match(/[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}/);
  return match ? match[0] : '';
}

function toValidateFieldValue(value: string | number | boolean): string {
  if (typeof value === 'boolean') {
    return value ? '1' : '0';
  }
  return String(value);
}

function escapeOData(value: string): string {
  return value.replace(/'/g, "''");
}

function trimEnd(value: string, char: string): string {
  return value.lastIndexOf(char) === value.length - 1 ? value.substring(0, value.length - 1) : value;
}

function siteOrigin(siteUrl: string): string {
  const match = siteUrl.match(/^https?:\/\/[^/]+/i);
  return match ? match[0] : siteUrl;
}

function buildLibraryFolderUrl(origin: string, libraryRoot: string, folderServerRelativeUrl: string): string {
  const encodedRoot = (libraryRoot || '')
    .split('/')
    .map((segment) => segment ? encodeURIComponent(segment) : '')
    .join('/');
  return `${origin}${encodedRoot}/Forms/AllItems.aspx?id=${encodeURIComponent(folderServerRelativeUrl || libraryRoot)}`;
}

async function readSharePointError(response: SPHttpClientResponse, fallback: string): Promise<string> {
  try {
    const json = await response.json() as {
      error?: { message?: string | { value?: string } };
      'odata.error'?: { message?: { value?: string } };
    };
    const verbose = json.error && typeof json.error.message === 'object' ? json.error.message.value : undefined;
    const simple = json.error && typeof json.error.message === 'string' ? json.error.message : undefined;
    const odata = json['odata.error'] && json['odata.error'].message ? json['odata.error'].message.value : undefined;
    return verbose || simple || odata || `${fallback} (${response.status})`;
  } catch {
    return `${fallback} (${response.status})`;
  }
}
