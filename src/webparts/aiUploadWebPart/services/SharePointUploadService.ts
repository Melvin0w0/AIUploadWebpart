import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { isNameField, isRegistrationNumberField } from '../constants/defaultFormFields';
import { INamedValue, IResolvedUploadDestination } from './uploadDestination';
import { assertActiveRootUrlMapping, ROOT_URL_MAPPING_LIST_TITLE } from './rootUrlMapping';
import {
  buildFieldPayload,
  ILibraryField,
  isLibraryMetadataField,
  REGISTRATION_NUMBER_INTERNAL
} from './libraryFieldMap';

export interface ISharePointUploadResult {
  fileUrl: string;
  folderUrl: string;
  fileName: string;
  metadataError: string;
}

interface IExistingRegistrationFile {
  fileName: string;
  folderPath: string;
  fileRef: string;
}

export class DuplicateDestinationFileError extends Error {
  public readonly fileName: string;

  public constructor(fileName: string) {
    super(`A file named "${fileName}" already exists in this folder.`);
    this.name = 'DuplicateDestinationFileError';
    this.fileName = fileName;
  }
}

export class DuplicateRegistrationNumberError extends Error {
  public readonly registrationNumber: string;
  public readonly folderPath: string;
  public readonly fileName: string;

  public constructor(registrationNumber: string, folderPath: string, fileName: string) {
    const location = folderPath ? ` in "${folderPath}"` : '';
    super(`Registration Number "${registrationNumber}" has already been uploaded${location}.`);
    this.name = 'DuplicateRegistrationNumberError';
    this.registrationNumber = registrationNumber;
    this.folderPath = folderPath;
    this.fileName = fileName;
  }
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
    const projectFolderUrl = projectFolderServerUrl(libraryRoot, destination.folderPath);

    const registrationNumber = registrationNumberFromFields(fieldValues);
    onStatus('Checking Registration Number…');
    if (!registrationNumber) {
      throw new Error('Registration Number is required.');
    }
    const existing = await this._findExistingRegistration(
      siteUrl,
      destination.libraryName,
      libraryRoot,
      registrationNumber,
      projectFolderUrl,
      digest
    );
    if (existing) {
      throw new DuplicateRegistrationNumberError(
        registrationNumber,
        existing.folderPath,
        existing.fileName
      );
    }

    if (destination.folderPath) {
      onStatus('Creating folder…');
      await this._ensureFolderPath(siteUrl, libraryRoot, destination.folderPath, digest);
    }

    onStatus('Checking if the file already exists…');
    const destinationFileUrl = `${trimEnd(folderServerRelativeUrl, '/')}/${fileName}`;
    if (await this._fileExists(siteUrl, destinationFileUrl, digest)) {
      throw new DuplicateDestinationFileError(fileName);
    }

    onStatus('Uploading PDF…');
    const buffer = originalPdfBytes && originalPdfBytes.byteLength > 0
      ? copyArrayBuffer(originalPdfBytes)
      : await file.arrayBuffer();
    const addUrl =
      `${siteUrl}/_api/web/GetFolderByServerRelativeUrl('${escapeOData(folderServerRelativeUrl)}')` +
      `/Files/add(overwrite=false,url='${escapeOData(fileName)}')`;
    const uploaded = await this._http.post(addUrl, SPHttpClient.configurations.v1, {
      headers: {
        Accept: 'application/json;odata=nometadata',
        'Content-Type': 'application/octet-stream',
        'X-RequestDigest': digest
      },
      body: buffer
    });
    if (!uploaded.ok) {
      const details = await readSharePointError(uploaded, 'Could not upload the PDF to Project Documents.');
      if (uploaded.status === 409 || isAlreadyExistsMessage(details)) {
        throw new DuplicateDestinationFileError(fileName);
      }
      throw new Error(details);
    }

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

  private async _findExistingRegistration(
    siteUrl: string,
    libraryTitle: string,
    libraryRoot: string,
    registrationNumber: string,
    projectFolderUrl: string,
    digest: string
  ): Promise<IExistingRegistrationFile | undefined> {
    const libraryHit = await this._queryRegistrationNumber(
      siteUrl,
      libraryTitle,
      registrationNumber,
      '',
      digest
    );
    if (libraryHit && typeof libraryHit !== 'string') {
      return this._hydrateExistingFile(siteUrl, libraryTitle, libraryRoot, libraryHit, digest);
    }
    if (libraryHit === false) {
      return undefined;
    }

    const folderHit = await this._queryRegistrationNumber(
      siteUrl,
      libraryTitle,
      registrationNumber,
      projectFolderUrl,
      digest
    );
    if (folderHit && typeof folderHit !== 'string') {
      return this._hydrateExistingFile(siteUrl, libraryTitle, libraryRoot, folderHit, digest);
    }
    if (folderHit === false) {
      return undefined;
    }

    throw new Error(
      (typeof libraryHit === 'string' ? libraryHit : '')
      || (typeof folderHit === 'string' ? folderHit : '')
      || 'Could not check Registration Number.'
    );
  }

  private async _hydrateExistingFile(
    siteUrl: string,
    libraryTitle: string,
    libraryRoot: string,
    item: { [key: string]: unknown },
    digest: string
  ): Promise<IExistingRegistrationFile> {
    const id = itemId(item);
    const picked = pickFileFields(item);
    let serverRelativeUrl = picked.fileRef;
    let fileName = picked.fileName;
    if (id) {
      const file = await this._getListItemFile(siteUrl, libraryTitle, id, digest);
      if (file) {
        serverRelativeUrl = file.ServerRelativeUrl || serverRelativeUrl;
        fileName = file.Name || fileName;
      }
    }
    const folderServer = picked.fileDir || serverRelativeUrl.replace(/\/[^/]+$/, '');
    return {
      fileName,
      fileRef: serverRelativeUrl,
      folderPath: folderPathAfterLibrary(folderServer, libraryRoot, libraryTitle)
    };
  }

  private async _getListItemFile(
    siteUrl: string,
    libraryTitle: string,
    itemId: number,
    digest: string
  ): Promise<{ ServerRelativeUrl?: string; Name?: string } | undefined> {
    const urls = [
      `${siteUrl}/_api/web/lists/GetByTitle('${escapeOData(libraryTitle)}')/items(${itemId})` +
      `/File?$select=ServerRelativeUrl,Name`,
      `${siteUrl}/_api/web/lists/GetByTitle('${escapeOData(libraryTitle)}')/items(${itemId})` +
      `?$select=Id,FileLeafRef,FileRef,FileDirRef,File/ServerRelativeUrl,File/Name&$expand=File`
    ];
    for (let i = 0; i < urls.length; i++) {
      const response = await this._http.get(urls[i], SPHttpClient.configurations.v1, {
        headers: {
          Accept: 'application/json;odata=nometadata',
          'X-RequestDigest': digest
        }
      });
      if (!response.ok) {
        continue;
      }
      const json = await response.json() as {
        ServerRelativeUrl?: string;
        Name?: string;
        FileLeafRef?: string;
        FileRef?: string;
        FileDirRef?: string;
        File?: { ServerRelativeUrl?: string; Name?: string };
      };
      const serverRelativeUrl = json.ServerRelativeUrl
        || (json.File && json.File.ServerRelativeUrl)
        || json.FileRef
        || json.FileDirRef
        || '';
      const name = json.Name || (json.File && json.File.Name) || json.FileLeafRef || '';
      if (serverRelativeUrl || name) {
        return {
          ServerRelativeUrl: serverRelativeUrl,
          Name: name
        };
      }
    }
    return undefined;
  }

  private async _queryRegistrationNumber(
    siteUrl: string,
    libraryTitle: string,
    registrationNumber: string,
    folderServerRelativeUrl: string,
    digest: string
  ): Promise<{ [key: string]: unknown } | false | string> {
    const caml = await this._queryRegistrationByCaml(
      siteUrl,
      libraryTitle,
      registrationNumber,
      folderServerRelativeUrl,
      digest
    );
    if (caml !== undefined && typeof caml !== 'string') {
      return caml;
    }
    const rest = await this._queryRegistrationByRest(
      siteUrl,
      libraryTitle,
      registrationNumber,
      digest
    );
    if (rest !== undefined && typeof rest !== 'string') {
      return rest;
    }
    return (typeof caml === 'string' ? caml : '') || (typeof rest === 'string' ? rest : '');
  }

  private async _queryRegistrationByCaml(
    siteUrl: string,
    libraryTitle: string,
    registrationNumber: string,
    folderServerRelativeUrl: string,
    digest: string
  ): Promise<{ [key: string]: unknown } | false | string> {
    const viewXml =
      `<View Scope="RecursiveAll">` +
      `<Query><Where><Eq>` +
      `<FieldRef Name="${REGISTRATION_NUMBER_INTERNAL}"/>` +
      `<Value Type="Text">${escapeXml(registrationNumber)}</Value>` +
      `</Eq></Where></Query>` +
      `<ViewFields>` +
      `<FieldRef Name="ID"/>` +
      `<FieldRef Name="FileLeafRef"/>` +
      `<FieldRef Name="FileRef"/>` +
      `<FieldRef Name="FileDirRef"/>` +
      `<FieldRef Name="${REGISTRATION_NUMBER_INTERNAL}"/>` +
      `</ViewFields>` +
      `<RowLimit>1</RowLimit>` +
      `</View>`;
    const url =
      `${siteUrl}/_api/web/lists/GetByTitle('${escapeOData(libraryTitle)}')/GetItems`;
    const query: { [key: string]: unknown } = {
      __metadata: { type: 'SP.CamlQuery' },
      ViewXml: viewXml
    };
    if (folderServerRelativeUrl) {
      query.FolderServerRelativeUrl = folderServerRelativeUrl;
    }
    const response = await this._http.post(url, SPHttpClient.configurations.v1, {
      headers: {
        Accept: 'application/json;odata=nometadata',
        'Content-Type': 'application/json;odata=verbose',
        'odata-version': '3.0',
        'X-RequestDigest': digest
      },
      body: JSON.stringify({ query })
    });
    if (response.ok) {
      const item = firstListItem(await response.json());
      return item || false;
    }
    const details = await readSharePointError(response, 'Could not check Registration Number.');
    if (isMissingFolderError(details, response.status)) {
      return false;
    }
    if (isListThresholdError(details)) {
      return details;
    }
    throw new Error(details);
  }

  private async _queryRegistrationByRest(
    siteUrl: string,
    libraryTitle: string,
    registrationNumber: string,
    digest: string
  ): Promise<{ [key: string]: unknown } | false | string> {
    const url =
      `${siteUrl}/_api/web/lists/GetByTitle('${escapeOData(libraryTitle)}')` +
      `/items?$filter=${encodeURIComponent(`${REGISTRATION_NUMBER_INTERNAL} eq '${escapeOData(registrationNumber)}'`)}` +
      `&$select=Id,FileLeafRef,FileRef,FileDirRef,File/ServerRelativeUrl,File/Name` +
      `&$expand=File` +
      `&$top=1`;
    const response = await this._http.get(url, SPHttpClient.configurations.v1, {
      headers: {
        Accept: 'application/json;odata=nometadata',
        'X-RequestDigest': digest
      }
    });
    if (response.ok) {
      const item = firstListItem(await response.json());
      return item || false;
    }
    const details = await readSharePointError(response, 'Could not check Registration Number.');
    if (isListThresholdError(details)) {
      return details;
    }
    throw new Error(details);
  }

  private async _fileExists(siteUrl: string, serverRelativeUrl: string, digest: string): Promise<boolean> {
    const urls = [
      `${siteUrl}/_api/web/GetFileByServerRelativePath(decodedUrl='${escapeOData(serverRelativeUrl)}')` +
      `?$select=Exists,Name`,
      `${siteUrl}/_api/web/GetFileByServerRelativeUrl('${escapeOData(serverRelativeUrl)}')` +
      `?$select=Exists,Name`
    ];
    let lastError = 'Could not check whether the file already exists.';
    for (let i = 0; i < urls.length; i++) {
      const response = await this._http.get(urls[i], SPHttpClient.configurations.v1, {
        headers: {
          Accept: 'application/json;odata=nometadata',
          'X-RequestDigest': digest
        }
      });
      if (response.status === 404) {
        return false;
      }
      if (response.ok) {
        const json = await response.json() as { Exists?: boolean };
        return json.Exists !== false;
      }
      lastError = await readSharePointError(response, lastError);
      if (isMissingFileError(lastError, response.status)) {
        return false;
      }
      if (response.status === 401 || response.status === 403) {
        throw new Error(lastError);
      }
    }
    throw new Error(lastError);
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

    // A single bad field (usually Issue Date format) must not drop the rest.
    // Retry the fields that SharePoint accepted, then MERGE whatever is still missing.
    if (validated.attempted && validated.failed.length > 0) {
      const rest = omitKeys(payload, validated.failed);
      if (Object.keys(rest).length > 0) {
        const retry = await this._validateUpdateListItem(siteUrl, libraryTitle, item.Id, rest, digest);
        if (retry.attempted && retry.failed.length === 0) {
          const leftover = pickKeys(payload, validated.failed);
          const leftoverFailed = await this._mergeLibraryFields(siteUrl, libraryTitle, item.Id, leftover, digest);
          if (leftoverFailed.length === 0) {
            return '';
          }
          return `The file was uploaded. Some library fields could not be updated: ${leftoverFailed.join(', ')}.`;
        }
      }
    }

    const mergeFailed = await this._mergeLibraryFields(siteUrl, libraryTitle, item.Id, payload, digest);
    if (mergeFailed.length === 0) {
      return '';
    }
    if (mergeFailed.length < Object.keys(payload).length) {
      return `The file was uploaded. Some library fields could not be updated: ${mergeFailed.join(', ')}.`;
    }
    const extra = validated.failed.length > 0 ? ` ${validated.failed.join(', ')}.` : '';
    return `The file was uploaded, but library fields were not updated.${extra}`;
  }

  private async _mergeLibraryFields(
    siteUrl: string,
    libraryTitle: string,
    itemId: number,
    payload: { [name: string]: string | number | boolean },
    digest: string
  ): Promise<string[]> {
    const keys = Object.keys(payload);
    if (keys.length === 0) {
      return [];
    }
    const mergeUrl = `${siteUrl}/_api/web/lists/GetByTitle('${escapeOData(libraryTitle)}')/items(${itemId})`;
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
      return [];
    }

    const failed: string[] = [];
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
    return failed;
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
    return (json.value || []).filter(isLibraryMetadataField);
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
  const dateText = toValidateDateText(value);
  if (dateText) {
    return dateText;
  }
  return String(value);
}

function toValidateDateText(value: string | number | boolean): string {
  if (typeof value !== 'string') {
    return '';
  }
  const iso = value.match(/^(\d{4})-(\d{2})-(\d{2})(?:[T\s].*)?$/);
  if (!iso) {
    return '';
  }
  // ValidateUpdateListItem rejects ISO with milliseconds/Z, which was aborting the other fields.
  return `${iso[1]}-${iso[2]}-${iso[3]} 00:00:00`;
}

function omitKeys(
  payload: { [name: string]: string | number | boolean },
  names: string[]
): { [name: string]: string | number | boolean } {
  const skip: { [name: string]: boolean } = {};
  names.forEach((name) => {
    skip[name.toLowerCase()] = true;
  });
  const next: { [name: string]: string | number | boolean } = {};
  Object.keys(payload).forEach((key) => {
    if (!skip[key.toLowerCase()]) {
      next[key] = payload[key];
    }
  });
  return next;
}

function pickKeys(
  payload: { [name: string]: string | number | boolean },
  names: string[]
): { [name: string]: string | number | boolean } {
  const keep: { [name: string]: boolean } = {};
  names.forEach((name) => {
    keep[name.toLowerCase()] = true;
  });
  const next: { [name: string]: string | number | boolean } = {};
  Object.keys(payload).forEach((key) => {
    if (keep[key.toLowerCase()]) {
      next[key] = payload[key];
    }
  });
  return next;
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

function isAlreadyExistsMessage(text: string): boolean {
  const value = (text || '').toLowerCase();
  return value.indexOf('already exists') >= 0
    || value.indexOf('already exist') >= 0
    || value.indexOf('name already being used') >= 0;
}

function isMissingFileError(text: string, status: number): boolean {
  if (status === 404) {
    return true;
  }
  const value = (text || '').toLowerCase();
  return value.indexOf('does not exist') >= 0
    || value.indexOf('cannot find') >= 0
    || value.indexOf('file not found') >= 0
    || value.indexOf('not found') >= 0;
}

function isMissingFolderError(text: string, status: number): boolean {
  if (status === 404) {
    return true;
  }
  const value = (text || '').toLowerCase();
  return value.indexOf('folder') >= 0 && (
    value.indexOf('does not exist') >= 0
    || value.indexOf('cannot find') >= 0
    || value.indexOf('not found') >= 0
  );
}

function isListThresholdError(text: string): boolean {
  const value = (text || '').toLowerCase();
  return value.indexOf('threshold') >= 0
    || value.indexOf('lookup column') >= 0
    || value.indexOf('exceeds the list view') >= 0;
}

function firstListItem(json: {
  value?: { [key: string]: unknown }[];
  d?: {
    results?: { [key: string]: unknown }[];
    GetItems?: { results?: { [key: string]: unknown }[] };
  };
}): { [key: string]: unknown } | undefined {
  const items = json.value
    || (json.d && json.d.results)
    || (json.d && json.d.GetItems && json.d.GetItems.results)
    || [];
  return items.length > 0 ? items[0] : undefined;
}

function itemId(item: { [key: string]: unknown }): number {
  const raw = item.Id !== undefined ? item.Id : item.ID;
  const id = typeof raw === 'number' ? raw : parseInt(String(raw || ''), 10);
  return isNaN(id) ? 0 : id;
}

function pickFileFields(item: { [key: string]: unknown }): { fileName: string; fileRef: string; fileDir: string } {
  const file = item.File && typeof item.File === 'object'
    ? item.File as { ServerRelativeUrl?: string; Name?: string }
    : undefined;
  return {
    fileName: stringField(item, ['FileLeafRef', 'Name']) || (file && file.Name) || '',
    fileRef: stringField(item, ['FileRef', 'File_x0020_Ref']) || (file && file.ServerRelativeUrl) || '',
    fileDir: stringField(item, ['FileDirRef', 'File_x0020_DirRef'])
  };
}

function stringField(item: { [key: string]: unknown }, names: string[]): string {
  for (let i = 0; i < names.length; i++) {
    const value = item[names[i]];
    if (typeof value === 'string' && value.trim()) {
      return value.trim();
    }
  }
  const wanted = names.map((name) => name.toLowerCase());
  const key = Object.keys(item).filter((itemKey) => wanted.indexOf(itemKey.toLowerCase()) >= 0)[0];
  const matched = key ? item[key] : undefined;
  return typeof matched === 'string' ? matched.trim() : '';
}

function folderPathAfterLibrary(folderServerUrl: string, libraryRoot: string, libraryTitle: string): string {
  const folderParts = decodePath(folderServerUrl).split('/').filter((segment) => segment.length > 0);
  const rootParts = decodePath(libraryRoot).split('/').filter((segment) => segment.length > 0);
  if (rootParts.length > 0 && startsWithPath(folderParts, rootParts)) {
    return folderParts.slice(rootParts.length).join('/');
  }

  const libraryNames = [
    libraryTitle,
    (libraryTitle || '').replace(/\s+/g, ''),
    'Project Documents',
    'ProjectDocuments'
  ].filter((name) => !!name);
  for (let i = 0; i < libraryNames.length; i++) {
    const key = libraryNames[i].toLowerCase();
    const index = folderParts.map((segment) => segment.toLowerCase()).indexOf(key);
    if (index >= 0) {
      return folderParts.slice(index + 1).join('/');
    }
  }
  return folderParts.slice(-3).join('/') || decodePath(folderServerUrl).replace(/^\/+/, '');
}

function startsWithPath(parts: string[], prefix: string[]): boolean {
  if (prefix.length === 0 || parts.length < prefix.length) {
    return false;
  }
  for (let i = 0; i < prefix.length; i++) {
    if (parts[i].toLowerCase() !== prefix[i].toLowerCase()) {
      return false;
    }
  }
  return true;
}

function decodePath(value: string): string {
  const raw = (value || '').replace(/\\/g, '/');
  try {
    return decodeURIComponent(raw);
  } catch {
    return raw;
  }
}

function registrationNumberFromFields(fieldValues: INamedValue[]): string {
  const fromRegistration = fieldValues.filter((field) => isRegistrationNumberField(field.label))[0];
  const fromName = fieldValues.filter((field) => isNameField(field.label))[0];
  return ((fromRegistration && fromRegistration.value) || (fromName && fromName.value) || '').trim();
}

function projectFolderServerUrl(libraryRoot: string, folderPath: string): string {
  const projectFolder = (folderPath || '').split('/').filter((segment) => segment.length > 0)[0] || '';
  if (!projectFolder) {
    return libraryRoot;
  }
  return `${trimEnd(libraryRoot, '/')}/${projectFolder}`;
}

function escapeXml(value: string): string {
  return (value || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
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
