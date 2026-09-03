import { SPHttpClient } from '@microsoft/sp-http';

export interface IAiUploadProps {
  hasTeamsContext: boolean;
  formFields: string;
  azureOpenAiEndpoint: string;
  azureOpenAiApiKey: string;
  azureOpenAiDeployment: string;
  azureOpenAiApiVersion: string;
  tenantUrl: string;
  currentWebUrl: string;
  libraryName: string;
  folderPathTemplate: string;
  spHttpClient: SPHttpClient;
}
