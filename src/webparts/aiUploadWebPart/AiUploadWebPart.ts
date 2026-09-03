import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'AiUploadWebPartStrings';
import AiUpload from './components/AiUpload';
import { IAiUploadProps } from './components/IAiUploadProps';
import { DEFAULT_FORM_FIELDS_TEXT } from './constants/defaultFormFields';

export interface IAiUploadWebPartProps {
  formFields: string;
  azureOpenAiEndpoint: string;
  azureOpenAiApiKey: string;
  azureOpenAiDeployment: string;
  azureOpenAiApiVersion: string;
  libraryName: string;
  folderPathTemplate: string;
}

export default class AiUploadWebPart extends BaseClientSideWebPart<IAiUploadWebPartProps> {

  public render(): void {
    const tenantUrl = this._tenantUrl();
    const element: React.ReactElement<IAiUploadProps> = React.createElement(
      AiUpload,
      {
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        formFields: this._resolveFormFields(),
        azureOpenAiEndpoint: this.properties.azureOpenAiEndpoint || '',
        azureOpenAiApiKey: this.properties.azureOpenAiApiKey || '',
        azureOpenAiDeployment: this.properties.azureOpenAiDeployment || '',
        azureOpenAiApiVersion: this.properties.azureOpenAiApiVersion || '2024-08-01-preview',
        tenantUrl,
        currentWebUrl: this.context.pageContext.web.absoluteUrl,
        libraryName: this.properties.libraryName || 'Project Documents',
        folderPathTemplate: this._resolveFolderPathTemplate(),
        spHttpClient: this.context.spHttpClient
      }
    );

    ReactDom.render(element, this.domElement);
  }

  private _tenantUrl(): string {
    const absoluteUrl = this.context.pageContext.web.absoluteUrl;
    const host = absoluteUrl.replace(/^https?:\/\//i, '').split('/')[0];
    return `https://${host}`;
  }

  private _resolveFormFields(): string {
    const configured = (this.properties.formFields || '').trim();
    if (!configured || configured.indexOf('欄位') >= 0) {
      return DEFAULT_FORM_FIELDS_TEXT;
    }
    return configured;
  }

  private _resolveFolderPathTemplate(): string {
    const configured = (this.properties.folderPathTemplate || '').trim();
    if (!configured || configured === '{Sub-Project Number}/{Registration Number}') {
      return '{Project Number}';
    }
    return configured;
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.5');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('formFields', {
                  label: strings.FormFieldsFieldLabel,
                  description: strings.FormFieldsFieldDescription,
                  multiline: true,
                  rows: 18
                })
              ]
            },
            {
              groupName: strings.UploadGroupName,
              groupFields: [
                PropertyPaneTextField('libraryName', {
                  label: strings.LibraryNameLabel,
                  description: strings.LibraryNameDescription,
                  placeholder: 'Project Documents'
                }),
                PropertyPaneTextField('folderPathTemplate', {
                  label: strings.FolderPathTemplateLabel,
                  description: strings.FolderPathTemplateDescription,
                  placeholder: '{Project Number}'
                })
              ]
            },
            {
              groupName: strings.AiGroupName,
              groupFields: [
                PropertyPaneTextField('azureOpenAiEndpoint', {
                  label: strings.AzureEndpointLabel,
                  description: strings.AzureEndpointDescription
                }),
                PropertyPaneTextField('azureOpenAiDeployment', {
                  label: strings.AzureDeploymentLabel,
                  description: strings.AzureDeploymentDescription
                }),
                PropertyPaneTextField('azureOpenAiApiKey', {
                  label: strings.AzureApiKeyLabel,
                  description: strings.AzureApiKeyDescription
                }),
                PropertyPaneTextField('azureOpenAiApiVersion', {
                  label: strings.AzureApiVersionLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
