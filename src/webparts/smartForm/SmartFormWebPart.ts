import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneSlider,
  PropertyPaneTextField,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'SmartFormWebPartStrings';
import SmartForm from './components/SmartForm';
import { ISmartFormProps } from './components/ISmartFormProps';
import { PropertyPaneJsonFilePicker } from './propertyPane/PropertyPaneJsonFilePicker';
import { PropertyPaneCheckboxMultiSelect } from './propertyPane/PropertyPaneCheckboxMultiSelect';

export interface ISmartFormWebPartProps {
  targetLibraryName: string;
  fileNamePattern: string;
  persistSchema: boolean;
  showSuccessMessage: boolean;
  schemaJson: string;
  submittedItemsColumns: string[];
  submittedItemsPageSize: number;
}

export default class SmartFormWebPart extends BaseClientSideWebPart<ISmartFormWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  private getSchemaFieldOptions(): { key: string; text: string }[] {
    const schemaText = this.properties.schemaJson || '';
    if (!schemaText.trim()) {
      return [];
    }

    try {
      const parsed = JSON.parse(schemaText);
      if (!Array.isArray(parsed)) {
        return [];
      }

      return parsed
        .filter((f) => f && typeof f.fieldName === 'string')
        .map((f) => ({
          key: f.fieldName,
          text: (typeof f.fieldLabel === 'string' && f.fieldLabel) ? f.fieldLabel : f.fieldName,
        }));
    } catch {
      return [];
    }
  }

  public render(): void {
    const element: React.ReactElement<ISmartFormProps> = React.createElement(
      SmartForm,
      {
        context: this.context,
        schemaJson: this.properties.schemaJson || '',
        targetLibraryName: this.properties.targetLibraryName || 'Form Submissions',
        fileNamePattern: this.properties.fileNamePattern || 'FormSubmission_{timestamp}_{userName}.json',
        persistSchema: this.properties.persistSchema !== undefined ? this.properties.persistSchema : true,
        showSuccessMessage: this.properties.showSuccessMessage !== undefined ? this.properties.showSuccessMessage : true,
        submittedItemsColumns: this.properties.submittedItemsColumns || [],
        submittedItemsPageSize: this.properties.submittedItemsPageSize || 50,
        isDarkTheme: this._isDarkTheme
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });
  }



  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
            case 'TeamsModern':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: 'Configure the Smart Form settings'
          },
          groups: [
            {
              groupName: 'Form Schema',
              groupFields: [
                PropertyPaneJsonFilePicker('schemaJson', {
                  label: 'Import schema from file',
                  buttonText: 'Import JSON file',
                  onFileTextLoaded: (fileText: string) => {
                    this.properties.schemaJson = fileText;
                    this.context.propertyPane.refresh();
                    this.render();
                  }
                }),
                PropertyPaneTextField('schemaJson', {
                  label: 'Schema JSON',
                  description: 'Paste the JSON array that defines your form fields.',
                  multiline: true,
                  rows: 20,
                  resizable: false,
                  value: this.properties.schemaJson || ''
                })
              ]
            },
            {
              groupName: 'SharePoint Integration',
              groupFields: [
                PropertyPaneTextField('targetLibraryName', {
                  label: 'Target Document Library',
                  description: 'Name of the SharePoint library where form submissions will be saved',
                  value: this.properties.targetLibraryName || 'Form Submissions'
                }),
                PropertyPaneTextField('fileNamePattern', {
                  label: 'File Naming Pattern',
                  description: 'Use placeholders: {timestamp}, {userName}, {date}, {time}, {guid}',
                  value: this.properties.fileNamePattern || 'FormSubmission_{timestamp}_{userName}.json'
                }),
                PropertyPaneCheckboxMultiSelect('submittedItemsColumns', {
                  label: 'Submitted Items - columns (form fields)',
                  description: 'These are shared for all users. Values are stored into library columns for fast listing.',
                  options: this.getSchemaFieldOptions(),
                  selectedKeys: this.properties.submittedItemsColumns || [],
                  onSelectionChanged: (keys: string[]) => {
                    this.properties.submittedItemsColumns = keys;
                    this.context.propertyPane.refresh();
                    this.render();
                  }
                }),
                PropertyPaneSlider('submittedItemsPageSize', {
                  label: 'Submitted Items - page size',
                  min: 10,
                  max: 200,
                  step: 10,
                  value: this.properties.submittedItemsPageSize || 50,
                  showValue: true
                })
              ]
            },
            {
              groupName: 'Form Behavior',
              groupFields: [
                PropertyPaneToggle('persistSchema', {
                  label: 'Save Schema with Submission',
                  onText: 'Yes',
                  offText: 'No',
                  checked: this.properties.persistSchema !== undefined ? this.properties.persistSchema : true
                }),
                PropertyPaneToggle('showSuccessMessage', {
                  label: 'Show Success Notification',
                  onText: 'Yes',
                  offText: 'No',
                  checked: this.properties.showSuccessMessage !== undefined ? this.properties.showSuccessMessage : true
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
