import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface ISmartFormProps {
  context: WebPartContext;
  schemaJson: string;
  targetLibraryName: string;
  fileNamePattern: string;
  persistSchema: boolean;
  showSuccessMessage: boolean;
  submittedItemsColumns: string[];
  submittedItemsPageSize: number;
  showSchemaBanner: boolean;
  showDownloadJson: boolean;
  formTitle: string;
  isDarkTheme: boolean;
}
