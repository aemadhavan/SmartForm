import { WebPartContext } from '@microsoft/sp-webpart-base';
import { FormSchema } from '../../models/IFormSchema';

export interface ISubmissionsListProps {
  context: WebPartContext;
  targetLibraryName: string;
  onEditSubmission?: (fileName: string, submissionData: any) => void;
  /**
   * Increment to force reload (e.g., after a submission is created/updated).
   */
  refreshKey?: number;
  /**
   * Used to populate labels for selected columns.
   */
  schema?: FormSchema | null;
  /**
   * Columns configured at the web part level (shared for all users).
   */
  selectedFieldColumns?: string[];
  /**
   * Page size for the list.
   */
  pageSize?: number;
  isDarkTheme: boolean;
}
