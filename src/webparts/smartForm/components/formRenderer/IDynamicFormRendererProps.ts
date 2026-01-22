import { FormSchema } from '../../models/IFormSchema';
import { IFormData, FormFieldValue } from '../../models/IFormData';
import { IValidationError } from '../../models/IValidationError';
import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IDynamicFormRendererProps {
  schema: FormSchema;
  formData: IFormData;
  validationErrors: IValidationError[];
  onFieldChange: (fieldName: string, value: FormFieldValue) => void;
  onSubmit: () => void;
  onReset: () => void;
  onDownloadJson: () => void;
  onSaveAsCopy?: () => void;
  isSubmitting: boolean;
  isEditing?: boolean;
  showDownloadJson: boolean;
  formTitle: string;
  context: WebPartContext;
}
