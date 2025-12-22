import { IFormFieldSchema } from '../../models/IFormSchema';
import { FormFieldValue } from '../../models/IFormData';
import { IValidationError } from '../../models/IValidationError';
import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IFieldRendererProps {
  field: IFormFieldSchema;
  value: FormFieldValue;
  onChange: (fieldName: string, value: FormFieldValue) => void;
  error?: IValidationError;
  context?: WebPartContext; // Required for People Picker
}
