import { FormSchema } from '../../models/IFormSchema';

export interface ISchemaInputProps {
  onSchemaLoaded: (schema: FormSchema) => void;
  onError: (message: string) => void;
}
