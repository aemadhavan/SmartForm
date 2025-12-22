export interface IHyperlinkValue {
  url: string;
  description: string;
}

export interface IPeopleValue {
  id: number;
  email: string;
  displayName: string;
  loginName: string;
}

export type FormFieldValue =
  | string
  | number
  | boolean
  | Date
  | string[]
  | IPeopleValue
  | IPeopleValue[]
  | IHyperlinkValue
  | null
  | undefined;

export interface IFormData {
  [fieldName: string]: FormFieldValue;
}

export interface IFormSubmission {
  submittedAt: string;
  submittedBy: string;
  editedAt?: string;
  editedBy?: string;
  formData: IFormData;
  schemaVersion?: string;
  schema?: any;
}
