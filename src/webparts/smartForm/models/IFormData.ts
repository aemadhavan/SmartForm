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

export interface IAttachmentValue {
  fileName: string;
  serverRelativeUrl?: string; // Present after upload
  size: number;
  fileContent?: File; // Present before upload
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
  | IAttachmentValue[]
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
