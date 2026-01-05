import { FieldType } from './FormFieldTypes';

export interface IFormFieldSchema {
  fieldLabel: string;
  fieldName: string;
  fieldType: FieldType;
  required?: boolean;

  // Tab / section support (optional)
  /**
   * Logical tab/section this field belongs to.
   * If omitted, field will appear in a default tab or in a flat layout
   * depending on the renderer implementation.
   */
  tabName?: string;

  /**
   * Optional order for tabs (lower comes first).
   * If multiple fields define different orders for the same tab,
   * the lowest is used.
   */
  tabOrder?: number;

  // Text/Textarea specific
  placeholder?: string;
  maxLength?: number;
  minLength?: number;

  // Number specific
  min?: number;
  max?: number;
  step?: number;

  // Choice/MultiChoice specific
  choices?: string[];

  // People specific
  allowMultiple?: boolean;

  // Date specific
  minDate?: string;
  maxDate?: string;

  // Attachment specific
  accept?: string;
  maxSize?: number;
  maxFiles?: number;

  // General
  description?: string;
  defaultValue?: any;
}

export type FormSchema = IFormFieldSchema[];

export interface ISchemaValidationResult {
  isValid: boolean;
  errors: string[];
  warnings?: string[];
}
