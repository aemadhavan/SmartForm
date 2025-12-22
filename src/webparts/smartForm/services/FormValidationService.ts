import { FormSchema } from '../models/IFormSchema';
import { IFormData, IHyperlinkValue, IPeopleValue } from '../models/IFormData';
import { IFormValidationResult, IValidationError } from '../models/IValidationError';
import { FieldType } from '../models/FormFieldTypes';
import { ValidationRules } from '../utils/ValidationRules';
import { Constants } from '../utils/Constants';

export class FormValidationService {
  /**
   * Validates form data against the schema
   */
  public validateForm(schema: FormSchema, formData: IFormData): IFormValidationResult {
    const errors: IValidationError[] = [];

    schema.forEach((field) => {
      const value = formData[field.fieldName];
      const fieldErrors = this.validateField(field, value);
      errors.push(...fieldErrors);
    });

    return {
      isValid: errors.length === 0,
      errors,
    };
  }

  /**
   * Validates a single field value against its schema
   */
  private validateField(
    field: any,
    value: any
  ): IValidationError[] {
    const errors: IValidationError[] = [];

    // Check required fields
    if (field.required) {
      if (this.isEmpty(value)) {
        errors.push({
          fieldName: field.fieldName,
          message: Constants.VALIDATION_MESSAGES.REQUIRED_FIELD,
          severity: 'error',
        });
        return errors; // If required and empty, no need to check other validations
      }
    }

    // If value is empty and not required, it's valid
    if (this.isEmpty(value)) {
      return errors;
    }

    // Type-specific validation
    switch (field.fieldType) {
      case FieldType.Text:
      case FieldType.Textarea:
        this.validateTextField(field, value as string, errors);
        break;

      case FieldType.Number:
        this.validateNumberField(field, value, errors);
        break;

      case FieldType.Date:
        this.validateDateField(field, value, errors);
        break;

      case FieldType.Boolean:
        this.validateBooleanField(field, value, errors);
        break;

      case FieldType.Choice:
        this.validateChoiceField(field, value as string, errors);
        break;

      case FieldType.MultiChoice:
        this.validateMultiChoiceField(field, value as string[], errors);
        break;

      case FieldType.People:
        this.validatePeopleField(field, value, errors);
        break;

      case FieldType.Hyperlink:
        this.validateHyperlinkField(field, value as IHyperlinkValue, errors);
        break;
    }

    return errors;
  }

  /**
   * Validates text field
   */
  private validateTextField(field: any, value: string, errors: IValidationError[]): void {
    if (typeof value !== 'string') {
      errors.push({
        fieldName: field.fieldName,
        message: 'Value must be text',
        severity: 'error',
      });
      return;
    }

    if (field.minLength !== undefined && value.length < field.minLength) {
      errors.push({
        fieldName: field.fieldName,
        message: Constants.VALIDATION_MESSAGES.TEXT_TOO_SHORT.replace(
          '{min}',
          field.minLength.toString()
        ),
        severity: 'error',
      });
    }

    if (field.maxLength !== undefined && value.length > field.maxLength) {
      errors.push({
        fieldName: field.fieldName,
        message: Constants.VALIDATION_MESSAGES.TEXT_TOO_LONG.replace(
          '{max}',
          field.maxLength.toString()
        ),
        severity: 'error',
      });
    }
  }

  /**
   * Validates number field
   */
  private validateNumberField(field: any, value: any, errors: IValidationError[]): void {
    const numValue = typeof value === 'string' ? parseFloat(value) : value;

    if (typeof numValue !== 'number' || isNaN(numValue)) {
      errors.push({
        fieldName: field.fieldName,
        message: Constants.VALIDATION_MESSAGES.INVALID_NUMBER,
        severity: 'error',
      });
      return;
    }

    if (field.min !== undefined && numValue < field.min) {
      errors.push({
        fieldName: field.fieldName,
        message: Constants.VALIDATION_MESSAGES.NUMBER_TOO_LOW.replace(
          '{min}',
          field.min.toString()
        ),
        severity: 'error',
      });
    }

    if (field.max !== undefined && numValue > field.max) {
      errors.push({
        fieldName: field.fieldName,
        message: Constants.VALIDATION_MESSAGES.NUMBER_TOO_HIGH.replace(
          '{max}',
          field.max.toString()
        ),
        severity: 'error',
      });
    }
  }

  /**
   * Validates date field
   */
  private validateDateField(field: any, value: any, errors: IValidationError[]): void {
    if (!ValidationRules.isValidDate(value)) {
      errors.push({
        fieldName: field.fieldName,
        message: Constants.VALIDATION_MESSAGES.INVALID_DATE,
        severity: 'error',
      });
      return;
    }

    const dateValue = value instanceof Date ? value : new Date(value);

    if (field.minDate) {
      const minDate = new Date(field.minDate);
      if (dateValue < minDate) {
        errors.push({
          fieldName: field.fieldName,
          message: Constants.VALIDATION_MESSAGES.DATE_TOO_EARLY.replace(
            '{minDate}',
            minDate.toLocaleDateString()
          ),
          severity: 'error',
        });
      }
    }

    if (field.maxDate) {
      const maxDate = new Date(field.maxDate);
      if (dateValue > maxDate) {
        errors.push({
          fieldName: field.fieldName,
          message: Constants.VALIDATION_MESSAGES.DATE_TOO_LATE.replace(
            '{maxDate}',
            maxDate.toLocaleDateString()
          ),
          severity: 'error',
        });
      }
    }
  }

  /**
   * Validates boolean field
   */
  private validateBooleanField(field: any, value: any, errors: IValidationError[]): void {
    if (typeof value !== 'boolean') {
      errors.push({
        fieldName: field.fieldName,
        message: 'Value must be true or false',
        severity: 'error',
      });
    }
  }

  /**
   * Validates choice field
   */
  private validateChoiceField(field: any, value: string, errors: IValidationError[]): void {
    if (!field.choices || !Array.isArray(field.choices)) {
      return;
    }

    if (!field.choices.includes(value)) {
      errors.push({
        fieldName: field.fieldName,
        message: `Please select a valid option`,
        severity: 'error',
      });
    }
  }

  /**
   * Validates multi-choice field
   */
  private validateMultiChoiceField(field: any, value: string[], errors: IValidationError[]): void {
    if (!Array.isArray(value)) {
      errors.push({
        fieldName: field.fieldName,
        message: 'Value must be an array',
        severity: 'error',
      });
      return;
    }

    if (!field.choices || !Array.isArray(field.choices)) {
      return;
    }

    const invalidChoices = value.filter((v) => !field.choices.includes(v));
    if (invalidChoices.length > 0) {
      errors.push({
        fieldName: field.fieldName,
        message: `Invalid selection(s): ${invalidChoices.join(', ')}`,
        severity: 'error',
      });
    }
  }

  /**
   * Validates people field
   */
  private validatePeopleField(field: any, value: any, errors: IValidationError[]): void {
    // Value can be IPeopleValue or IPeopleValue[]
    const people: IPeopleValue[] = Array.isArray(value) ? value : [value];

    people.forEach((person: IPeopleValue) => {
      if (!person || typeof person !== 'object') {
        errors.push({
          fieldName: field.fieldName,
          message: 'Invalid people selection',
          severity: 'error',
        });
        return;
      }

      if (person.email && !ValidationRules.isValidEmail(person.email)) {
        errors.push({
          fieldName: field.fieldName,
          message: Constants.VALIDATION_MESSAGES.INVALID_EMAIL,
          severity: 'error',
        });
      }
    });
  }

  /**
   * Validates hyperlink field
   */
  private validateHyperlinkField(
    field: any,
    value: IHyperlinkValue,
    errors: IValidationError[]
  ): void {
    if (!value || typeof value !== 'object') {
      errors.push({
        fieldName: field.fieldName,
        message: 'Invalid hyperlink value',
        severity: 'error',
      });
      return;
    }

    if (!value.url || !ValidationRules.isValidUrl(value.url)) {
      errors.push({
        fieldName: field.fieldName,
        message: Constants.VALIDATION_MESSAGES.INVALID_URL,
        severity: 'error',
      });
    }
  }

  /**
   * Checks if a value is empty
   */
  private isEmpty(value: any): boolean {
    if (value === null || value === undefined) {
      return true;
    }
    if (typeof value === 'string' && value.trim() === '') {
      return true;
    }
    if (Array.isArray(value) && value.length === 0) {
      return true;
    }
    return false;
  }
}
