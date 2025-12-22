import { FormSchema, IFormFieldSchema, ISchemaValidationResult } from '../models/IFormSchema';
import { FieldType, SUPPORTED_FIELD_TYPES } from '../models/FormFieldTypes';
import { Constants } from '../utils/Constants';

export class SchemaValidationService {
  /**
   * Validates a JSON string to ensure it's a valid form schema
   */
  public validateSchema(jsonString: string): ISchemaValidationResult {
    const errors: string[] = [];
    const warnings: string[] = [];

    // Step 1: Validate JSON parse
    let parsedSchema: any;
    try {
      parsedSchema = JSON.parse(jsonString);
    } catch {
      errors.push(Constants.SCHEMA_VALIDATION_MESSAGES.INVALID_JSON);
      return { isValid: false, errors, warnings };
    }

    // Step 2: Validate that it's an array
    if (!Array.isArray(parsedSchema)) {
      errors.push(Constants.SCHEMA_VALIDATION_MESSAGES.NOT_ARRAY);
      return { isValid: false, errors, warnings };
    }

    // Step 3: Validate that array is not empty
    if (parsedSchema.length === 0) {
      errors.push(Constants.SCHEMA_VALIDATION_MESSAGES.EMPTY_SCHEMA);
      return { isValid: false, errors, warnings };
    }

    // Step 4: Validate each field
    const fieldNames = new Set<string>();

    parsedSchema.forEach((field: any, index: number) => {
      // Check required properties
      if (!field.fieldLabel) {
        errors.push(
          Constants.SCHEMA_VALIDATION_MESSAGES.MISSING_FIELD_LABEL.replace(
            '{index}',
            index.toString()
          )
        );
      }

      if (!field.fieldName) {
        errors.push(
          Constants.SCHEMA_VALIDATION_MESSAGES.MISSING_FIELD_NAME.replace(
            '{index}',
            index.toString()
          )
        );
      } else {
        // Check for duplicate field names
        if (fieldNames.has(field.fieldName)) {
          errors.push(
            Constants.SCHEMA_VALIDATION_MESSAGES.DUPLICATE_FIELD_NAME.replace(
              '{fieldName}',
              field.fieldName
            )
          );
        }
        fieldNames.add(field.fieldName);
      }

      if (!field.fieldType) {
        errors.push(
          Constants.SCHEMA_VALIDATION_MESSAGES.MISSING_FIELD_TYPE.replace(
            '{index}',
            index.toString()
          )
        );
      } else {
        // Validate field type is supported
        if (SUPPORTED_FIELD_TYPES.indexOf(field.fieldType) === -1) {
          errors.push(
            Constants.SCHEMA_VALIDATION_MESSAGES.INVALID_FIELD_TYPE
              .replace('{fieldName}', field.fieldName || `Field ${index}`)
              .replace('{fieldType}', field.fieldType)
          );
        }

        // Field-specific validation
        this.validateFieldSpecificProperties(field, errors, warnings);
      }
    });

    return {
      isValid: errors.length === 0,
      errors,
      warnings,
    };
  }

  /**
   * Validates field-specific properties based on field type
   */
  private validateFieldSpecificProperties(
    field: IFormFieldSchema,
    errors: string[],
    warnings: string[]
  ): void {
    const fieldType = field.fieldType;
    const fieldName = field.fieldName || 'Unknown field';

    // Choice and MultiChoice must have choices array
    if (fieldType === FieldType.Choice || fieldType === FieldType.MultiChoice) {
      if (!field.choices) {
        errors.push(
          Constants.SCHEMA_VALIDATION_MESSAGES.MISSING_CHOICES.replace(
            '{fieldName}',
            fieldName
          )
        );
      } else if (!Array.isArray(field.choices) || field.choices.length === 0) {
        errors.push(
          Constants.SCHEMA_VALIDATION_MESSAGES.EMPTY_CHOICES.replace(
            '{fieldName}',
            fieldName
          )
        );
      }
    }

    // Number field: validate min/max
    if (fieldType === FieldType.Number) {
      if (
        field.min !== undefined &&
        field.max !== undefined &&
        field.min > field.max
      ) {
        warnings.push(
          `Field "${fieldName}": min value (${field.min}) is greater than max value (${field.max})`
        );
      }
    }

    // Text/Textarea: validate length constraints
    if (fieldType === FieldType.Text || fieldType === FieldType.Textarea) {
      if (
        field.minLength !== undefined &&
        field.maxLength !== undefined &&
        field.minLength > field.maxLength
      ) {
        warnings.push(
          `Field "${fieldName}": minLength (${field.minLength}) is greater than maxLength (${field.maxLength})`
        );
      }
    }

    // Date: validate date range
    if (fieldType === FieldType.Date) {
      if (field.minDate && field.maxDate) {
        const minDate = new Date(field.minDate);
        const maxDate = new Date(field.maxDate);
        if (minDate > maxDate) {
          warnings.push(
            `Field "${fieldName}": minDate is after maxDate`
          );
        }
      }
    }
  }

  /**
   * Parses a JSON string and returns a typed FormSchema
   */
  public parseSchema(jsonString: string): FormSchema | null {
    const validation = this.validateSchema(jsonString);
    if (!validation.isValid) {
      return null;
    }

    try {
      return JSON.parse(jsonString) as FormSchema;
    } catch {
      return null;
    }
  }
}
