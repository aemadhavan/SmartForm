export class Constants {
  // File upload limits
  public static readonly MAX_FILE_SIZE = 5 * 1024 * 1024; // 5MB
  public static readonly ALLOWED_FILE_EXTENSION = '.json';

  // Default configuration values
  public static readonly DEFAULT_LIBRARY_NAME = 'Form Submissions';
  public static readonly DEFAULT_FILE_NAME_PATTERN = 'FormSubmission_{timestamp}_{userName}.json';

  // Validation messages
  public static readonly VALIDATION_MESSAGES = {
    REQUIRED_FIELD: 'This field is required',
    INVALID_EMAIL: 'Please enter a valid email address',
    INVALID_URL: 'Please enter a valid URL (http:// or https://)',
    INVALID_NUMBER: 'Please enter a valid number',
    INVALID_DATE: 'Please enter a valid date',
    NUMBER_TOO_LOW: 'Value must be greater than or equal to {min}',
    NUMBER_TOO_HIGH: 'Value must be less than or equal to {max}',
    TEXT_TOO_SHORT: 'Text must be at least {min} characters',
    TEXT_TOO_LONG: 'Text must not exceed {max} characters',
    DATE_TOO_EARLY: 'Date must be on or after {minDate}',
    DATE_TOO_LATE: 'Date must be on or before {maxDate}',
  };

  // Schema validation messages
  public static readonly SCHEMA_VALIDATION_MESSAGES = {
    INVALID_JSON: 'The uploaded file is not valid JSON',
    NOT_ARRAY: 'Schema must be an array of field definitions',
    EMPTY_SCHEMA: 'Schema cannot be empty',
    MISSING_FIELD_LABEL: 'Field at index {index} is missing "fieldLabel"',
    MISSING_FIELD_NAME: 'Field at index {index} is missing "fieldName"',
    MISSING_FIELD_TYPE: 'Field at index {index} is missing "fieldType"',
    INVALID_FIELD_TYPE: 'Field "{fieldName}" has unsupported field type "{fieldType}"',
    DUPLICATE_FIELD_NAME: 'Duplicate field name "{fieldName}" found',
    MISSING_CHOICES: 'Field "{fieldName}" is of type Choice/MultiChoice but missing "choices" array',
    EMPTY_CHOICES: 'Field "{fieldName}" has an empty "choices" array',
  };

  // Success messages
  public static readonly SUCCESS_MESSAGES = {
    FORM_SUBMITTED: 'Form submitted successfully!',
    SCHEMA_LOADED: 'Schema loaded successfully',
    FILE_UPLOADED: 'File uploaded to {libraryName}',
  };

  // Error messages
  public static readonly ERROR_MESSAGES = {
    SCHEMA_LOAD_FAILED: 'Failed to load schema',
    FORM_SUBMISSION_FAILED: 'Failed to submit form. Please try again.',
    LIBRARY_NOT_FOUND: 'Document library "{libraryName}" not found',
    NO_WRITE_PERMISSION: 'You do not have permission to upload files to this library',
    FILE_TOO_LARGE: 'File size exceeds {maxSize}MB limit',
  };

  // Performance thresholds
  public static readonly LARGE_FORM_THRESHOLD = 50; // Number of fields considered "large"
  public static readonly DEBOUNCE_DELAY_MS = 300; // Debounce delay for text inputs
}
