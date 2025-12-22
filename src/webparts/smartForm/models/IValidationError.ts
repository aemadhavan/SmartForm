export interface IValidationError {
  fieldName: string;
  message: string;
  severity: 'error' | 'warning';
}

export interface IFormValidationResult {
  isValid: boolean;
  errors: IValidationError[];
}
