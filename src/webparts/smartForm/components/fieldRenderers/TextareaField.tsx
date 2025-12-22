import * as React from 'react';
import { TextField } from '@fluentui/react/lib/TextField';
import { IFieldRendererProps } from './IFieldRendererProps';

const TextareaField: React.FC<IFieldRendererProps> = ({ field, value, onChange, error }) => {
  const handleChange = (_: any, newValue?: string): void => {
    onChange(field.fieldName, newValue || '');
  };

  return (
    <TextField
      label={field.fieldLabel}
      value={(value as string) || ''}
      onChange={handleChange}
      multiline
      rows={4}
      placeholder={field.placeholder}
      maxLength={field.maxLength}
      required={field.required}
      description={field.description}
      errorMessage={error?.message}
      resizable={false}
      aria-label={field.fieldLabel}
      aria-required={field.required}
      aria-invalid={!!error}
    />
  );
};

export default TextareaField;
