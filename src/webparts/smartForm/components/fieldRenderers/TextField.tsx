import * as React from 'react';
import { TextField as FluentTextField } from '@fluentui/react/lib/TextField';
import { IFieldRendererProps } from './IFieldRendererProps';

const TextField: React.FC<IFieldRendererProps> = ({ field, value, onChange, error }) => {
  const handleChange = (_: any, newValue?: string): void => {
    onChange(field.fieldName, newValue || '');
  };

  return (
    <FluentTextField
      label={field.fieldLabel}
      value={(value as string) || ''}
      onChange={handleChange}
      placeholder={field.placeholder}
      maxLength={field.maxLength}
      required={field.required}
      description={field.description}
      errorMessage={error?.message}
      aria-label={field.fieldLabel}
      aria-required={field.required}
      aria-invalid={!!error}
      aria-describedby={error ? `${field.fieldName}-error` : undefined}
    />
  );
};

export default TextField;
