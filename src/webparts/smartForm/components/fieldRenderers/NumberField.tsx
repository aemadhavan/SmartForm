import * as React from 'react';
import { TextField } from '@fluentui/react/lib/TextField';
import { IFieldRendererProps } from './IFieldRendererProps';

const NumberField: React.FC<IFieldRendererProps> = ({ field, value, onChange, error }) => {
  const handleChange = (_: any, newValue?: string): void => {
    if (newValue === '') {
      onChange(field.fieldName, null);
      return;
    }

    const numValue = parseFloat(newValue || '0');
    if (!isNaN(numValue)) {
      onChange(field.fieldName, numValue);
    }
  };

  return (
    <TextField
      label={field.fieldLabel}
      value={value !== null && value !== undefined ? value.toString() : ''}
      onChange={handleChange}
      type="number"
      placeholder={field.placeholder}
      required={field.required}
      description={field.description}
      errorMessage={error?.message}
      min={field.min}
      max={field.max}
      step={field.step}
      aria-label={field.fieldLabel}
      aria-required={field.required}
      aria-invalid={!!error}
    />
  );
};

export default NumberField;
