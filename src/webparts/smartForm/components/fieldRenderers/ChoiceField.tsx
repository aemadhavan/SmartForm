import * as React from 'react';
import { Dropdown, IDropdownOption } from '@fluentui/react/lib/Dropdown';
import { IFieldRendererProps } from './IFieldRendererProps';

const ChoiceField: React.FC<IFieldRendererProps> = ({ field, value, onChange, error }) => {
  const options: IDropdownOption[] = (field.choices || []).map((choice) => ({
    key: choice,
    text: choice,
  }));

  const handleChange = (_: any, option?: IDropdownOption): void => {
    onChange(field.fieldName, option?.key as string || null);
  };

  return (
    <Dropdown
      label={field.fieldLabel}
      selectedKey={value as string || undefined}
      onChange={handleChange}
      placeholder={field.placeholder || 'Select an option'}
      options={options}
      required={field.required}
      errorMessage={error?.message}
      aria-label={field.fieldLabel}
      aria-required={field.required}
      aria-invalid={!!error}
    />
  );
};

export default ChoiceField;
