import * as React from 'react';
import { Dropdown, IDropdownOption } from '@fluentui/react/lib/Dropdown';
import { IFieldRendererProps } from './IFieldRendererProps';

const MultiChoiceField: React.FC<IFieldRendererProps> = ({ field, value, onChange, error }) => {
  const options: IDropdownOption[] = (field.choices || []).map((choice) => ({
    key: choice,
    text: choice,
  }));

  const handleChange = (_: any, option?: IDropdownOption): void => {
    const currentValues = (value as string[]) || [];

    let newValues: string[];
    if (option) {
      if (option.selected) {
        newValues = [...currentValues, option.key as string];
      } else {
        newValues = currentValues.filter((v) => v !== option.key);
      }
    } else {
      newValues = [];
    }

    onChange(field.fieldName, newValues);
  };

  const selectedKeys = (value as string[]) || [];

  return (
    <Dropdown
      label={field.fieldLabel}
      selectedKeys={selectedKeys}
      onChange={handleChange}
      placeholder={field.placeholder || 'Select options'}
      multiSelect
      options={options}
      required={field.required}
      errorMessage={error?.message}
      aria-label={field.fieldLabel}
      aria-required={field.required}
      aria-invalid={!!error}
    />
  );
};

export default MultiChoiceField;
