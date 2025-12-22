import * as React from 'react';
import { DatePicker } from '@fluentui/react/lib/DatePicker';
import { IFieldRendererProps } from './IFieldRendererProps';
import styles from './FieldRenderer.module.scss';

const DateField: React.FC<IFieldRendererProps> = ({ field, value, onChange, error }) => {
  const handleSelectDate = (date: Date | null | undefined): void => {
    onChange(field.fieldName, date || null);
  };

  const dateValue = value instanceof Date ? value : value ? new Date(value as string) : undefined;

  const minDate = field.minDate ? new Date(field.minDate) : undefined;
  const maxDate = field.maxDate ? new Date(field.maxDate) : undefined;

  return (
    <div className={styles.fieldWrapper}>
      <DatePicker
        label={field.fieldLabel}
        value={dateValue}
        onSelectDate={handleSelectDate}
        placeholder={field.placeholder || 'Select a date'}
        isRequired={field.required}
        minDate={minDate}
        maxDate={maxDate}
        ariaLabel={field.fieldLabel}
      />
      {field.description && <div className={styles.fieldDescription}>{field.description}</div>}
      {error && <div className={styles.errorMessage}>{error.message}</div>}
    </div>
  );
};

export default DateField;
