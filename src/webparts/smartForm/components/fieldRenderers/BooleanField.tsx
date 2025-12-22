import * as React from 'react';
import { Checkbox } from '@fluentui/react/lib/Checkbox';
import { IFieldRendererProps } from './IFieldRendererProps';
import styles from './FieldRenderer.module.scss';

const BooleanField: React.FC<IFieldRendererProps> = ({ field, value, onChange, error }) => {
  const handleChange = (_: any, checked?: boolean): void => {
    onChange(field.fieldName, checked || false);
  };

  return (
    <div className={styles.fieldWrapper}>
      <Checkbox
        label={field.fieldLabel}
        checked={value as boolean || false}
        onChange={handleChange}
        aria-label={field.fieldLabel}
        aria-required={field.required}
        aria-invalid={!!error}
      />
      {field.description && <div className={styles.fieldDescription}>{field.description}</div>}
      {error && <div className={styles.errorMessage}>{error.message}</div>}
    </div>
  );
};

export default BooleanField;
