import * as React from 'react';
import { TextField } from '@fluentui/react/lib/TextField';
import { Label } from '@fluentui/react/lib/Label';
import { IFieldRendererProps } from './IFieldRendererProps';
import { IHyperlinkValue } from '../../models/IFormData';
import styles from './FieldRenderer.module.scss';

const HyperlinkField: React.FC<IFieldRendererProps> = ({ field, value, onChange, error }) => {
  const hyperlinkValue = (value as IHyperlinkValue) || { url: '', description: '' };

  const handleUrlChange = (_: any, newValue?: string): void => {
    onChange(field.fieldName, {
      ...hyperlinkValue,
      url: newValue || '',
    });
  };

  const handleDescriptionChange = (_: any, newValue?: string): void => {
    onChange(field.fieldName, {
      ...hyperlinkValue,
      description: newValue || '',
    });
  };

  return (
    <div className={styles.fieldWrapper}>
      <Label required={field.required}>{field.fieldLabel}</Label>
      {field.description && <div className={styles.fieldDescription}>{field.description}</div>}

      <TextField
        label="URL"
        value={hyperlinkValue.url || ''}
        onChange={handleUrlChange}
        placeholder="https://example.com"
        required={field.required}
        type="url"
        errorMessage={error?.message}
        aria-label={`${field.fieldLabel} - URL`}
        aria-required={field.required}
        aria-invalid={!!error}
      />

      <TextField
        label="Description (optional)"
        value={hyperlinkValue.description || ''}
        onChange={handleDescriptionChange}
        placeholder="Link description"
        aria-label={`${field.fieldLabel} - Description`}
      />
    </div>
  );
};

export default HyperlinkField;
