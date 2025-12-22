import * as React from 'react';
import { useState, useRef } from 'react';
import { PrimaryButton, DefaultButton } from '@fluentui/react/lib/Button';
import { TextField } from '@fluentui/react/lib/TextField';
import { ChoiceGroup, IChoiceGroupOption } from '@fluentui/react/lib/ChoiceGroup';
import { Stack } from '@fluentui/react/lib/Stack';
import { Label } from '@fluentui/react/lib/Label';
import { ISchemaInputProps } from './ISchemaInputProps';
import { SchemaValidationService } from '../../services/SchemaValidationService';
import { Constants } from '../../utils/Constants';
import styles from './SchemaInput.module.scss';

const SchemaInput: React.FC<ISchemaInputProps> = (props) => {
  const [uploadMethod, setUploadMethod] = useState<'file' | 'paste'>('file');
  const [schemaText, setSchemaText] = useState<string>('');
  const [fileName, setFileName] = useState<string>('');
  const fileInputRef = useRef<HTMLInputElement>(null);

  const validationService = new SchemaValidationService();

  const uploadMethodOptions: IChoiceGroupOption[] = [
    { key: 'file', text: 'Upload JSON File' },
    { key: 'paste', text: 'Paste JSON' },
  ];

  const handleFileChange = (event: React.ChangeEvent<HTMLInputElement>): void => {
    const file = event.target.files?.[0];
    if (!file) {
      return;
    }

    // Validate file extension
    if (!file.name.toLowerCase().endsWith(Constants.ALLOWED_FILE_EXTENSION)) {
      props.onError('Please select a JSON file (.json extension)');
      return;
    }

    // Validate file size
    if (file.size > Constants.MAX_FILE_SIZE) {
      props.onError(
        Constants.ERROR_MESSAGES.FILE_TOO_LARGE.replace(
          '{maxSize}',
          (Constants.MAX_FILE_SIZE / (1024 * 1024)).toString()
        )
      );
      return;
    }

    setFileName(file.name);

    // Read file content
    const reader = new FileReader();
    reader.onload = (e: ProgressEvent<FileReader>) => {
      const content = e.target?.result as string;
      setSchemaText(content);
    };
    reader.onerror = () => {
      props.onError('Failed to read file');
    };
    reader.readAsText(file);
  };

  const handleValidateAndLoad = (): void => {
    if (!schemaText || schemaText.trim() === '') {
      props.onError('Please upload a file or paste JSON content');
      return;
    }

    // Validate schema
    const validationResult = validationService.validateSchema(schemaText);

    if (!validationResult.isValid) {
      const errorMessage = validationResult.errors.join('\n');
      props.onError(errorMessage);
      return;
    }

    // Parse and load schema
    const schema = validationService.parseSchema(schemaText);
    if (!schema) {
      props.onError('Failed to parse schema');
      return;
    }

    // Show warnings if any
    if (validationResult.warnings && validationResult.warnings.length > 0) {
      console.warn('Schema warnings:', validationResult.warnings);
    }

    // Success - pass schema to parent
    props.onSchemaLoaded(schema);
  };

  const handleClear = (): void => {
    setSchemaText('');
    setFileName('');
    if (fileInputRef.current) {
      fileInputRef.current.value = '';
    }
  };

  const handleUploadClick = (): void => {
    fileInputRef.current?.click();
  };

  return (
    <div className={styles.schemaInput}>
      <Stack tokens={{ childrenGap: 20 }}>
        <div className={styles.header}>
          <Label className={styles.title}>Load Form Schema</Label>
          <p className={styles.description}>
            Upload a JSON file or paste JSON content to define your 
          </p>
        </div>

        <ChoiceGroup
          selectedKey={uploadMethod}
          options={uploadMethodOptions}
          onChange={(_, option) => setUploadMethod(option?.key as 'file' | 'paste')}
          label="Choose input method"
        />

        {uploadMethod === 'file' ? (
          <div className={styles.fileUpload}>
            <input
              ref={fileInputRef}
              type="file"
              accept=".json"
              onChange={handleFileChange}
              className={styles.hiddenFileInput}
              aria-label="Upload JSON schema file"
            />
            <PrimaryButton
              text={fileName || 'Choose File'}
              onClick={handleUploadClick}
              iconProps={{ iconName: 'Upload' }}
            />
            {fileName && (
              <div className={styles.fileInfo}>
                <span className={styles.fileName}>{fileName}</span>
              </div>
            )}
          </div>
        ) : (
          <TextField
            label="JSON Schema"
            multiline
            rows={15}
            value={schemaText}
            onChange={(_, newValue) => setSchemaText(newValue || '')}
            placeholder='Paste your JSON schema here, e.g.,
[
  {
    "fieldLabel": "Full Name",
    "fieldName": "fullName",
    "fieldType": "Text",
    "required": true
  }
]'
            resizable={false}
            aria-label="JSON schema text area"
          />
        )}

        <Stack horizontal tokens={{ childrenGap: 10 }}>
          <PrimaryButton
            text="Validate & Load Schema"
            onClick={handleValidateAndLoad}
            disabled={!schemaText}
            iconProps={{ iconName: 'CheckMark' }}
          />
          <DefaultButton
            text="Clear"
            onClick={handleClear}
            iconProps={{ iconName: 'Clear' }}
          />
        </Stack>

        <div className={styles.helpText}>
          <Label>Supported Field Types:</Label>
          <p>Text, Textarea, Number, Date, Boolean, Choice, MultiChoice, People, Hyperlink</p>
        </div>
      </Stack>
    </div>
  );
};

export default SchemaInput;
