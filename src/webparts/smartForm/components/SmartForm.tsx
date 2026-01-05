import * as React from 'react';
import { useState, useMemo, useCallback, useEffect } from 'react';
import { DefaultButton } from '@fluentui/react/lib/Button';
import { Stack } from '@fluentui/react/lib/Stack';
import { Label } from '@fluentui/react/lib/Label';
import { Pivot, PivotItem } from '@fluentui/react/lib/Pivot';
import { ISmartFormProps } from './ISmartFormProps';
import DynamicFormRenderer from './formRenderer/DynamicFormRenderer';
import MessageDisplay from './common/MessageDisplay';
import SubmissionsList from './submissionsList/SubmissionsList';
import { FormSchema } from '../models/IFormSchema';
import { FieldType } from '../models/FormFieldTypes';
import { IFormData, FormFieldValue, IFormSubmission, IAttachmentValue } from '../models/IFormData';
import { IValidationError } from '../models/IValidationError';
import { SharePointService } from '../services/SharePointService';
import { FormValidationService } from '../services/FormValidationService';
import { SchemaValidationService } from '../services/SchemaValidationService';
import { FileNameGenerator } from '../utils/FileNameGenerator';
import { Constants } from '../utils/Constants';
import styles from './SmartForm.module.scss';

interface IMessage {
  type: 'success' | 'error' | 'info' | 'warning';
  text: string;
}

const formatFieldValueForMetadata = (value: any): string => {
  if (value === null || value === undefined) {
    return '';
  }

  if (typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean') {
    return String(value);
  }

  if (value instanceof Date) {
    return value.toISOString();
  }

  if (Array.isArray(value)) {
    return value
      .map((v) => formatFieldValueForMetadata(v))
      .filter((v) => !!v)
      .join(', ');
  }

  if (typeof value === 'object') {
    // People values
    if (typeof value.displayName === 'string') {
      return value.displayName;
    }

    // Hyperlink values
    if (typeof value.url === 'string') {
      return value.description ? `${value.description} (${value.url})` : value.url;
    }

    try {
      return JSON.stringify(value);
    } catch {
      return String(value);
    }
  }

  return String(value);
};

const SmartForm: React.FC<ISmartFormProps> = (props) => {
  const [schema, setSchema] = useState<FormSchema | null>(null);
  const [formData, setFormData] = useState<IFormData>({});
  const [validationErrors, setValidationErrors] = useState<IValidationError[]>([]);
  const [isSubmitting, setIsSubmitting] = useState<boolean>(false);
  const [message, setMessage] = useState<IMessage | null>(null);

  const [pivotKey, setPivotKey] = useState<string>('new');
  const [editFileName, setEditFileName] = useState<string | null>(null);
  const [editOriginalData, setEditOriginalData] = useState<IFormData | null>(null);
  const [editOriginalSubmissionMeta, setEditOriginalSubmissionMeta] = useState<
    Pick<IFormSubmission, 'submittedAt' | 'submittedBy'> | null
  >(null);
  const [submissionsRefreshKey, setSubmissionsRefreshKey] = useState<number>(0);

  // Initialize services
  const spService = useMemo(() => new SharePointService(props.context), [props.context]);
  const validationService = useMemo(() => new FormValidationService(), []);
  const schemaValidationService = useMemo(() => new SchemaValidationService(), []);

  const initializeFormFromSchema = useCallback((loadedSchema: FormSchema, showSuccessMessage?: boolean): void => {
    setSchema(loadedSchema);
    setEditFileName(null);
    setEditOriginalData(null);
    setEditOriginalSubmissionMeta(null);

    if (showSuccessMessage) {
      setMessage({
        type: 'success',
        text: Constants.SUCCESS_MESSAGES.SCHEMA_LOADED,
      });
    } else {
      setMessage(null);
    }

    const initialData: IFormData = {};
    loadedSchema.forEach((field) => {
      if (field.defaultValue !== undefined) {
        initialData[field.fieldName] = field.defaultValue;
      } else {
        initialData[field.fieldName] = null;
      }
    });

    setFormData(initialData);
    setValidationErrors([]);
  }, []);

  const loadSchemaFromWebPartProperties = useCallback((): void => {
    const schemaText = props.schemaJson ?? '';

    if (!schemaText.trim()) {
      setSchema(null);
      setFormData({});
      setValidationErrors([]);
      setMessage({
        type: 'info',
        text: 'No schema configured. Open the web part properties to paste or import a schema JSON.',
      });
      return;
    }

    const validationResult = schemaValidationService.validateSchema(schemaText);
    if (!validationResult.isValid) {
      setSchema(null);
      setFormData({});
      setValidationErrors([]);
      setMessage({
        type: 'error',
        text: validationResult.errors.join('\n'),
      });
      return;
    }

    const parsed = schemaValidationService.parseSchema(schemaText);
    if (!parsed) {
      setSchema(null);
      setFormData({});
      setValidationErrors([]);
      setMessage({
        type: 'error',
        text: 'Failed to parse schema JSON.',
      });
      return;
    }

    initializeFormFromSchema(parsed, false);
  }, [props.schemaJson, schemaValidationService, initializeFormFromSchema]);

  useEffect(() => {
    loadSchemaFromWebPartProperties();
  }, [loadSchemaFromWebPartProperties]);

  /**
   * Handle field value change
   */
  const handleFieldChange = useCallback((fieldName: string, value: FormFieldValue): void => {
    setFormData((prevData) => ({
      ...prevData,
      [fieldName]: value,
    }));

    // Clear validation error for this field
    setValidationErrors((prevErrors) =>
      prevErrors.filter((err) => err.fieldName !== fieldName)
    );
  }, []);

  /**
   * Handle form reset
   */
  const handleReset = useCallback((): void => {
    if (!schema) {
      return;
    }

    // If editing, reset back to the loaded submission values.
    if (editFileName && editOriginalData) {
      setFormData(editOriginalData);
      setValidationErrors([]);
      setMessage(null);
      return;
    }

    // Reset to default values
    const initialData: IFormData = {};
    schema.forEach((field) => {
      if (field.defaultValue !== undefined) {
        initialData[field.fieldName] = field.defaultValue;
      } else {
        initialData[field.fieldName] = null;
      }
    });

    setFormData(initialData);
    setValidationErrors([]);
    setMessage(null);
  }, [schema, editFileName, editOriginalData]);

  /**
   * Handle form submission
   */
  const handleSubmit = useCallback(async (): Promise<void> => {
    if (!schema) {
      return;
    }

    // Validate form
    const validationResult = validationService.validateForm(schema, formData);

    if (!validationResult.isValid) {
      setValidationErrors(validationResult.errors);
      setMessage({
        type: 'error',
        text: `Please fix ${validationResult.errors.length} validation error(s) before submitting.`,
      });
      return;
    }

    setIsSubmitting(true);
    setValidationErrors([]);

    try {
      // Check if library exists
      const libraryExists = await spService.ensureLibraryExists(props.targetLibraryName);
      if (!libraryExists) {
        throw new Error(
          Constants.ERROR_MESSAGES.LIBRARY_NOT_FOUND.replace(
            '{libraryName}',
            props.targetLibraryName
          )
        );
      }

      // Get current user
      const currentUser = await spService.getCurrentUser();

      // Build submission object
      const nowIso = new Date().toISOString();

      // Handle attachments if any
      const submissionData = { ...formData };
      for (const field of schema) {
        if (field.fieldType === FieldType.Attachment) {
          const attachments = submissionData[field.fieldName] as IAttachmentValue[];
          if (attachments && attachments.length > 0) {
            const folderPath = `Submissions/Attachments/${FileNameGenerator.generateFileName(props.fileNamePattern, currentUser.displayName).replace('.json', '')}`;
            const folderUrl = await spService.ensureFolderExists(props.targetLibraryName, folderPath);

            const uploadedAttachments: IAttachmentValue[] = [];
            for (const attachment of attachments) {
              if (attachment.fileContent) {
                const uploadedUrl = await spService.uploadFile(folderUrl, attachment.fileContent);
                uploadedAttachments.push({
                  fileName: attachment.fileName,
                  size: attachment.size,
                  serverRelativeUrl: uploadedUrl
                });
              } else {
                uploadedAttachments.push(attachment);
              }
            }
            submissionData[field.fieldName] = uploadedAttachments;
          }
        }
      }

      const submission: IFormSubmission = editFileName && editOriginalSubmissionMeta
        ? {
          submittedAt: editOriginalSubmissionMeta.submittedAt,
          submittedBy: editOriginalSubmissionMeta.submittedBy,
          editedAt: nowIso,
          editedBy: currentUser.displayName,
          formData: submissionData,
        }
        : {
          submittedAt: nowIso,
          submittedBy: currentUser.displayName,
          formData: submissionData,
        };

      // Add schema if configured
      if (props.persistSchema) {
        submission.schema = schema;
      }

      // Generate file name (or keep original when editing)
      const fileName = editFileName
        ? editFileName
        : FileNameGenerator.generateFileName(props.fileNamePattern, currentUser.displayName);

      // Convert to JSON
      const jsonContent = JSON.stringify(submission, null, 2);

      // Upload file
      const uploadedServerRelativeUrl = await spService.uploadJsonFile(
        props.targetLibraryName,
        fileName,
        jsonContent
      );

      // Persist selected Submitted Items columns into SharePoint metadata for fast listing.
      if (props.submittedItemsColumns && props.submittedItemsColumns.length > 0) {
        try {
          const columnRequests = props.submittedItemsColumns.map((fieldName) => {
            const match = schema.find((f) => f.fieldName === fieldName);
            return {
              internalName: SharePointService.toInternalName(fieldName),
              displayName: match?.fieldLabel || fieldName,
              maxLength: 255,
            };
          });

          await spService.ensureTextFieldsExist(props.targetLibraryName, columnRequests);

          const fieldsToUpdate: Record<string, any> = {};
          for (const fieldName of props.submittedItemsColumns) {
            const internalName = SharePointService.toInternalName(fieldName);
            fieldsToUpdate[internalName] = formatFieldValueForMetadata(formData[fieldName]);
          }

          await spService.updateFileListItemFields(uploadedServerRelativeUrl, fieldsToUpdate);
        } catch (metaError: any) {
          console.warn('Failed to persist metadata fields:', metaError);
          // Don't fail the submission; just inform the user.
          setMessage({
            type: 'warning',
            text:
              'Submission saved, but updating SharePoint columns failed. The Submitted Items grid may not show field columns until this is fixed. ' +
              (metaError?.message || ''),
          });
        }
      }

      // Show success message (only if metadata update didn't set a warning already)
      if (props.showSuccessMessage) {
        setMessage((prev) => {
          if (prev?.type === 'warning') {
            return prev;
          }

          return {
            type: 'success',
            text: editFileName
              ? `Submission updated successfully! File overwritten: ${fileName}`
              : `${Constants.SUCCESS_MESSAGES.FORM_SUBMITTED} File saved as: ${fileName}`,
          };
        });
      }

      // If we were editing, exit edit mode after save.
      if (editFileName) {
        setEditFileName(null);
        setEditOriginalData(null);
        setEditOriginalSubmissionMeta(null);
      }

      // Reset form
      handleReset();

      // Refresh submissions list
      setSubmissionsRefreshKey((k) => k + 1);
    } catch (error: any) {
      console.error('Form submission error:', error);
      setMessage({
        type: 'error',
        text: `${Constants.ERROR_MESSAGES.FORM_SUBMISSION_FAILED} ${error.message || ''
          }`,
      });
    } finally {
      setIsSubmitting(false);
    }
  }, [
    schema,
    formData,
    props,
    spService,
    validationService,
    handleReset,
    editFileName,
    editOriginalSubmissionMeta,
  ]);

  /**
   * Handle download JSON of current form data
   */
  const handleDownloadJson = useCallback((): void => {
    const jsonContent = JSON.stringify(formData, null, 2);
    const blob = new Blob([jsonContent], { type: 'application/json' });
    const url = URL.createObjectURL(blob);

    const link = document.createElement('a');
    link.href = url;
    link.download = `form-data-${new Date().getTime()}.json`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);

    setMessage({
      type: 'info',
      text: 'Form data downloaded successfully',
    });
  }, [formData]);

  /**
   * Open property pane so the author can edit the schema
   */
  const handleEditSchema = useCallback((): void => {
    props.context.propertyPane.open();
  }, [props.context]);

  const handlePivotLinkClick = useCallback((item?: PivotItem): void => {
    if (!item) {
      return;
    }

    setPivotKey(item.props.itemKey || 'new');
  }, []);

  const handleEditSubmission = useCallback(
    (fileName: string, submissionData: any): void => {
      // Use schema from the file if present; otherwise fallback to the currently loaded schema.
      const schemaFromFile = submissionData?.schema as FormSchema | undefined;
      const schemaToUse = schemaFromFile || schema;

      if (!schemaToUse) {
        setMessage({
          type: 'error',
          text:
            'Cannot edit this submission because no schema is available. Enable "Save Schema with Submission" or configure the schema in the web part properties.',
        });
        return;
      }

      setSchema(schemaToUse);

      const dataFromFile = (submissionData?.formData ?? submissionData) as IFormData;
      setFormData(dataFromFile);
      setEditOriginalData(dataFromFile);
      setEditFileName(fileName);

      const submittedAt = submissionData?.submittedAt;
      const submittedBy = submissionData?.submittedBy;
      if (submittedAt && submittedBy) {
        setEditOriginalSubmissionMeta({ submittedAt, submittedBy });
      } else {
        // Fallback: treat this as an update, but preserve metadata as best as possible.
        setEditOriginalSubmissionMeta({
          submittedAt: new Date().toISOString(),
          submittedBy: 'Unknown',
        });
      }

      setValidationErrors([]);
      setMessage({
        type: 'info',
        text: `Editing submission: ${fileName}. Click Submit to overwrite the file in SharePoint.`,
      });

      setPivotKey('new');
    },
    [schema]
  );

  const handleCancelEdit = useCallback((): void => {
    setEditFileName(null);
    setEditOriginalData(null);
    setEditOriginalSubmissionMeta(null);
    setValidationErrors([]);
    setMessage({ type: 'info', text: 'Edit cancelled.' });
    handleReset();
  }, [handleReset]);

  /**
   * Handle message dismiss
   */
  const handleDismissMessage = useCallback((): void => {
    setMessage(null);
  }, []);

  return (
    <section className={`${styles.smartForm} ${props.isDarkTheme ? styles.dark : ''}`}>
      <div className={styles.container}>
        {message && (
          <MessageDisplay
            message={message.text}
            messageType={message.type}
            onDismiss={handleDismissMessage}
          />
        )}

        <Pivot selectedKey={pivotKey} onLinkClick={handlePivotLinkClick}>
          <PivotItem headerText="New Submission" itemKey="new">
            {!schema ? (
              <div className={styles.formContainer}>
                <div className={styles.schemaInfo}>
                  <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 16 }}>
                    <div>
                      <Label className={styles.schemaLabel}>Schema not configured</Label>
                    </div>
                    <DefaultButton
                      text="Configure Schema"
                      onClick={handleEditSchema}
                      iconProps={{ iconName: 'Settings' }}
                    />
                  </Stack>
                </div>
              </div>
            ) : (
              <div className={styles.formContainer}>
                <div className={styles.schemaInfo}>
                  <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 16 }}>
                    <div>
                      <Label className={styles.schemaLabel}>
                        {editFileName
                          ? `Editing: ${editFileName}`
                          : `Schema Loaded - ${schema.length} field(s)`}
                      </Label>
                    </div>
                    <Stack horizontal tokens={{ childrenGap: 8 }}>
                      <DefaultButton
                        text="Edit Schema"
                        onClick={handleEditSchema}
                        iconProps={{ iconName: 'Settings' }}
                      />
                      {editFileName && (
                        <DefaultButton
                          text="Cancel Edit"
                          onClick={handleCancelEdit}
                          iconProps={{ iconName: 'Cancel' }}
                        />
                      )}
                    </Stack>
                  </Stack>
                </div>

                <DynamicFormRenderer
                  schema={schema}
                  formData={formData}
                  validationErrors={validationErrors}
                  onFieldChange={handleFieldChange}
                  onSubmit={handleSubmit}
                  onReset={handleReset}
                  onDownloadJson={handleDownloadJson}
                  isSubmitting={isSubmitting}
                  context={props.context}
                />
              </div>
            )}
          </PivotItem>

          <PivotItem headerText="Submitted Items" itemKey="submitted">
            <SubmissionsList
              context={props.context}
              targetLibraryName={props.targetLibraryName}
              onEditSubmission={handleEditSubmission}
              refreshKey={submissionsRefreshKey}
              schema={schema}
              selectedFieldColumns={props.submittedItemsColumns}
              pageSize={props.submittedItemsPageSize}
              isDarkTheme={props.isDarkTheme}
            />
          </PivotItem>
        </Pivot>
      </div>
    </section>
  );
};

export default SmartForm;
