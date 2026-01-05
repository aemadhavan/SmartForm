import * as React from 'react';
import { useCallback, useRef } from 'react';
import { Stack } from '@fluentui/react/lib/Stack';
import { Label } from '@fluentui/react/lib/Label';
import { ActionButton, IconButton } from '@fluentui/react/lib/Button';
import { Text } from '@fluentui/react/lib/Text';
import { IFieldRendererProps } from './IFieldRendererProps';
import { IAttachmentValue } from '../../models/IFormData';
import styles from './FieldRenderer.module.scss';

const AttachmentField: React.FC<IFieldRendererProps> = ({ field, value, onChange, error }) => {
    const fileInputRef = useRef<HTMLInputElement>(null);
    const attachments = (value as IAttachmentValue[]) || [];

    const handleFileChange = useCallback((event: React.ChangeEvent<HTMLInputElement>): void => {
        const files = event.target.files;
        if (!files || files.length === 0) return;

        const newAttachments: IAttachmentValue[] = [...attachments];

        for (let i = 0; i < files.length; i++) {
            const file = files[i];

            // Basic validation
            if (field.maxSize && file.size > field.maxSize) {
                console.warn(`File ${file.name} exceeds max size of ${field.maxSize} bytes`);
                continue;
            }

            const isDuplicate = newAttachments.some(a => a.fileName === file.name);
            if (!isDuplicate) {
                newAttachments.push({
                    fileName: file.name,
                    size: file.size,
                    fileContent: file
                });
            }
        }

        onChange(field.fieldName, newAttachments);

        // Reset input so the same file can be selected again if removed
        if (fileInputRef.current) {
            fileInputRef.current.value = '';
        }
    }, [attachments, field, onChange]);

    const removeAttachment = useCallback((fileName: string): void => {
        const newAttachments = attachments.filter(a => a.fileName !== fileName);
        onChange(field.fieldName, newAttachments);
    }, [attachments, field.fieldName, onChange]);

    const onAddClick = (): void => {
        if (fileInputRef.current) {
            fileInputRef.current.click();
        }
    };

    const formatSize = (bytes: number): string => {
        if (bytes === 0) return '0 Bytes';
        const k = 1024;
        const sizes = ['Bytes', 'KB', 'MB', 'GB'];
        const i = Math.floor(Math.log(bytes) / Math.log(k));
        return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
    };

    const canAddMore = !field.maxFiles || attachments.length < field.maxFiles;

    return (
        <Stack tokens={{ childrenGap: 4 }}>
            <Label required={field.required}>{field.fieldLabel}</Label>
            {field.description && <Text variant="small" style={{ color: '#605e5c', marginBottom: 4 }}>{field.description}</Text>}

            <input
                type="file"
                ref={fileInputRef}
                onChange={handleFileChange}
                style={{ display: 'none' }}
                multiple={!field.maxFiles || field.maxFiles > 1}
                accept={field.accept}
            />

            <Stack tokens={{ childrenGap: 8 }}>
                {attachments.map((file) => (
                    <Stack
                        horizontal
                        verticalAlign="center"
                        key={file.fileName}
                        tokens={{ childrenGap: 8 }}
                        className={styles.attachmentItem}
                        style={{
                            padding: '4px 8px',
                            background: '#f3f2f1',
                            borderRadius: 2,
                            border: '1px solid #edebe9'
                        }}
                    >
                        <Stack.Item grow>
                            <Text variant="medium">
                                {file.fileName} <Text variant="small" style={{ color: '#605e5c' }}>({formatSize(file.size)})</Text>
                            </Text>
                        </Stack.Item>
                        <IconButton
                            iconProps={{ iconName: 'Cancel' }}
                            title="Remove"
                            ariaLabel="Remove"
                            onClick={() => removeAttachment(file.fileName)}
                        />
                    </Stack>
                ))}

                {canAddMore && (
                    <ActionButton
                        iconProps={{ iconName: 'Add' }}
                        onClick={onAddClick}
                        disabled={!canAddMore}
                    >
                        Add Attachment
                    </ActionButton>
                )}
            </Stack>

            {error && (
                <Text variant="small" style={{ color: '#a4262c', marginTop: 4 }}>
                    {error.message}
                </Text>
            )}
        </Stack>
    );
};

export default AttachmentField;
