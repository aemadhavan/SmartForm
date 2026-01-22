import * as React from 'react';
import { useCallback, useEffect, useMemo, useState } from 'react';
import { DetailsList, IColumn, SelectionMode } from '@fluentui/react/lib/DetailsList';
import { DefaultButton, IconButton } from '@fluentui/react/lib/Button';
import { Stack } from '@fluentui/react/lib/Stack';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { Text } from '@fluentui/react/lib/Text';
import { Dialog, DialogFooter, DialogType } from '@fluentui/react/lib/Dialog';
import MessageDisplay from '../common/MessageDisplay';
import { SharePointService } from '../../services/SharePointService';
import { FormSchema } from '../../models/IFormSchema';
import { ISubmissionsListProps } from './ISubmissionsListProps';
import styles from './SubmissionsList.module.scss';

interface IMessage {
  type: 'success' | 'error' | 'info' | 'warning';
  text: string;
}

const formatBytes = (bytes?: number): string => {
  if (!bytes && bytes !== 0) {
    return '';
  }

  if (bytes < 1024) {
    return `${bytes} B`;
  }

  const kb = bytes / 1024;
  if (kb < 1024) {
    return `${kb.toFixed(1)} KB`;
  }

  const mb = kb / 1024;
  return `${mb.toFixed(1)} MB`;
};

const formatDateTime = (iso?: string): string => {
  if (!iso) {
    return '';
  }

  const date = new Date(iso);
  if (isNaN(date.getTime())) {
    return iso;
  }

  return date.toLocaleString();
};

interface IListItemAllFields {
  [key: string]: any;
}

interface ISubmissionListItem {
  name: string;
  serverRelativeUrl: string;
  timeLastModified: string;
  length?: number;
  ListItemAllFields?: IListItemAllFields;
  _fieldValues?: Record<string, string>;
}

const formatFieldValue = (value: any): string => {
  if (value === null || value === undefined) {
    return '';
  }

  if (typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean') {
    return String(value);
  }

  if (Array.isArray(value)) {
    return value
      .map((v) => formatFieldValue(v))
      .filter((v) => !!v)
      .join(', ');
  }

  if (typeof value === 'object') {
    if (typeof value.displayName === 'string') {
      return value.displayName;
    }

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

const SubmissionsList: React.FC<ISubmissionsListProps> = ({
  context,
  targetLibraryName,
  onEditSubmission,
  refreshKey,
  schema,
  selectedFieldColumns,
  pageSize,
  isDarkTheme,
}) => {
  const spService = useMemo(() => new SharePointService(context), [context]);

  const [items, setItems] = useState<ISubmissionListItem[]>([]);
  const [isLoading, setIsLoading] = useState<boolean>(false);
  const [isLoadingMore, setIsLoadingMore] = useState<boolean>(false);
  const [nextLink, setNextLink] = useState<string | undefined>(undefined);
  const [message, setMessage] = useState<IMessage | null>(null);

  const [fileToDelete, setFileToDelete] = useState<{ name: string; serverRelativeUrl: string } | null>(null);
  const [isDeleting, setIsDeleting] = useState<boolean>(false);

  const effectivePageSize = pageSize || 50;
  const effectiveSelectedColumns = selectedFieldColumns || [];

  const internalNames = useMemo(
    () => effectiveSelectedColumns.map((f) => SharePointService.toInternalName(f)),
    [effectiveSelectedColumns]
  );

  const availableFields = useMemo(() => {
    const s: FormSchema = schema ?? [];
    return s.map((f) => ({
      fieldName: f.fieldName,
      label: f.fieldLabel || f.fieldName,
    }));
  }, [schema]);

  const fieldColumns = useMemo(() => {
    return effectiveSelectedColumns.map((fieldName) => {
      const internalName = SharePointService.toInternalName(fieldName);
      const label = availableFields.find((f) => f.fieldName === fieldName)?.label || fieldName;

      return { fieldName, internalName, label };
    });
  }, [availableFields, effectiveSelectedColumns]);

  const openPropertyPane = useCallback((): void => {
    context.propertyPane.open();
  }, [context]);

  const loadFirstPage = useCallback(async (): Promise<void> => {
    setIsLoading(true);
    setMessage(null);

    try {
      const libraryExists = await spService.ensureLibraryExists(targetLibraryName);
      if (!libraryExists) {
        setItems([]);
        setNextLink(undefined);
        setMessage({
          type: 'error',
          text: `Document library "${targetLibraryName}" not found. Check web part properties.`,
        });
        return;
      }

      const page = await spService.listLibraryFilesPage<ISubmissionListItem>(
        targetLibraryName,
        internalNames,
        effectivePageSize
      );

      const mapped = await Promise.all(
        page.items.map(async (f) => {
          const fieldValues: Record<string, string> = {};
          const missing: { internalName: string; fieldName: string }[] = [];

          for (const c of fieldColumns) {
            const v = formatFieldValue((f as any)?.ListItemAllFields?.[c.internalName]);
            fieldValues[c.internalName] = v;
            if (!v) {
              missing.push({ internalName: c.internalName, fieldName: c.fieldName });
            }
          }

          // Fallback: if metadata is missing/empty, read the JSON file and extract values from formData.
          if (missing.length > 0) {
            try {
              const fileText = await spService.getFileTextByServerRelativeUrl(f.serverRelativeUrl);
              const parsed = JSON.parse(fileText);
              const formData = (parsed?.formData ?? parsed) as any;

              const fieldsToBackfill: Record<string, any> = {};
              for (const m of missing) {
                const fromJson = formatFieldValue(formData?.[m.fieldName]);
                if (fromJson) {
                  fieldValues[m.internalName] = fromJson;
                  fieldsToBackfill[m.internalName] = fromJson;
                }
              }

              // Best-effort: backfill metadata so next load is fast.
              if (Object.keys(fieldsToBackfill).length > 0) {
                try {
                  await spService.updateFileListItemFields(f.serverRelativeUrl, fieldsToBackfill);
                } catch {
                  // ignore (permissions/field existence)
                }
              }
            } catch {
              // ignore
            }
          }

          return { ...f, _fieldValues: fieldValues };
        })
      );

      setItems(mapped);
      setNextLink(page.nextLink);
    } catch (e: any) {
      setItems([]);
      setNextLink(undefined);
      setMessage({
        type: 'error',
        text: `Failed to load submissions. ${e?.message || ''}`,
      });
    } finally {
      setIsLoading(false);
    }
  }, [effectivePageSize, fieldColumns, internalNames, spService, targetLibraryName]);

  const loadMore = useCallback(async (): Promise<void> => {
    if (!nextLink) {
      return;
    }

    setIsLoadingMore(true);
    setMessage(null);

    try {
      const page = await spService.listLibraryFilesPage<ISubmissionListItem>(
        targetLibraryName,
        internalNames,
        effectivePageSize,
        nextLink
      );

      const mapped = await Promise.all(
        page.items.map(async (f) => {
          const fieldValues: Record<string, string> = {};
          const missing: { internalName: string; fieldName: string }[] = [];

          for (const c of fieldColumns) {
            const v = formatFieldValue((f as any)?.ListItemAllFields?.[c.internalName]);
            fieldValues[c.internalName] = v;
            if (!v) {
              missing.push({ internalName: c.internalName, fieldName: c.fieldName });
            }
          }

          if (missing.length > 0) {
            try {
              const fileText = await spService.getFileTextByServerRelativeUrl(f.serverRelativeUrl);
              const parsed = JSON.parse(fileText);
              const formData = (parsed?.formData ?? parsed) as any;

              const fieldsToBackfill: Record<string, any> = {};
              for (const m of missing) {
                const fromJson = formatFieldValue(formData?.[m.fieldName]);
                if (fromJson) {
                  fieldValues[m.internalName] = fromJson;
                  fieldsToBackfill[m.internalName] = fromJson;
                }
              }

              if (Object.keys(fieldsToBackfill).length > 0) {
                try {
                  await spService.updateFileListItemFields(f.serverRelativeUrl, fieldsToBackfill);
                } catch {
                  // ignore
                }
              }
            } catch {
              // ignore
            }
          }

          return { ...f, _fieldValues: fieldValues };
        })
      );

      setItems((prev) => [...prev, ...mapped]);
      setNextLink(page.nextLink);
    } catch (e: any) {
      setMessage({
        type: 'error',
        text: `Failed to load more items. ${e?.message || ''}`,
      });
    } finally {
      setIsLoadingMore(false);
    }
  }, [effectivePageSize, fieldColumns, internalNames, nextLink, spService, targetLibraryName]);

  useEffect(() => {
    void loadFirstPage();
  }, [loadFirstPage, refreshKey]);

  const handleDismissMessage = useCallback((): void => {
    setMessage(null);
  }, []);

  const handleEditClick = useCallback(
    async (file: ISubmissionListItem): Promise<void> => {
      if (!onEditSubmission) {
        return;
      }

      setIsLoading(true);
      setMessage(null);

      try {
        const fileText = await spService.getFileTextByServerRelativeUrl(file.serverRelativeUrl);
        const parsed = JSON.parse(fileText);
        onEditSubmission(file.name, parsed);
      } catch (e: any) {
        setMessage({
          type: 'error',
          text: `Failed to open submission for editing. ${e?.message || ''}`,
        });
      } finally {
        setIsLoading(false);
      }
    },
    [onEditSubmission, spService]
  );

  const handleRequestDelete = useCallback(
    (file: { name: string; serverRelativeUrl: string }): void => {
      setMessage(null);
      setFileToDelete(file);
    },
    []
  );

  const handleCancelDelete = useCallback((): void => {
    setFileToDelete(null);
  }, []);

  const handleConfirmDelete = useCallback(async (): Promise<void> => {
    if (!fileToDelete) {
      return;
    }

    setIsDeleting(true);
    setMessage(null);

    try {
      // 1. Derive the attachment folder path from the file's server relative URL
      // File path example: /sites/MySite/MyLibrary/FormSubmission_123_User.json
      // Attachment folder path: /sites/MySite/MyLibrary/Submissions/Attachments/FormSubmission_123_User
      const fileUrl = fileToDelete.serverRelativeUrl;
      const lastSlashIndex = fileUrl.lastIndexOf('/');
      const libraryPath = fileUrl.substring(0, lastSlashIndex);
      const fileNameWithoutExtension = fileToDelete.name.replace(/\.[^/.]+$/, "");
      const attachmentFolderPath = `${libraryPath}/Submissions/Attachments/${fileNameWithoutExtension}`;

      // 2. Try to delete the attachment folder first (ignore errors if it doesn't exist)
      try {
        await spService.deleteFolderByServerRelativeUrl(attachmentFolderPath);
      } catch (e) {
        console.warn(`Could not delete attachment folder ${attachmentFolderPath}:`, e);
      }

      // 3. Delete the JSON submission file
      await spService.deleteFileByServerRelativeUrl(fileUrl);

      setMessage({ type: 'success', text: `Deleted submission: ${fileToDelete.name}` });
      setFileToDelete(null);
      await loadFirstPage();
    } catch (e: any) {
      setMessage({
        type: 'error',
        text: `Failed to delete file. ${e?.message || ''}`,
      });
    } finally {
      setIsDeleting(false);
    }
  }, [fileToDelete, loadFirstPage, spService]);

  const columns: IColumn[] = useMemo(() => {
    const base: IColumn[] = [
      {
        key: 'name',
        name: 'File',
        fieldName: 'name',
        minWidth: 220,
        isResizable: true,
        onRender: (item: ISubmissionListItem) => <Text>{item.name}</Text>,
      },
      {
        key: 'modified',
        name: 'Modified',
        fieldName: 'timeLastModified',
        minWidth: 160,
        isResizable: true,
        onRender: (item: ISubmissionListItem) => <Text>{formatDateTime(item.timeLastModified)}</Text>,
      },
      {
        key: 'size',
        name: 'Size',
        fieldName: 'length',
        minWidth: 90,
        isResizable: true,
        onRender: (item: ISubmissionListItem) => <Text>{formatBytes(item.length)}</Text>,
      },
    ];

    const dynamic: IColumn[] = fieldColumns.map((c) => ({
      key: `field:${c.internalName}`,
      name: c.label,
      fieldName: c.internalName,
      minWidth: 140,
      isResizable: true,
      onRender: (item: ISubmissionListItem) => (
        <Text>{item._fieldValues?.[c.internalName] ?? ''}</Text>
      ),
    }));

    const actions: IColumn = {
      key: 'actions',
      name: 'Actions',
      minWidth: 120,
      isResizable: false,
      onRender: (item: ISubmissionListItem) => (
        <Stack horizontal tokens={{ childrenGap: 4 }}>
          <IconButton
            iconProps={{ iconName: 'Edit' }}
            title="Edit"
            ariaLabel={`Edit ${item.name}`}
            disabled={!onEditSubmission || isLoading || isDeleting || isLoadingMore}
            onClick={() => void handleEditClick(item)}
          />
          <IconButton
            iconProps={{ iconName: 'Delete' }}
            title="Delete"
            ariaLabel={`Delete ${item.name}`}
            disabled={isLoading || isDeleting || isLoadingMore}
            onClick={() => handleRequestDelete({ name: item.name, serverRelativeUrl: item.serverRelativeUrl })}
          />
        </Stack>
      ),
    };

    return [...base, ...dynamic, actions];
  }, [fieldColumns, handleEditClick, handleRequestDelete, isDeleting, isLoading, isLoadingMore, onEditSubmission]);

  return (
    <div className={`${styles.submissionsList} ${isDarkTheme ? styles.dark : ''}`}>
      {message && (
        <MessageDisplay
          message={message.text}
          messageType={message.type}
          onDismiss={handleDismissMessage}
        />
      )}

      <Dialog
        hidden={!fileToDelete}
        onDismiss={handleCancelDelete}
        dialogContentProps={{
          type: DialogType.normal,
          title: 'Delete submission',
          subText: fileToDelete
            ? `Are you sure you want to delete "${fileToDelete.name}"? This cannot be undone.`
            : '',
        }}
        modalProps={{ isBlocking: true }}
      >
        <DialogFooter>
          <DefaultButton
            text="Cancel"
            onClick={handleCancelDelete}
            disabled={isDeleting}
          />
          <DefaultButton
            text={isDeleting ? 'Deleting...' : 'Delete'}
            onClick={() => void handleConfirmDelete()}
            disabled={isDeleting}
            iconProps={{ iconName: 'Delete' }}
          />
        </DialogFooter>
      </Dialog>

      <Stack horizontal verticalAlign="center" horizontalAlign="space-between" tokens={{ childrenGap: 12 }}>
        <Text variant="large">Submitted items</Text>
        <Stack horizontal tokens={{ childrenGap: 8 }}>
          <DefaultButton
            text="Configure"
            iconProps={{ iconName: 'Settings' }}
            disabled={isLoading || isDeleting || isLoadingMore}
            onClick={openPropertyPane}
          />
          <DefaultButton
            text="Refresh"
            iconProps={{ iconName: 'Refresh' }}
            disabled={isLoading || isDeleting || isLoadingMore}
            onClick={() => void loadFirstPage()}
          />
        </Stack>
      </Stack>

      <div className={styles.listContainer}>
        {isLoading ? (
          <Spinner size={SpinnerSize.medium} label="Loading submissions..." />
        ) : (
          <>
            <DetailsList
              items={items}
              columns={columns}
              selectionMode={SelectionMode.none}
              setKey="submissionsList"
              compact={false}
            />

            {nextLink && (
              <div className={styles.loadMore}>
                <DefaultButton
                  text={isLoadingMore ? 'Loading...' : 'Load more'}
                  disabled={isLoadingMore || isDeleting}
                  onClick={() => void loadMore()}
                />
              </div>
            )}
          </>
        )}

        {!isLoading && items.length === 0 && (
          <div className={styles.emptyState}>
            <Text>
              No submitted items found in &quot;{targetLibraryName}&quot;.
            </Text>
          </div>
        )}
      </div>
    </div>
  );
};

export default React.memo(SubmissionsList);
