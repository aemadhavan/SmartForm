import * as React from 'react';
import { PrimaryButton, DefaultButton } from '@fluentui/react/lib/Button';
import { Stack } from '@fluentui/react/lib/Stack';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { Label } from '@fluentui/react/lib/Label';
import { Pivot, PivotItem } from '@fluentui/react/lib/Pivot';
import { IDynamicFormRendererProps } from './IDynamicFormRendererProps';
import { IValidationError } from '../../models/IValidationError';
import FieldRenderer from '../fieldRenderers/FieldRenderer';
import styles from './DynamicFormRenderer.module.scss';

const DynamicFormRenderer: React.FC<IDynamicFormRendererProps> = ({
  schema,
  formData,
  validationErrors,
  onFieldChange,
  onSubmit,
  onReset,
  onDownloadJson,
  isSubmitting,
  context,
}) => {
  const [activeTabKey, setActiveTabKey] = React.useState<string | undefined>(
    undefined
  );

  const getFieldError = (fieldName: string): IValidationError | undefined => {
    return validationErrors.find((err) => err.fieldName === fieldName);
  };

  const handleKeyDown = (e: React.KeyboardEvent): void => {
    // Prevent form submission on Enter key
    if (e.key === 'Enter' && e.target instanceof HTMLInputElement) {
      if ((e.target as HTMLInputElement).type !== 'textarea') {
        e.preventDefault();
      }
    }
  };

  // Determine if schema uses tabs and group fields accordingly
  const { hasTabs, tabs, fieldsByTab } = React.useMemo(() => {
    const anyTabs = schema.some((f) => !!f.tabName);

    if (!anyTabs) {
      return {
        hasTabs: false,
        tabs: [] as { key: string; name: string }[],
        fieldsByTab: {} as Record<string, typeof schema>,
      };
    }

    const map = new Map<
      string,
      { name: string; order: number; fields: typeof schema }
    >();

    for (const field of schema) {
      const tabName = field.tabName || 'General';
      const tabOrder = field.tabOrder ?? 0;

      if (!map.has(tabName)) {
        map.set(tabName, {
          name: tabName,
          order: tabOrder,
          fields: [],
        });
      }

      map.get(tabName)!.fields.push(field);

      // Keep the lowest order number for the tab
      if (tabOrder < map.get(tabName)!.order) {
        map.get(tabName)!.order = tabOrder;
      }
    }

    const tabArray = Array.from(map.values())
      .sort((a, b) => a.order - b.order || a.name.localeCompare(b.name))
      .map((t) => ({
        key: t.name,
        name: t.name,
      }));

    const fieldsMap: Record<string, typeof schema> = {};
    for (const [name, info] of Array.from(map.entries())) {
      fieldsMap[name] = info.fields;
    }

    return {
      hasTabs: true,
      tabs: tabArray,
      fieldsByTab: fieldsMap,
    };
  }, [schema]);

  // Ensure active tab is set when tabs are present
  React.useEffect(() => {
    if (hasTabs && tabs.length > 0 && !activeTabKey) {
      setActiveTabKey(tabs[0].key);
    }
  }, [hasTabs, tabs, activeTabKey]);

  const handleTabClick = (item?: PivotItem): void => {
    if (item?.props?.itemKey) {
      setActiveTabKey(item.props.itemKey);
    }
  };

  const renderField = (field: any, index: number) => (
    <FieldRenderer
      key={field.fieldName || index}
      field={field}
      value={formData[field.fieldName]}
      onChange={onFieldChange}
      error={getFieldError(field.fieldName)}
      context={context}
    />
  );

  return (
    <div className={styles.formRenderer}>
      {isSubmitting && (
        <div className={styles.spinnerOverlay}>
          <Spinner size={SpinnerSize.large} label="Submitting form..." />
        </div>
      )}

      <div className={styles.formHeader}>
        <Label className={styles.formTitle}>Powers Form</Label>
        <p className={styles.formDescription}>
          Fill out the form below. Fields marked with * are required.
        </p>
        <div className={styles.fieldCount}>
          Total fields: {schema.length}
        </div>
      </div>

      <div className={styles.formFields} onKeyDown={handleKeyDown}>
        {!hasTabs && (
          <>
            {schema.map((field, index) => renderField(field, index))}
          </>
        )}

        {hasTabs && (
          <Pivot selectedKey={activeTabKey} onLinkClick={handleTabClick}>
            {tabs.map((tab) => (
              <PivotItem
                key={tab.key}
                itemKey={tab.key}
                headerText={tab.name}
              >
                <div className={styles.tabContent}>
                  {fieldsByTab[tab.name]?.map((field, index) =>
                    renderField(field, index)
                  )}
                </div>
              </PivotItem>
            ))}
          </Pivot>
        )}
      </div>

      <div className={styles.formActions}>
        <Stack horizontal tokens={{ childrenGap: 10 }} wrap>
          <PrimaryButton
            text="Submit"
            onClick={onSubmit}
            disabled={isSubmitting}
            iconProps={{ iconName: 'Send' }}
            aria-label="Submit form"
          />
          <DefaultButton
            text="Reset Form"
            onClick={onReset}
            disabled={isSubmitting}
            iconProps={{ iconName: 'Refresh' }}
            aria-label="Reset form"
          />
          <DefaultButton
            text="Download JSON"
            onClick={onDownloadJson}
            disabled={isSubmitting}
            iconProps={{ iconName: 'Download' }}
            aria-label="Download form data as JSON"
          />
        </Stack>

        {validationErrors.length > 0 && (
          <div className={styles.validationSummary} role="alert" aria-live="polite">
            <p>Please fix the following errors:</p>
            <ul>
              {validationErrors.map((error, index) => (
                <li key={`${error.fieldName}-${index}`}>
                  <strong>{error.fieldName}:</strong> {error.message}
                </li>
              ))}
            </ul>
          </div>
        )}
      </div>
    </div>
  );
};

export default React.memo(DynamicFormRenderer);
