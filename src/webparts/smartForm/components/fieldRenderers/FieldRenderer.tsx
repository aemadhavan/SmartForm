import * as React from 'react';
import { FieldType } from '../../models/FormFieldTypes';
import { IFieldRendererProps } from './IFieldRendererProps';
import TextField from './TextField';
import TextareaField from './TextareaField';
import NumberField from './NumberField';
import DateField from './DateField';
import BooleanField from './BooleanField';
import ChoiceField from './ChoiceField';
import MultiChoiceField from './MultiChoiceField';
import PeopleField from './PeopleField';
import HyperlinkField from './HyperlinkField';
import styles from './FieldRenderer.module.scss';

const FieldRenderer: React.FC<IFieldRendererProps> = (props) => {
  const { field } = props;

  const renderField = (): JSX.Element => {
    switch (field.fieldType) {
      case FieldType.Text:
        return <TextField {...props} />;

      case FieldType.Textarea:
        return <TextareaField {...props} />;

      case FieldType.Number:
        return <NumberField {...props} />;

      case FieldType.Date:
        return <DateField {...props} />;

      case FieldType.Boolean:
        return <BooleanField {...props} />;

      case FieldType.Choice:
        return <ChoiceField {...props} />;

      case FieldType.MultiChoice:
        return <MultiChoiceField {...props} />;

      case FieldType.People:
        return <PeopleField {...props} />;

      case FieldType.Hyperlink:
        return <HyperlinkField {...props} />;

      default:
        return (
          <div className={styles.unsupportedField}>
            <p>Unsupported field type: {field.fieldType}</p>
          </div>
        );
    }
  };

  return (
    <div className={styles.fieldContainer}>
      {renderField()}
    </div>
  );
};

export default React.memo(FieldRenderer);
