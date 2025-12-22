export enum FieldType {
  Text = 'Text',
  Textarea = 'Textarea',
  Number = 'Number',
  Date = 'Date',
  Boolean = 'Boolean',
  Choice = 'Choice',
  MultiChoice = 'MultiChoice',
  People = 'People',
  Hyperlink = 'Hyperlink'
}

export const SUPPORTED_FIELD_TYPES: string[] = [
  FieldType.Text,
  FieldType.Textarea,
  FieldType.Number,
  FieldType.Date,
  FieldType.Boolean,
  FieldType.Choice,
  FieldType.MultiChoice,
  FieldType.People,
  FieldType.Hyperlink
];
