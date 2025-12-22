import {
  PropertyPaneFieldType,
  type IPropertyPaneCustomFieldProps,
  type IPropertyPaneField
} from '@microsoft/sp-property-pane';

export interface IPropertyPaneJsonFilePickerProps {
  label?: string;
  buttonText?: string;
  accept?: string;
  onFileTextLoaded: (fileText: string, fileName?: string) => void;
}

interface IPropertyPaneJsonFilePickerInternalProps
  extends IPropertyPaneCustomFieldProps,
    IPropertyPaneJsonFilePickerProps {
  onRender: (elem: HTMLElement) => void;
  onDispose: (elem: HTMLElement) => void;
}

class PropertyPaneJsonFilePickerBuilder implements IPropertyPaneField<IPropertyPaneJsonFilePickerInternalProps> {
  public type: PropertyPaneFieldType = PropertyPaneFieldType.Custom;
  public targetProperty: string;
  public properties: IPropertyPaneJsonFilePickerInternalProps;

  private _inputId: string;

  public constructor(targetProperty: string, properties: IPropertyPaneJsonFilePickerProps) {
    this.targetProperty = targetProperty;
    this._inputId = `pp-json-file-${targetProperty}-${Date.now()}-${Math.random().toString(16).slice(2)}`;

    this.properties = {
      key: this._inputId,
      ...properties,
      accept: properties.accept ?? '.json,application/json',
      buttonText: properties.buttonText ?? 'Import JSON file',
      onRender: this._render.bind(this),
      onDispose: this._dispose.bind(this)
    };
  }

  private _render(elem: HTMLElement): void {
    const labelHtml = this.properties.label
      ? `<div style="font-weight: 600; margin: 0 0 8px 0;">${this._escapeHtml(this.properties.label)}</div>`
      : '';

    elem.innerHTML = `
      <div>
        ${labelHtml}
        <div style="display: flex; align-items: center; gap: 8px;">
          <input id="${this._inputId}" type="file" accept="${this._escapeHtml(this.properties.accept!)}" style="display:none" />
          <button type="button" class="ms-Button ms-Button--default">
            <span class="ms-Button-label">${this._escapeHtml(this.properties.buttonText!)}</span>
          </button>
          <span data-pp-json-file-name style="font-size: 12px; color: #605e5c;"></span>
        </div>
        <div data-pp-json-file-error style="margin-top: 6px; font-size: 12px; color: #a4262c;"></div>
      </div>
    `.trim();

    // this._inputId only contains safe characters (letters/numbers/dash), so we can use it directly.
    const input = elem.querySelector(`#${this._inputId}`) as HTMLInputElement | null;
    const button = elem.querySelector('button') as HTMLButtonElement | null;
    const fileNameEl = elem.querySelector('[data-pp-json-file-name]') as HTMLSpanElement | null;
    const errorEl = elem.querySelector('[data-pp-json-file-error]') as HTMLDivElement | null;

    if (!input || !button) {
      return;
    }

    const setError = (msg: string): void => {
      if (errorEl) {
        errorEl.textContent = msg;
      }
    };

    const clearError = (): void => setError('');

    button.onclick = () => {
      clearError();
      input.click();
    };

    input.onchange = () => {
      clearError();

      const file = input.files?.[0];
      if (!file) {
        return;
      }

      if (fileNameEl) {
        fileNameEl.textContent = file.name;
      }

      // Simple validation
      if (!file.name.toLowerCase().endsWith('.json')) {
        setError('Please select a .json file.');
        input.value = '';
        return;
      }

      const reader = new FileReader();
      reader.onerror = () => {
        setError('Failed to read file.');
      };

      reader.onload = () => {
        const text = (reader.result as string) ?? '';
        this.properties.onFileTextLoaded(text, file.name);
        // allow picking same file again
        input.value = '';
      };

      reader.readAsText(file);
    };
  }

  private _dispose(elem: HTMLElement): void {
    elem.innerHTML = '';
  }

  private _escapeHtml(s: string): string {
    return s
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#039;');
  }
}

export function PropertyPaneJsonFilePicker(
  targetProperty: string,
  properties: IPropertyPaneJsonFilePickerProps
): IPropertyPaneField<IPropertyPaneJsonFilePickerProps> {
  return new PropertyPaneJsonFilePickerBuilder(targetProperty, properties) as unknown as IPropertyPaneField<IPropertyPaneJsonFilePickerProps>;
}
