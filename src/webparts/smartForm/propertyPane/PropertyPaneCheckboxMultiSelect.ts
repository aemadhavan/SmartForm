import {
  PropertyPaneFieldType,
  type IPropertyPaneCustomFieldProps,
  type IPropertyPaneField
} from '@microsoft/sp-property-pane';

export interface IPropertyPaneCheckboxMultiSelectOption {
  key: string;
  text: string;
}

export interface IPropertyPaneCheckboxMultiSelectProps {
  label?: string;
  description?: string;
  options: IPropertyPaneCheckboxMultiSelectOption[];
  selectedKeys: string[];
  onSelectionChanged: (selectedKeys: string[]) => void;
}

interface IPropertyPaneCheckboxMultiSelectInternalProps
  extends IPropertyPaneCustomFieldProps,
    IPropertyPaneCheckboxMultiSelectProps {
  onRender: (elem: HTMLElement) => void;
  onDispose: (elem: HTMLElement) => void;
}

class PropertyPaneCheckboxMultiSelectBuilder
  implements IPropertyPaneField<IPropertyPaneCheckboxMultiSelectInternalProps>
{
  public type: PropertyPaneFieldType = PropertyPaneFieldType.Custom;
  public targetProperty: string;
  public properties: IPropertyPaneCheckboxMultiSelectInternalProps;

  private _id: string;

  public constructor(targetProperty: string, properties: IPropertyPaneCheckboxMultiSelectProps) {
    this.targetProperty = targetProperty;
    this._id = `pp-ms-${targetProperty}-${Date.now()}-${Math.random().toString(16).slice(2)}`;

    this.properties = {
      key: this._id,
      ...properties,
      selectedKeys: properties.selectedKeys || [],
      options: properties.options || [],
      onRender: this._render.bind(this),
      onDispose: this._dispose.bind(this)
    };
  }

  private _render(elem: HTMLElement): void {
    const labelHtml = this.properties.label
      ? `<div style="font-weight: 600; margin: 0 0 6px 0;">${this._escapeHtml(this.properties.label)}</div>`
      : '';

    const descHtml = this.properties.description
      ? `<div style="font-size: 12px; color: #605e5c; margin: 0 0 8px 0;">${this._escapeHtml(this.properties.description)}</div>`
      : '';

    const optionsHtml = (this.properties.options || [])
      .map((o) => {
        const checked = (this.properties.selectedKeys || []).indexOf(o.key) >= 0 ? 'checked' : '';
        const optionId = `${this._id}-${this._safeId(o.key)}`;
        return `
          <label for="${optionId}" style="display:flex; gap:8px; align-items:center; margin: 6px 0;">
            <input id="${optionId}" type="checkbox" data-key="${this._escapeHtml(o.key)}" ${checked} />
            <span>${this._escapeHtml(o.text)}</span>
          </label>
        `.trim();
      })
      .join('');

    elem.innerHTML = `
      <div>
        ${labelHtml}
        ${descHtml}
        <div style="max-height: 220px; overflow: auto; border: 1px solid #edebe9; padding: 8px; border-radius: 2px;">
          ${optionsHtml || '<div style="font-size:12px; color:#605e5c;">No fields available (load a schema first).</div>'}
        </div>
      </div>
    `.trim();

    const checkboxes = Array.from(elem.querySelectorAll('input[type="checkbox"][data-key]')) as HTMLInputElement[];
    const onChanged = (): void => {
      const selected: string[] = [];
      for (const cb of checkboxes) {
        const key = cb.getAttribute('data-key') || '';
        if (cb.checked && key) {
          selected.push(key);
        }
      }

      // update internal state so re-render reflects current selection
      this.properties.selectedKeys = selected;
      this.properties.onSelectionChanged(selected);
    };

    for (const cb of checkboxes) {
      cb.onchange = onChanged;
    }
  }

  private _dispose(elem: HTMLElement): void {
    elem.innerHTML = '';
  }

  private _safeId(s: string): string {
    return (s || '').replace(/[^A-Za-z0-9_-]/g, '_');
  }

  private _escapeHtml(s: string): string {
    return (s || '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#039;');
  }
}

export function PropertyPaneCheckboxMultiSelect(
  targetProperty: string,
  properties: IPropertyPaneCheckboxMultiSelectProps
): IPropertyPaneField<IPropertyPaneCheckboxMultiSelectProps> {
  return new PropertyPaneCheckboxMultiSelectBuilder(targetProperty, properties) as unknown as IPropertyPaneField<IPropertyPaneCheckboxMultiSelectProps>;
}
