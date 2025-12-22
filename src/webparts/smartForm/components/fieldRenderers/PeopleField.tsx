import * as React from 'react';
import { useState, useEffect, useCallback } from 'react';
import { Label } from '@fluentui/react/lib/Label';
import { TextField } from '@fluentui/react/lib/TextField';
import { Icon } from '@fluentui/react/lib/Icon';
import { IFieldRendererProps } from './IFieldRendererProps';
import { IPeopleValue } from '../../models/IFormData';
import { SharePointService } from '../../services/SharePointService';
import styles from './FieldRenderer.module.scss';

const PeopleField: React.FC<IFieldRendererProps> = ({ field, value, onChange, error, context }) => {
  const [searchText, setSearchText] = useState<string>('');
  const [suggestions, setSuggestions] = useState<any[]>([]);
  const [showSuggestions, setShowSuggestions] = useState<boolean>(false);
  const [selectedPeople, setSelectedPeople] = useState<IPeopleValue[]>([]);
  const [spService] = useState(() => context ? new SharePointService(context) : null);

  // Initialize selected people from value
  useEffect(() => {
    if (value) {
      const people = field.allowMultiple
        ? (Array.isArray(value) ? value as IPeopleValue[] : [])
        : [value as IPeopleValue];
      setSelectedPeople(people);
    } else {
      setSelectedPeople([]);
    }
  }, [value, field.allowMultiple]);

  // Search for users
  const handleSearch = useCallback(async (text: string): Promise<void> => {
    if (!spService || !text || text.trim().length < 2) {
      setSuggestions([]);
      return;
    }

    try {
      const users = await spService.searchUsers(text, 10);
      setSuggestions(users);
      setShowSuggestions(true);
    } catch (err) {
      console.error('Error searching users:', err);
      setSuggestions([]);
    }
  }, [spService]);

  const handleSearchTextChange = (event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string): void => {
    const text = newValue || '';
    setSearchText(text);
    if (text.trim().length >= 2) {
      void handleSearch(text);
    } else {
      setSuggestions([]);
      setShowSuggestions(false);
    }
  };

  const handleSelectUser = (user: any): void => {
    const newPerson: IPeopleValue = {
      id: user.Id,
      email: user.Email,
      displayName: user.Title,
      loginName: user.LoginName,
    };

    let newSelectedPeople: IPeopleValue[];

    if (field.allowMultiple) {
      // Check if already selected
      const alreadySelected = selectedPeople.some(p => p.email === newPerson.email);
      if (!alreadySelected) {
        newSelectedPeople = [...selectedPeople, newPerson];
      } else {
        newSelectedPeople = selectedPeople;
      }
    } else {
      newSelectedPeople = [newPerson];
    }

    setSelectedPeople(newSelectedPeople);
    onChange(
      field.fieldName,
      field.allowMultiple ? newSelectedPeople : newSelectedPeople[0]
    );
    setSearchText('');
    setSuggestions([]);
    setShowSuggestions(false);
  };

  const handleRemovePerson = (person: IPeopleValue): void => {
    const newSelectedPeople = selectedPeople.filter(p => p.email !== person.email);
    setSelectedPeople(newSelectedPeople);
    onChange(
      field.fieldName,
      field.allowMultiple ? newSelectedPeople : (newSelectedPeople.length > 0 ? newSelectedPeople[0] : null)
    );
  };

  if (!context || !spService) {
    return (
      <div className={styles.fieldWrapper}>
        <Label required={field.required}>{field.fieldLabel}</Label>
        <div className={styles.errorMessage}>
          Context is required for People Picker
        </div>
      </div>
    );
  }

  return (
    <div className={styles.fieldWrapper}>
      <Label required={field.required}>{field.fieldLabel}</Label>

      {/* Selected people chips */}
      {selectedPeople.length > 0 && (
        <div style={{ marginBottom: '8px', display: 'flex', flexWrap: 'wrap', gap: '4px' }}>
          {selectedPeople.map((person, index) => (
            <div
              key={index}
              style={{
                display: 'inline-flex',
                alignItems: 'center',
                padding: '4px 8px',
                backgroundColor: '#f3f2f1',
                borderRadius: '4px',
                fontSize: '12px',
              }}
            >
              <Icon iconName="Contact" style={{ marginRight: '4px' }} />
              <span>{person.displayName}</span>
              <Icon
                iconName="Cancel"
                style={{ marginLeft: '8px', cursor: 'pointer', fontSize: '10px' }}
                onClick={() => handleRemovePerson(person)}
              />
            </div>
          ))}
        </div>
      )}

      {/* Search input - only show if we can still add more */}
      {(field.allowMultiple || selectedPeople.length === 0) && (
        <div style={{ position: 'relative' }}>
          <TextField
            placeholder={field.placeholder || 'Start typing to search...'}
            value={searchText}
            onChange={handleSearchTextChange}
            onFocus={() => {
              if (suggestions.length > 0) {
                setShowSuggestions(true);
              }
            }}
            iconProps={{ iconName: 'Search' }}
          />

          {/* Suggestions dropdown */}
          {showSuggestions && suggestions.length > 0 && (
            <div
              style={{
                position: 'absolute',
                top: '100%',
                left: 0,
                right: 0,
                backgroundColor: 'white',
                border: '1px solid #ccc',
                maxHeight: '200px',
                overflowY: 'auto',
                zIndex: 1000,
                boxShadow: '0 2px 4px rgba(0,0,0,0.1)',
              }}
            >
              {suggestions.map((user, index) => (
                <div
                  key={index}
                  style={{
                    padding: '8px 12px',
                    cursor: 'pointer',
                    borderBottom: '1px solid #f3f2f1',
                  }}
                  onClick={() => handleSelectUser(user)}
                  onMouseEnter={(e) => {
                    e.currentTarget.style.backgroundColor = '#f3f2f1';
                  }}
                  onMouseLeave={(e) => {
                    e.currentTarget.style.backgroundColor = 'white';
                  }}
                >
                  <div style={{ fontWeight: 600 }}>{user.Title}</div>
                  <div style={{ fontSize: '12px', color: '#666' }}>{user.Email}</div>
                </div>
              ))}
            </div>
          )}
        </div>
      )}

      {field.description && <div className={styles.fieldDescription}>{field.description}</div>}
      {error && <div className={styles.errorMessage}>{error.message}</div>}
    </div>
  );
};

export default PeopleField;
