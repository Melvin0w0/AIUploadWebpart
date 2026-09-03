import { isIssueDateField } from '../constants/issueDate';
import { isNameField, isRegistrationNumberField } from '../constants/defaultFormFields';
import { isLeadingBlField } from '../constants/blSiteMap';
import { isSubProjectNumberField, SUB_PROJECT_NONE } from '../constants/subProjectNumber';
import { isYesNoChoiceField } from '../constants/yesNo';

const STORAGE_KEY: string = 'aiUpload.fieldHistory.v1';
const MAX_VALUES: number = 12;
const MAX_RECORDS: number = 8;

export interface IHistoryField {
  label: string;
  value: string;
}

export interface IHistoryRecord {
  id: string;
  savedAt: number;
  summary: string;
  fields: IHistoryField[];
}

export interface IFieldHistory {
  values: { [label: string]: string[] };
  records: IHistoryRecord[];
}

export function emptyFieldHistory(): IFieldHistory {
  return { values: {}, records: [] };
}

export function loadFieldHistory(): IFieldHistory {
  try {
    if (typeof window === 'undefined' || !window.localStorage) {
      return emptyFieldHistory();
    }
    const raw = window.localStorage.getItem(STORAGE_KEY);
    if (!raw) {
      return emptyFieldHistory();
    }
    const parsed = JSON.parse(raw) as IFieldHistory;
    return {
      values: parsed && parsed.values ? parsed.values : {},
      records: parsed && parsed.records ? parsed.records : []
    };
  } catch {
    return emptyFieldHistory();
  }
}

export function isHistoryTextField(label: string): boolean {
  return !isYesNoChoiceField(label)
    && !isLeadingBlField(label)
    && !isSubProjectNumberField(label)
    && !isRegistrationNumberField(label)
    && !isNameField(label)
    && !isIssueDateField(label);
}

export function suggestionsFor(history: IFieldHistory, label: string, query: string): string[] {
  const items = (history.values[label] || []).slice();
  const needle = (query || '').trim().toLowerCase();
  if (!needle) {
    return items;
  }
  return items.filter((item) => item.toLowerCase().indexOf(needle) >= 0 && item.toLowerCase() !== needle);
}

export function rememberFieldValue(history: IFieldHistory, label: string, value: string): IFieldHistory {
  const trimmed = (value || '').replace(/\s+/g, ' ').trim();
  if (!isHistoryTextField(label) || !trimmed || trimmed === SUB_PROJECT_NONE) {
    return history;
  }
  const current = history.values[label] || [];
  const next = [trimmed].concat(current.filter((item) => item.toLowerCase() !== trimmed.toLowerCase()));
  return {
    values: {
      ...history.values,
      [label]: next.slice(0, MAX_VALUES)
    },
    records: history.records
  };
}

export function rememberRecord(history: IFieldHistory, fields: IHistoryField[]): IFieldHistory {
  const savedFields = fields
    .map((field) => ({ label: field.label, value: (field.value || '').replace(/\s+/g, ' ').trim() }))
    .filter((field) => field.value.length > 0);
  if (savedFields.length === 0) {
    return history;
  }

  let next = history;
  savedFields.forEach((field) => {
    next = rememberFieldValue(next, field.label, field.value);
  });

  const record: IHistoryRecord = {
    id: `record-${Date.now()}`,
    savedAt: Date.now(),
    summary: recordSummary(savedFields),
    fields: savedFields
  };
  const records = [record].concat(next.records).slice(0, MAX_RECORDS);
  return {
    values: next.values,
    records
  };
}

export function saveFieldHistory(history: IFieldHistory): void {
  try {
    if (typeof window === 'undefined' || !window.localStorage) {
      return;
    }
    window.localStorage.setItem(STORAGE_KEY, JSON.stringify(history));
  } catch {
    // Ignore quota / private mode.
  }
}

function recordSummary(fields: IHistoryField[]): string {
  const get = (label: string): string => {
    const match = fields.filter((field) => field.label.toLowerCase() === label.toLowerCase())[0];
    return match ? match.value : '';
  };
  const parts = [get('Name') || get('Ref No'), get('Organization'), get('Issue Date')]
    .filter((part) => part.length > 0);
  if (parts.length === 0) {
    return savedFieldsFallback(fields);
  }
  return parts.join(' · ');
}

function savedFieldsFallback(fields: IHistoryField[]): string {
  const first = fields.filter((field) => field.value)[0];
  return first ? first.value : 'Saved record';
}
