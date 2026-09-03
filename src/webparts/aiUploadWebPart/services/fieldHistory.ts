import { isIssueDateField } from '../constants/issueDate';
import { isNameField, isRegistrationNumberField } from '../constants/defaultFormFields';
import { isLeadingBlField } from '../constants/blSiteMap';
import { isSubProjectNumberField, SUB_PROJECT_NONE } from '../constants/subProjectNumber';
import { isYesNoChoiceField } from '../constants/yesNo';

const STORAGE_KEY: string = 'aiUpload.fieldHistory.v1';
const MAX_VALUES: number = 12;

export interface IFieldHistory {
  values: { [label: string]: string[] };
}

export function emptyFieldHistory(): IFieldHistory {
  return { values: {} };
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
    const parsed = JSON.parse(raw) as { values?: { [label: string]: string[] } };
    return {
      values: parsed && parsed.values ? parsed.values : {}
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
    }
  };
}

export function rememberFieldValues(
  history: IFieldHistory,
  fields: { label: string; value: string }[]
): IFieldHistory {
  let next = history;
  fields.forEach((field) => {
    next = rememberFieldValue(next, field.label, field.value);
  });
  return next;
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
