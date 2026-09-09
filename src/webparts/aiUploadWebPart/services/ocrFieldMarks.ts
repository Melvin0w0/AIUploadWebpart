// Maps filled field values back to OCR words for the Extracted text Debug overlay.
import { isLeadingBlField } from '../constants/blSiteMap';
import { isNameField, isRegistrationNumberField } from '../constants/defaultFormFields';
import { isSubProjectNumberField } from '../constants/subProjectNumber';
import { isYesNoChoiceField } from '../constants/yesNo';
import { IOcrPageResult, IOcrWord } from './IPdfOcr';

export interface IOcrFieldMark {
  label: string;
  source: string;
  pageNumber: number;
  indexes: number[];
  color: string;
}

const FIELD_MARK_COLORS: string[] = ['#ff9f0a', '#34c759', '#af52de', '#ff2d55', '#5ac8fa', '#007aff', '#ffcc00'];

export function buildOcrFieldMarks(
  pages: IOcrPageResult[],
  fields: { label: string; value: string; debugSource?: string }[]
): IOcrFieldMark[] {
  const sorted = (fields || []).slice().sort((left, right) => (right.value || '').length - (left.value || '').length);
  const marks: IOcrFieldMark[] = [];
  const used: { [key: string]: boolean } = {};
  sorted.forEach((field) => {
    if (skipDebugField(field.label) || !(field.value || '').trim()) {
      return;
    }
    for (let pageIndex = 0; pageIndex < (pages || []).length; pageIndex++) {
      const page = pages[pageIndex];
      const indexes = findWordIndexesForValue(page.words || [], field.value);
      if (indexes.length === 0) {
        continue;
      }
      const key = page.pageNumber + ':' + indexes.join(',');
      if (used[key]) {
        continue;
      }
      used[key] = true;
      marks.push({
        label: field.label,
        source: field.debugSource || '',
        pageNumber: page.pageNumber,
        indexes,
        color: FIELD_MARK_COLORS[marks.length % FIELD_MARK_COLORS.length]
      });
      break;
    }
  });
  return marks;
}

export function formatOcrTextWithDebugMarks(
  words: IOcrWord[],
  marks: IOcrFieldMark[],
  withStyles: boolean
): string {
  const list = words || [];
  if (list.length === 0) {
    return '';
  }
  const order = visualWordOrder(list);
  const tagOf: string[] = [];
  for (let index = 0; index < list.length; index++) {
    tagOf[index] = '';
  }
  (marks || []).forEach((mark) => {
    const tag = fieldTagName(mark.label);
    (mark.indexes || []).forEach((index) => {
      if (index >= 0 && index < list.length && !tagOf[index]) {
        tagOf[index] = mark.source ? (tag + ':' + sanitizeTagPart(mark.source)) : tag;
      }
    });
  });

  let text = '';
  let openField = '';
  let openBold = false;
  let openUnderline = false;
  const closeStyles = (): void => {
    if (openBold) {
      text += '</b>';
      openBold = false;
    }
    if (openUnderline) {
      text += '</u>';
      openUnderline = false;
    }
  };
  const closeField = (): void => {
    if (!openField) {
      return;
    }
    closeStyles();
    text += '</' + fieldTagBase(openField) + '>';
    openField = '';
  };
  const openStyles = (word: IOcrWord): void => {
    if (withStyles && word.underline && !openUnderline) {
      text += '<u>';
      openUnderline = true;
    }
    if (withStyles && word.bold && !openBold) {
      text += '<b>';
      openBold = true;
    }
  };
  const openFieldTag = (tag: string): void => {
    if (!tag || tag === openField) {
      return;
    }
    closeField();
    text += '<' + tag + '>';
    openField = tag;
  };

  for (let cursor = 0; cursor < order.length; cursor++) {
    const index = order[cursor];
    const word = list[index];
    const previous = cursor > 0 ? list[order[cursor - 1]] : undefined;
    const tag = tagOf[index] || '';
    const lineHeight = previous ? Math.max(previous.y1 - previous.y0, 1) : 1;
    const isNewLine = previous ? (word.y0 - previous.y0 > lineHeight * 0.6) : false;
    if (isNewLine) {
      closeField();
      text += '\n';
    } else if (previous) {
      if (tag !== openField) {
        closeField();
      } else if (withStyles && (!!previous.bold !== !!word.bold || !!previous.underline !== !!word.underline)) {
        closeStyles();
      }
      text += joinGap(previous.text || '', word.text || '');
    }
    openFieldTag(tag);
    openStyles(word);
    text += word.text || '';
  }
  closeField();
  return text.trim();
}

function skipDebugField(label: string): boolean {
  return isNameField(label) ||
    isRegistrationNumberField(label) ||
    isLeadingBlField(label) ||
    isSubProjectNumberField(label) ||
    isYesNoChoiceField(label);
}

function findWordIndexesForValue(words: IOcrWord[], value: string): number[] {
  const target = normalizeDebugText(value);
  if (!target || target.length < 2) {
    return [];
  }
  const items: { index: number; norm: string }[] = [];
  for (let index = 0; index < (words || []).length; index++) {
    const norm = normalizeDebugText(words[index].text || '');
    if (norm) {
      items.push({ index, norm });
    }
  }
  let contains: number[] = [];
  for (let start = 0; start < items.length; start++) {
    let joined = '';
    const indexes: number[] = [];
    for (let end = start; end < items.length && end < start + 28; end++) {
      joined = joined ? (joined + ' ' + items[end].norm) : items[end].norm;
      indexes.push(items[end].index);
      if (joined === target) {
        return indexes;
      }
      if (joined.indexOf(target) >= 0 && joined.length <= target.length + 16) {
        contains = indexes.slice();
      }
      if (joined.length > target.length + 16 && joined.indexOf(target) < 0) {
        break;
      }
    }
  }
  return contains;
}

function visualWordOrder(words: IOcrWord[]): number[] {
  const order: number[] = [];
  for (let index = 0; index < words.length; index++) {
    order.push(index);
  }
  order.sort((left, right) => {
    const a = words[left];
    const b = words[right];
    const lineHeight = Math.max(a.y1 - a.y0, b.y1 - b.y0, 1);
    if (Math.abs(a.y0 - b.y0) > lineHeight * 0.5) {
      return a.y0 - b.y0;
    }
    return a.x0 - b.x0;
  });
  return order;
}

function normalizeDebugText(value: string): string {
  return (value || '')
    .replace(/<[^>]+>/g, ' ')
    .toLowerCase()
    .replace(/[^a-z0-9\u4e00-\u9fff]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function fieldTagName(label: string): string {
  const tag = sanitizeTagPart(label);
  return tag || 'Field';
}

function fieldTagBase(tag: string): string {
  const cut = tag.indexOf(':');
  return cut >= 0 ? tag.substring(0, cut) : tag;
}

function sanitizeTagPart(value: string): string {
  return (value || '').replace(/[<>]/g, '').replace(/\s+/g, '').trim();
}

function joinGap(left: string, right: string): string {
  const last = (left || '').charAt((left || '').length - 1);
  const first = (right || '').charAt(0);
  if (/[\u3400-\u9FFF]/.test(last) && /[\u3400-\u9FFF]/.test(first)) {
    return '';
  }
  return ' ';
}
