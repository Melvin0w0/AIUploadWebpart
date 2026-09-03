import { isOrganizationField, isReceiverField, isRefNoField, isSenderField, isSubjectField } from '../constants/defaultFormFields';
import { isProjectNumberField } from '../constants/projectNumber';
import { IOcrPageResult, IOcrWord } from './IPdfOcr';
import { joinOcrWords } from './ocrSelection';

const FIELD_ALIASES: { [label: string]: string[] } = {
  'Name': ['name'],
  'Registration Number': ['registration number', 'registration no', 'reg no'],
  'Leading BL': ['leading bl', 'leading b/l', 'leading b l', 'business line', 'bl'],
  'Project Number': ['project number', 'project no'],
  'Sub-Project Number': ['sub-project number', 'sub project number', 'sub-project no', 'sub project no'],
  'Organization': ['organization', 'organisation'],
  'Sender': ['sender'],
  'Receiver': ['receiver', 'recipient'],
  'Subject': ['subject'],
  'File No': ['file no', 'file number', 'file no.'],
  'Ref No': ['ref no', 'reference no', 'reference number', 'ref no.', 'ref.'],
  'Issue Date': ['issue date', 'issued date', 'date of issue'],
  'Attachment': ['attachment', 'attachments', 'encl', 'enclosure'],
  'Scan': ['scan'],
  'Remark': ['remark', 'remarks'],
  'Location': ['location'],
  'cc to AECOM': ['cc to aecom', 'cc aecom', 'copy to aecom']
};

interface ILabelHit {
  field: string;
  aliasLength: number;
  start: number;
  end: number;
}

export function extractFieldValues(pages: IOcrPageResult[], fieldLabels: string[]): { [label: string]: string } {
  const values: { [label: string]: string } = {};
  const words: IOcrWord[] = [];
  pages.forEach((page) => {
    (page.words || []).forEach((word) => words.push(word));
  });

  const fromLayout = extractFromWords(words, fieldLabels);
  const fromText = extractFromText(
    pages.map((page) => page.text || '').join('\n'),
    fieldLabels
  );

  fieldLabels.forEach((label) => {
    if (isSenderField(label) || isReceiverField(label) || isSubjectField(label) || isRefNoField(label) || isProjectNumberField(label) || isOrganizationField(label)) {
      return;
    }
    const layoutValue = (fromLayout[label] || '').trim();
    const textValue = (fromText[label] || '').trim();
    const chosen = layoutValue || textValue;
    if (chosen) {
      values[label] = chosen;
    }
  });

  return values;
}

export function extractOurRefOnly(pages: IOcrPageResult[]): string {
  return extractLabeledRef(pages, ourRefWordIndex, ourRefFromText);
}

export function extractOurRefNo(pages: IOcrPageResult[]): string {
  return extractOurRefOnly(pages)
    || extractLabeledRef(pages, plainRefWordIndex, plainRefFromText);
}

export function extractYourRefNo(pages: IOcrPageResult[]): string {
  return extractLabeledRef(pages, yourRefWordIndex, yourRefFromText);
}

function extractLabeledRef(
  pages: IOcrPageResult[],
  wordIndex: (words: IOcrWord[], index: number) => number,
  fromText: (text: string) => string
): string {
  const list = pages || [];
  for (let index = 0; index < list.length; index++) {
    const page = list[index];
    if (!page) {
      continue;
    }
    try {
      const words = (page.words || []).slice().sort((left, right) => {
        if (Math.abs(left.y0 - right.y0) > 8) {
          return left.y0 - right.y0;
        }
        return left.x0 - right.x0;
      });
      for (let wordIndexCursor = 0; wordIndexCursor < words.length; wordIndexCursor++) {
        const refIndex = wordIndex(words, wordIndexCursor);
        if (refIndex < 0) {
          continue;
        }
        const value = valueAfterRef(words, refIndex);
        if (value) {
          return value;
        }
      }
      const fromPageText = fromText(page.text || '');
      if (fromPageText) {
        return fromPageText;
      }
    } catch {
      continue;
    }
  }
  return '';
}

function ourRefWordIndex(words: IOcrWord[], index: number): number {
  const key = normalizeToken(words[index].text || '');
  if (/^ourr+e+fs?(no|number)?$/.test(key) || key === 'ourreference') {
    return index;
  }
  if (!isRefWord(key)) {
    return -1;
  }
  const prev = words[index - 1];
  const prevKey = prev ? normalizeToken(prev.text || '') : '';
  if (prevKey === 'our' || prevKey === '0ur' || prevKey === 'ou') {
    return index;
  }
  return -1;
}

function yourRefWordIndex(words: IOcrWord[], index: number): number {
  const key = normalizeToken(words[index].text || '');
  if (/^y(ou)?rr+e+fs?(no|number)?$/.test(key) || key === 'yourreference' || key === 'youref') {
    return index;
  }
  if (!isRefWord(key)) {
    return -1;
  }
  const prev = words[index - 1];
  const prevKey = prev ? normalizeToken(prev.text || '') : '';
  if (prevKey === 'your' || prevKey === 'you' || prevKey === 'yr' || prevKey === 'youf' || prevKey === 'yor') {
    return index;
  }
  return -1;
}

function plainRefWordIndex(words: IOcrWord[], index: number): number {
  const key = normalizeToken(words[index].text || '');
  if (/^ourr+e+fs?(no|number)?$/.test(key) || key === 'ourreference') {
    return -1;
  }
  if (/^y(ou)?rr+e+fs?(no|number)?$/.test(key) || key === 'yourreference' || key === 'youref') {
    return -1;
  }
  if (!isRefWord(key)) {
    return -1;
  }
  const prev = words[index - 1];
  const prevKey = prev ? normalizeToken(prev.text || '') : '';
  if (prevKey === 'our' || prevKey === '0ur' || prevKey === 'ou') {
    return -1;
  }
  if (prevKey === 'your' || prevKey === 'you' || prevKey === 'yr' || prevKey === 'youf' || prevKey === 'yor') {
    return -1;
  }
  if (!refLabelHasColon(words, index)) {
    return -1;
  }
  return index;
}

function refLabelHasColon(words: IOcrWord[], index: number): boolean {
  if (tokenHasColon(words[index].text || '')) {
    return true;
  }
  let cursor = index + 1;
  if (words[cursor] && /^[:.-]+$/.test((words[cursor].text || '').trim())) {
    return tokenHasColon(words[cursor].text || '');
  }
  const nextKey = words[cursor] ? normalizeToken(words[cursor].text || '') : '';
  if (nextKey === 'no' || nextKey === 'number') {
    if (tokenHasColon(words[cursor].text || '')) {
      return true;
    }
    cursor++;
    if (words[cursor] && /^[:.-]+$/.test((words[cursor].text || '').trim())) {
      return tokenHasColon(words[cursor].text || '');
    }
    return tokenHasColon(words[cursor] ? words[cursor].text || '' : '');
  }
  return false;
}

function tokenHasColon(text: string): boolean {
  return (text || '').indexOf(':') >= 0;
}

function isRefWord(key: string): boolean {
  return key === 'reference' ||
    key === 'refno' ||
    /^r+e+fs?(no|number)?$/.test(key);
}

function valueAfterRef(words: IOcrWord[], index: number): string {
  let start = index + 1;
  if (words[start] && /^[:.-]+$/.test((words[start].text || '').trim())) {
    start++;
  }
  const nextKey = words[start] ? normalizeToken(words[start].text || '') : '';
  if (nextKey === 'no' || nextKey === 'number') {
    start++;
    if (words[start] && /^[:.-]+$/.test((words[start].text || '').trim())) {
      start++;
    }
  }

  const label = words[index];
  const lineMid = (label.y0 + label.y1) / 2;
  const lineHeight = Math.max(label.y1 - label.y0, 1);
  const collected: IOcrWord[] = [];
  for (let cursor = start; cursor < words.length; cursor++) {
    const word = words[cursor];
    const wordMid = (word.y0 + word.y1) / 2;
    if (Math.abs(wordMid - lineMid) > lineHeight * 0.8) {
      break;
    }
    if (word.x0 < label.x1 - 6) {
      continue;
    }
    if (isRefValueStop(word.text || '')) {
      break;
    }
    collected.push(word);
  }
  const sameLine = joinOcrWords(collected).replace(/^[:.\s-]+/, '').trim();
  if (sameLine) {
    return sameLine;
  }

  const below: IOcrWord[] = [];
  for (let cursor = start; cursor < words.length; cursor++) {
    const word = words[cursor];
    if (word.y0 < label.y1 - lineHeight * 0.2) {
      continue;
    }
    if (word.y0 > label.y1 + lineHeight * 1.8) {
      break;
    }
    if (isRefValueStop(word.text || '')) {
      break;
    }
    below.push(word);
  }
  return joinOcrWords(below).replace(/^[:.\s-]+/, '').trim();
}

function isRefValueStop(text: string): boolean {
  const key = normalizeToken(text);
  return key === 'date' ||
    key === 'tel' ||
    key === 'fax' ||
    key === 'email' ||
    key === 'our' ||
    key === 'your' ||
    key === 'yr' ||
    key === 'page' ||
    key === 'dear' ||
    isRefLabelToken(key);
}

function isRefLabelToken(key: string): boolean {
  return /^(our|your|yr|my)?r+efs?(no|number)?$/.test(key) || key === 'reference' || key === 'referenceno';
}

function ourRefFromText(text: string): string {
  const lines = (text || '').split(/\r?\n/).map((line) => line.trim()).filter((line) => line.length > 0);
  const pattern = /\bour\s+r+e+f(?:erence)?(?:\s*no(?:\.)?)?\s*[:.-]?\s*(.+)$/i;
  for (let index = 0; index < lines.length; index++) {
    const match = lines[index].match(pattern);
    if (!match) {
      continue;
    }
    const value = (match[1] || '')
      .replace(/\b(your\s+r+e+f|date|tel|fax|email)\b.*$/i, '')
      .replace(/^[:.\s-]+/, '')
      .trim();
    if (value) {
      return value;
    }
  }
  return '';
}

function plainRefFromText(text: string): string {
  const lines = (text || '').split(/\r?\n/).map((line) => line.trim()).filter((line) => line.length > 0);
  const pattern = /\br+e+f(?:erence)?(?:\s*no(?:\.)?)?\s*:\s*(.+)$/i;
  for (let index = 0; index < lines.length; index++) {
    const line = lines[index];
    if (/\b(?:our|your?|yr)\s+r+e+f/i.test(line)) {
      continue;
    }
    const match = line.match(pattern);
    if (!match) {
      continue;
    }
    const value = (match[1] || '')
      .replace(/\b(our\s+r+e+f|your\s+r+e+f|date|tel|fax|email)\b.*$/i, '')
      .replace(/^[:.\s-]+/, '')
      .trim();
    if (value) {
      return value;
    }
  }
  return '';
}

function yourRefFromText(text: string): string {
  const lines = (text || '').split(/\r?\n/).map((line) => line.trim()).filter((line) => line.length > 0);
  const pattern = /\b(?:your?|yr)\s+r+e+f(?:erence)?(?:\s*no(?:\.)?)?\s*[:.-]?\s*(.+)$/i;
  for (let index = 0; index < lines.length; index++) {
    const match = lines[index].match(pattern);
    if (!match) {
      continue;
    }
    const value = (match[1] || '')
      .replace(/\b(our\s+r+e+f|date|tel|fax|email)\b.*$/i, '')
      .replace(/^[:.\s-]+/, '')
      .trim();
    if (value) {
      return value;
    }
  }
  return '';
}

function extractFromWords(words: IOcrWord[], fieldLabels: string[]): { [label: string]: string } {
  const values: { [label: string]: string } = {};
  if (words.length === 0) {
    return values;
  }

  const sorted = words.slice().sort((left, right) => {
    if (Math.abs(left.y0 - right.y0) > 8) {
      return left.y0 - right.y0;
    }
    return left.x0 - right.x0;
  });
  const tokens = sorted.map((word) => normalizeToken(word.text));
  const hits = findLabelHits(tokens, fieldLabels);
  const used = new Set<number>();

  hits.sort((left, right) => right.aliasLength - left.aliasLength);
  hits.forEach((hit) => {
    if (values[hit.field]) {
      return;
    }
    let overlap = false;
    for (let index = hit.start; index < hit.end; index++) {
      if (used.has(index)) {
        overlap = true;
        break;
      }
    }
    if (overlap) {
      return;
    }

    const valueWords = collectValueWords(sorted, hit, hits);
    const value = joinOcrWords(valueWords).replace(/^[:.\s-]+/, '').trim();
    if (value) {
      values[hit.field] = value;
      for (let index = hit.start; index < hit.end; index++) {
        used.add(index);
      }
    }
  });

  return values;
}

function extractFromText(text: string, fieldLabels: string[]): { [label: string]: string } {
  const values: { [label: string]: string } = {};
  const lines = text.split(/\r?\n/).map((line) => line.trim()).filter((line) => line.length > 0);

  const ordered = fieldLabels.slice().sort((left, right) => longestAliasLength(right) - longestAliasLength(left));
  ordered.forEach((label) => {
    if (values[label]) {
      return;
    }
    const aliases = aliasesFor(label);
    for (let lineIndex = 0; lineIndex < lines.length; lineIndex++) {
      const line = lines[lineIndex];
      const match = matchAliasOnLine(line, aliases);
      if (!match) {
        continue;
      }
      if (match.value) {
        values[label] = match.value;
        break;
      }
      const nextLine = lines[lineIndex + 1];
      if (nextLine && !lineStartsWithAnyAlias(nextLine, fieldLabels)) {
        values[label] = nextLine;
        break;
      }
    }
  });

  return values;
}

function findLabelHits(tokens: string[], fieldLabels: string[]): ILabelHit[] {
  const hits: ILabelHit[] = [];
  fieldLabels.forEach((label) => {
    aliasesFor(label).forEach((alias) => {
      const aliasTokens = tokenize(alias);
      if (aliasTokens.length === 0) {
        return;
      }
      for (let index = 0; index <= tokens.length - aliasTokens.length; index++) {
        if (sequenceEquals(tokens, index, aliasTokens)) {
          hits.push({
            field: label,
            aliasLength: aliasTokens.length,
            start: index,
            end: index + aliasTokens.length
          });
        }
      }
    });
  });
  return hits;
}

function collectValueWords(words: IOcrWord[], hit: ILabelHit, hits: ILabelHit[]): IOcrWord[] {
  const labelWord = words[hit.end - 1];
  const lineMid = (labelWord.y0 + labelWord.y1) / 2;
  const lineHeight = Math.max(labelWord.y1 - labelWord.y0, 1);
  const occupied = new Set<number>();
  hits.forEach((item) => {
    if (item.start === hit.start && item.end === hit.end) {
      return;
    }
    for (let index = item.start; index < item.end; index++) {
      occupied.add(index);
    }
  });

  const sameLine: IOcrWord[] = [];
  for (let index = hit.end; index < words.length; index++) {
    if (occupied.has(index)) {
      break;
    }
    const word = words[index];
    const wordMid = (word.y0 + word.y1) / 2;
    if (Math.abs(wordMid - lineMid) > lineHeight * 0.75) {
      break;
    }
    if (word.x0 < labelWord.x1 - 4) {
      continue;
    }
    sameLine.push(word);
  }
  if (sameLine.length > 0) {
    return sameLine;
  }

  const nextLine: IOcrWord[] = [];
  for (let index = hit.end; index < words.length; index++) {
    if (occupied.has(index)) {
      break;
    }
    const word = words[index];
    if (word.y0 < labelWord.y1 - lineHeight * 0.2) {
      continue;
    }
    if (word.y0 > labelWord.y1 + lineHeight * 1.8) {
      break;
    }
    nextLine.push(word);
  }
  return nextLine;
}

function sequenceEquals(tokens: string[], start: number, aliasTokens: string[]): boolean {
  for (let offset = 0; offset < aliasTokens.length; offset++) {
    if (!tokenEquals(tokens[start + offset], aliasTokens[offset])) {
      return false;
    }
  }
  return true;
}

function tokenEquals(left: string, right: string): boolean {
  if (left === right) {
    return true;
  }
  if ((left === 'no' || left === 'number') && (right === 'no' || right === 'number')) {
    return true;
  }
  if ((left === 'bl' || left === 'b') && (right === 'bl' || right === 'b')) {
    return true;
  }
  return false;
}

function matchAliasOnLine(line: string, aliases: string[]): { value: string } | undefined {
  const normalizedLine = normalizeKey(line);
  for (let index = 0; index < aliases.length; index++) {
    const alias = normalizeKey(aliases[index]);
    if (alias.length === 0) {
      continue;
    }
    if (normalizedLine === alias) {
      return { value: '' };
    }
    if (normalizedLine.indexOf(alias + ' ') === 0 || normalizedLine.indexOf(alias) === 0) {
      const originalIndex = indexOfNormalized(line, aliases[index]);
      if (originalIndex !== 0 && originalIndex !== -1) {
        continue;
      }
      const remainder = stripLeadingLabel(line, aliases[index]);
      if (remainder !== undefined) {
        return { value: remainder };
      }
    }
  }
  return undefined;
}

function lineStartsWithAnyAlias(line: string, fieldLabels: string[]): boolean {
  for (let fieldIndex = 0; fieldIndex < fieldLabels.length; fieldIndex++) {
    const aliases = aliasesFor(fieldLabels[fieldIndex]);
    if (matchAliasOnLine(line, aliases)) {
      return true;
    }
  }
  return false;
}

function stripLeadingLabel(line: string, alias: string): string | undefined {
  const lineChars = line.trim();
  const aliasChars = alias.trim();
  let linePos = 0;
  let aliasPos = 0;
  while (linePos < lineChars.length && aliasPos < aliasChars.length) {
    const lineChar = lineChars.charAt(linePos).toLowerCase();
    const aliasChar = aliasChars.charAt(aliasPos).toLowerCase();
    if (isIgnorable(lineChar)) {
      linePos++;
      continue;
    }
    if (isIgnorable(aliasChar)) {
      aliasPos++;
      continue;
    }
    if (lineChar !== aliasChar) {
      return undefined;
    }
    linePos++;
    aliasPos++;
  }
  let rest = lineChars.substring(linePos);
  rest = rest.replace(/^[\s:.-]+/, '').trim();
  return rest;
}

function indexOfNormalized(line: string, alias: string): number {
  const stripped = stripLeadingLabel(line, alias);
  return stripped === undefined ? -1 : 0;
}

function aliasesFor(label: string): string[] {
  const extra = FIELD_ALIASES[label] || [];
  const list = [label].concat(extra);
  return list.sort((left, right) => right.length - left.length);
}

function longestAliasLength(label: string): number {
  return aliasesFor(label).reduce((max, alias) => Math.max(max, alias.length), 0);
}

function tokenize(text: string): string[] {
  return normalizeKey(text).split(' ').filter((token) => token.length > 0);
}

function normalizeToken(text: string): string {
  return text.toLowerCase().replace(/[^a-z0-9]+/g, '');
}

function normalizeKey(text: string): string {
  return text.toLowerCase().replace(/[^a-z0-9]+/g, ' ').replace(/\s+/g, ' ').trim();
}

function isIgnorable(char: string): boolean {
  return char === ' ' || char === '.' || char === ':' || char === '-' || char === '/';
}
