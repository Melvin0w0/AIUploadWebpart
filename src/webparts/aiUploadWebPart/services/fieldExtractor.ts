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
    const layoutValue = (fromLayout[label] || '').trim();
    const textValue = (fromText[label] || '').trim();
    const chosen = layoutValue || textValue;
    if (chosen) {
      values[label] = chosen;
    }
  });

  return values;
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
