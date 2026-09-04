import { IOcrPageResult, IOcrWord } from './IPdfOcr';
import { joinOcrWords } from './ocrSelection';

interface IInkBand {
  pixels: Uint8ClampedArray;
  width: number;
  height: number;
  x0: number;
  y0: number;
}

export interface ISignatureRegion {
  x0: number;
  y0: number;
  x1: number;
  y1: number;
}

export interface ISignatureAnalysis {
  region: ISignatureRegion | undefined;
  senderName: string;
  textBelow: string;
}

export async function analyzeSignature(page?: IOcrPageResult): Promise<ISignatureAnalysis> {
  const fromClosing = extractNameBetweenClosingAndTitle(page);
  const empty: ISignatureAnalysis = {
    region: undefined,
    senderName: fromClosing,
    textBelow: ''
  };
  if (!page || !page.imageUrl || page.width <= 0 || page.height <= 0) {
    return empty;
  }

  try {
    const image = await loadImage(page.imageUrl);
    const regions = findInkSignatures(image, page.words || []);
    let chosen: ISignatureAnalysis = empty;
    for (let index = 0; index < regions.length; index++) {
      const region = regions[index];
      const textBelow = joinOcrWords(wordsBelowSignature(page, region));
      const senderName = firstPersonName(textBelow);
      if (senderName) {
        chosen = { region, textBelow, senderName };
      } else if (!chosen.region) {
        chosen = { region, textBelow, senderName: '' };
      }
    }
    if (chosen.senderName) {
      return chosen;
    }
    return {
      ...chosen,
      senderName: fromClosing
    };
  } catch {
    return empty;
  }
}

export function asPersonName(value: string): string {
  return personNameFromLine(value);
}

export function extractReceiverAboveDearSir(page?: IOcrPageResult): string {
  if (!page) {
    return '';
  }
  try {
    const fromAttn = extractAttnValue(page);
    if (fromAttn) {
      const cleaned = cleanReceiverName(fromAttn);
      if (cleaned) {
        return cleaned;
      }
    }
    const fromByHand = receiverBelowByHand(page);
    if (fromByHand) {
      return fromByHand;
    }
    return '';
  } catch {
    return '';
  }
}

export function extractOrganizationAboveAddressee(page?: IOcrPageResult): string {
  if (!page) {
    return '';
  }
  try {
    const fromByPost = departmentBelowByPost(page);
    if (fromByPost) {
      return fromByPost;
    }
    return '';
  } catch {
    return '';
  }
}

function departmentBelowByPost(page: IOcrPageResult): string {
  const lines = linesBelowDelivery(page, 'post');
  const departments = lines.filter((line) => isDepartmentLine(line) && !isDirectorLine(line));
  if (departments.length > 0) {
    return departments[0];
  }
  const anyDepartment = lines.filter((line) => isDepartmentLine(line));
  return anyDepartment.length > 0 ? anyDepartment[0] : '';
}

function receiverBelowByHand(page: IOcrPageResult): string {
  const lines = linesBelowDelivery(page, 'hand');
  for (let index = 0; index < lines.length; index++) {
    const name = cleanReceiverName(lines[index]);
    if (name && looksLikePersonName(name)) {
      return name;
    }
  }
  for (let index = 0; index < lines.length; index++) {
    const name = cleanReceiverName(lines[index]);
    if (name && !isDepartmentLine(name) && !isDirectorLine(name) && !isDeliveryLine(name)) {
      return name;
    }
  }
  return '';
}

function cleanReceiverName(value: string): string {
  let text = stripParenthetical(value);
  text = text
    .replace(/^(?:(?:mr|mrs|ms|miss|dr|ir|prof(?:essor)?|engr?|sir|madam|mdm|mx|messrs)\b\.?\s*)+/i, '')
    .replace(/\s*(?:先生|女士|小姐|太太)\s*$/g, '')
    .replace(/\s+/g, ' ')
    .trim();
  return text;
}

function stripParenthetical(value: string): string {
  let text = (value || '').trim();
  let previous = '';
  while (text !== previous) {
    previous = text;
    text = text
      .replace(/[(\uFF08][^)\uFF09]*[)\uFF09]/g, '')
      .replace(/[()\uFF08\uFF09]/g, '')
      .replace(/\s+/g, ' ')
      .trim();
  }
  return text;
}

function linesBelowDelivery(page: IOcrPageResult, kind: 'post' | 'hand'): string[] {
  const hit = findDeliveryHit(page, kind);
  if (!hit) {
    return linesBelowDeliveryFromText(page.text || '', kind);
  }
  const column = deliveryColumn(page, hit, kind);
  const dear = findSalutationHit(page.words || []);
  const stopY = dear ? dear.y0 - 2 : page.height;
  const columnWords = (page.words || []).filter((word) => {
    const midX = (word.x0 + word.x1) / 2;
    return midX >= column.left && midX <= column.right;
  });
  const lines = groupWordsIntoLines(columnWords);
  const collected: string[] = [];
  for (let index = 0; index < lines.length; index++) {
    const line = lines[index];
    if (line.y1 <= hit.y1 - 4) {
      continue;
    }
    if (line.y0 >= stopY) {
      break;
    }
    const text = (line.text || '').replace(/\s+/g, ' ').trim();
    if (!text || isDeliveryLine(text) || isAttnLine(text) || isSalutationLine(text)) {
      continue;
    }
    collected.push(text);
    if (collected.length >= 8) {
      break;
    }
  }
  return collected.length > 0 ? collected : linesBelowDeliveryFromText(page.text || '', kind);
}

function linesBelowDeliveryFromText(text: string, kind: 'post' | 'hand'): string[] {
  const lines = (text || '').split(/\r?\n/).map((line) => line.replace(/\s+/g, ' ').trim()).filter((line) => line.length > 0);
  const collected: string[] = [];
  let capturing = false;
  for (let index = 0; index < lines.length; index++) {
    const line = lines[index];
    if (isDeliveryKindLine(line, kind) || (kind === 'post' && isCombinedDeliveryLine(line)) || (kind === 'hand' && isCombinedDeliveryLine(line))) {
      capturing = true;
      continue;
    }
    if (!capturing) {
      continue;
    }
    if (isSalutationLine(line) || isAttnLine(line) || isDeliveryLine(line)) {
      break;
    }
    const key = normalizeKey(line);
    if (key.indexOf('our ref') === 0 || key.indexOf('your ref') === 0) {
      break;
    }
    collected.push(line);
    if (collected.length >= 8) {
      break;
    }
  }
  return collected;
}

function findDeliveryHit(
  page: IOcrPageResult,
  kind: 'post' | 'hand'
): { x0: number; y0: number; x1: number; y1: number } | undefined {
  const words = (page.words || []).slice().sort((left, right) => {
    if (Math.abs(left.y0 - right.y0) > 8) {
      return left.y0 - right.y0;
    }
    return left.x0 - right.x0;
  });
  for (let index = 0; index < words.length; index++) {
    const key = wordKey(words[index].text || '');
    const compact = kind === 'post' ? /^byp(o|e)?st$/ : /^byh(a|e)?nd$/;
    if (compact.test(key)) {
      return wordBox(words[index]);
    }
    if (key !== 'by' && key !== '8y') {
      continue;
    }
    const lineMid = (words[index].y0 + words[index].y1) / 2;
    const lineHeight = Math.max(10, words[index].y1 - words[index].y0);
    for (let cursor = index + 1; cursor < Math.min(words.length, index + 5); cursor++) {
      const next = words[cursor];
      if (Math.abs((next.y0 + next.y1) / 2 - lineMid) > lineHeight * 0.85) {
        break;
      }
      const nextKey = wordKey(next.text || '');
      if (kind === 'post' && isPostWord(nextKey)) {
        return mergeBoxes(wordBox(words[index]), wordBox(next));
      }
      if (kind === 'hand' && isHandWord(nextKey)) {
        return mergeBoxes(wordBox(words[index]), wordBox(next));
      }
    }
  }
  return undefined;
}

function deliveryColumn(
  page: IOcrPageResult,
  hit: { x0: number; y0: number; x1: number; y1: number },
  kind: 'post' | 'hand'
): { left: number; right: number } {
  const other = findDeliveryHit(page, kind === 'post' ? 'hand' : 'post');
  if (other && Math.abs(other.y0 - hit.y0) < Math.max(24, (hit.y1 - hit.y0) * 2)) {
    if (hit.x0 <= other.x0) {
      return { left: 0, right: (hit.x1 + other.x0) / 2 };
    }
    return { left: (other.x1 + hit.x0) / 2, right: page.width };
  }
  return {
    left: Math.max(0, hit.x0 - 28),
    right: Math.min(page.width, Math.max(hit.x1 + 90, hit.x0 + page.width * 0.45))
  };
}

function wordBox(word: IOcrWord): { x0: number; y0: number; x1: number; y1: number } {
  return { x0: word.x0, y0: word.y0, x1: word.x1, y1: word.y1 };
}

function mergeBoxes(
  left: { x0: number; y0: number; x1: number; y1: number },
  right: { x0: number; y0: number; x1: number; y1: number }
): { x0: number; y0: number; x1: number; y1: number } {
  return {
    x0: Math.min(left.x0, right.x0),
    y0: Math.min(left.y0, right.y0),
    x1: Math.max(left.x1, right.x1),
    y1: Math.max(left.y1, right.y1)
  };
}

function isPostWord(key: string): boolean {
  return key === 'post' || key === 'pest' || key === 'pst' || key === 'postage';
}

function isHandWord(key: string): boolean {
  return key === 'hand' || key === 'hnd' || key === 'hands';
}

function isDeliveryLine(line: string): boolean {
  return isDeliveryKindLine(line, 'post') || isDeliveryKindLine(line, 'hand') || isCombinedDeliveryLine(line);
}

function isCombinedDeliveryLine(line: string): boolean {
  const key = normalizeKey(line);
  return /\bby\s+hand\b/.test(key) && /\bby\s+post\b/.test(key);
}

function isDeliveryKindLine(line: string, kind: 'post' | 'hand'): boolean {
  const key = normalizeKey(line);
  if (kind === 'post') {
    return /^by\s+(post|pest|pst)\b/.test(key) || /\bby\s+post\b/.test(key);
  }
  return /^by\s+hand\b/.test(key) || /\bby\s+hand\b/.test(key);
}

function isDepartmentLine(line: string): boolean {
  const text = (line || '').replace(/\s+/g, ' ').trim();
  return /\bdepartments?\b/i.test(text) || /(?:署|處)\s*$/.test(text);
}

function isDirectorLine(line: string): boolean {
  return /^(?:the\s+)?directors?\b/i.test((line || '').replace(/\s+/g, ' ').trim());
}

function organizationAboveAttn(page: IOcrPageResult): string {
  const fromWords = firstLineAboveAttnFromWords(page);
  if (fromWords) {
    return fromWords;
  }
  return firstLineAboveAttnFromText(page.text || '');
}

function isAttnLine(line: string): boolean {
  return /^(?:for\s+the\s+)?(?:att(?:n|ention)|atin|attm|atln)\b/i.test((line || '').replace(/\s+/g, ' ').trim());
}

function firstLineAboveAttnFromWords(page: IOcrPageResult): string {
  const lines = groupWordsIntoLines(page.words || []);
  if (lines.length === 0) {
    return '';
  }
  const dear = findSalutationHit(page.words || []);
  const attnIndex = findAttnLineIndex(lines, dear ? dear.y0 : -1);
  if (attnIndex < 0) {
    return '';
  }
  const above: string[] = [];
  let nextBottom = lines[attnIndex].y0;
  for (let index = attnIndex - 1; index >= 0 && above.length < 4; index--) {
    const line = lines[index];
    const text = completeAddresseeLine(page, line);
    if (!text || isAttnLine(text) || isSalutationLine(text)) {
      continue;
    }
    if (isAddressBlockStop(text)) {
      break;
    }
    const lineHeight = Math.max(12, line.y1 - line.y0);
    if (nextBottom - line.y1 > lineHeight * 2.8) {
      break;
    }
    above.unshift(text);
    nextBottom = line.y0;
  }
  return above.length > 0 ? above[0] : '';
}

function firstLineAboveAttnFromText(text: string): string {
  const lines = (text || '').split(/\r?\n/).map((line) => (line || '').replace(/\s+/g, ' ').trim());
  let attnIndex = -1;
  for (let index = 0; index < lines.length; index++) {
    if (isAttnLine(lines[index])) {
      attnIndex = index;
      break;
    }
  }
  if (attnIndex < 0) {
    return '';
  }
  const above: string[] = [];
  for (let index = attnIndex - 1; index >= 0 && above.length < 4; index--) {
    const line = lines[index];
    if (!line) {
      if (above.length > 0) {
        break;
      }
      continue;
    }
    if (isAttnLine(line) || isSalutationLine(line)) {
      continue;
    }
    if (isAddressBlockStop(line)) {
      break;
    }
    above.unshift(line);
  }
  return above.length > 0 ? above[0] : '';
}

function findAttnLineIndex(lines: { text: string; y0: number; y1: number }[], dearY: number): number {
  const hits: number[] = [];
  for (let index = 0; index < lines.length; index++) {
    if (isAttnLine(lines[index].text)) {
      hits.push(index);
    }
  }
  if (hits.length === 0) {
    return -1;
  }
  if (dearY >= 0) {
    const aboveDear = hits.filter((index) => lines[index].y1 <= dearY + 6);
    if (aboveDear.length > 0) {
      return aboveDear[aboveDear.length - 1];
    }
  }
  return hits[0];
}

function organizationFromLines(lines: string[]): string {
  for (let index = 0; index < lines.length; index++) {
    const line = (lines[index] || '').replace(/\s+/g, ' ').trim();
    if (!line || isSalutationLine(line) || isAddressBlockStop(line)) {
      continue;
    }
    if (/^(?:for\s+the\s+)?att(?:n|ention)\b/i.test(line)) {
      continue;
    }
    if (!/^[A-Za-z]/.test(line)) {
      continue;
    }
    return line;
  }
  return '';
}

export async function extractSubjectBelowDearSir(page?: IOcrPageResult): Promise<string> {
  if (!page) {
    return '';
  }
  const fromFirstUnderline = subjectFromFirstUnderlinedLine(page);
  if (fromFirstUnderline) {
    return fromFirstUnderline;
  }
  const fromUnderline = await subjectFromUnderline(page);
  if (fromUnderline) {
    return fromUnderline;
  }
  const fromBold = subjectFromStyledWords(page, 'bold');
  if (fromBold) {
    return fromBold;
  }
  return subjectFromReBlock(page);
}

export function subjectAppearsInPage(page: IOcrPageResult, value: string): boolean {
  const needle = normalizeSubjectMatch(value);
  if (needle.length < 4) {
    return false;
  }
  const haystacks = [
    normalizeSubjectMatch(page.text || ''),
    normalizeSubjectMatch(joinOcrWords(page.words || []))
  ];
  for (let index = 0; index < haystacks.length; index++) {
    if (haystacks[index].indexOf(needle) >= 0) {
      return true;
    }
  }
  return false;
}

function normalizeSubjectMatch(text: string): string {
  return (text || '')
    .toLowerCase()
    .replace(/[\u2018\u2019\u201c\u201d']/g, '')
    .replace(/[^a-z0-9\u4e00-\u9fff]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

export function dearSirBandRegion(page?: IOcrPageResult): ISignatureRegion | undefined {
  if (!page || page.width <= 0 || page.height <= 0) {
    return undefined;
  }
  const hit = findSalutationHit(page.words || []);
  const closing = findClosingHit(page.words || []);
  if (!hit) {
    return {
      x0: 0,
      y0: Math.round(page.height * 0.12),
      x1: Math.round(page.width * 0.96),
      y1: closing ? Math.max(8, closing.y0) : Math.round(page.height * 0.68)
    };
  }
  const y1 = closing
    ? Math.max(hit.y1 + 8, closing.y0)
    : Math.min(page.height, Math.max(hit.y1 + Math.max(160, page.height * 0.42), page.height * 0.68));
  return {
    x0: 0,
    y0: hit.y1,
    x1: Math.round(page.width * 0.96),
    y1: Math.min(page.height, y1)
  };
}

function extractAttnValue(page: IOcrPageResult): string {
  const fromWords = attnFromWords(page);
  if (fromWords) {
    return fromWords;
  }
  return attnFromText(page.text || '');
}

function attnFromWords(page: IOcrPageResult): string {
  const words = (page.words || []).slice().sort((left, right) => {
    if (Math.abs(left.y0 - right.y0) > 8) {
      return left.y0 - right.y0;
    }
    return left.x0 - right.x0;
  });
  for (let index = 0; index < words.length; index++) {
    if (!isAttnLabelWord(words, index)) {
      continue;
    }
    const value = valueAfterAttn(words, index);
    if (value) {
      return value;
    }
  }
  return '';
}

function isAttnLabelWord(words: IOcrWord[], index: number): boolean {
  const key = wordKey(words[index].text || '');
  if (isAttnKey(key)) {
    return true;
  }
  if (key !== 'of') {
    return false;
  }
  const prev = words[index - 1];
  const prevKey = prev ? wordKey(prev.text || '') : '';
  return prevKey === 'attention';
}

function isAttnKey(key: string): boolean {
  return key === 'attn' ||
    key === 'atin' ||
    key === 'attm' ||
    key === 'atln' ||
    key === 'attention';
}

function valueAfterAttn(words: IOcrWord[], index: number): string {
  let start = index + 1;
  while (words[start]) {
    const raw = (words[start].text || '').trim();
    const key = wordKey(raw);
    if (/^[:.-]+$/.test(raw) || key === 'of' || key === 'to') {
      start++;
      continue;
    }
    break;
  }

  const label = words[index];
  const lineMid = (label.y0 + label.y1) / 2;
  const lineHeight = Math.max(label.y1 - label.y0, 1);
  const collected: IOcrWord[] = [];
  for (let cursor = start; cursor < words.length; cursor++) {
    const word = words[cursor];
    const wordMid = (word.y0 + word.y1) / 2;
    if (Math.abs(wordMid - lineMid) > lineHeight * 0.85) {
      break;
    }
    if (word.x0 < label.x1 - 6) {
      continue;
    }
    if (isAttnValueStop(word.text || '')) {
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
    if (isAttnValueStop(word.text || '')) {
      break;
    }
    below.push(word);
  }
  return joinOcrWords(below).replace(/^[:.\s-]+/, '').trim();
}

function isAttnValueStop(text: string): boolean {
  const key = wordKey(text);
  return key === 'date' ||
    key === 'tel' ||
    key === 'fax' ||
    key === 'email' ||
    key === 'dear' ||
    key === 'our' ||
    key === 'your' ||
    key === 'page' ||
    isAttnKey(key);
}

function attnFromText(text: string): string {
  const lines = (text || '').split(/\r?\n/).map((line) => line.trim()).filter((line) => line.length > 0);
  const pattern = /(?:for\s+the\s+)?att(?:n|ention|in|m)\s*(?:of)?\s*[:.-]?\s*(.+)$/i;
  for (let index = 0; index < lines.length; index++) {
    const match = lines[index].match(pattern);
    if (!match) {
      continue;
    }
    const value = (match[1] || '')
      .replace(/\b(date|tel|fax|email|dear|our\s+ref|your\s+ref)\b.*$/i, '')
      .replace(/^[:.\s-]+/, '')
      .trim();
    if (value) {
      return value;
    }
    const next = lines[index + 1];
    if (next && !isSalutationLine(next) && !isAddressBlockStop(next) && !/^att(?:n|ention)/i.test(next)) {
      return next.replace(/\s+/g, ' ').trim();
    }
  }
  return '';
}

function wordKey(text: string): string {
  return (text || '').toLowerCase().replace(/[^a-z0-9]+/g, '');
}

function addresseeBlockFromText(text: string): string {
  const raw = (text || '').split(/\r?\n/).map((line) => line.trim());
  let dearIndex = -1;
  for (let index = 0; index < raw.length; index++) {
    if (isSalutationLine(raw[index])) {
      dearIndex = index;
      break;
    }
  }
  if (dearIndex < 0) {
    const salutation = findSalutationIndex(text || '');
    if (salutation.index < 0) {
      return '';
    }
    const before = (text || '').substring(0, salutation.index).split(/\r?\n/).map((line) => line.trim());
    return firstLineOfConsecutiveBlock(before);
  }
  return firstLineOfConsecutiveBlock(raw.slice(0, dearIndex));
}

function addresseeBlockFromWords(page: IOcrPageResult): string {
  const block = consecutiveAddresseeLines(page);
  return block.length > 0 ? block[0].text : '';
}

function consecutiveAddresseeLines(page: IOcrPageResult): { text: string; y0: number; y1: number }[] {
  const hit = findSalutationHit(page.words || []);
  if (!hit) {
    return [];
  }
  const above = (page.words || []).filter((word) => {
    const midX = (word.x0 + word.x1) / 2;
    const midY = (word.y0 + word.y1) / 2;
    return midY < hit.y0 - 1 &&
      midY >= hit.y0 - Math.max(380, page.height * 0.58) &&
      midX <= page.width * 0.92;
  });
  const lines = groupWordsIntoLines(above).filter((line) => line.y1 < hit.y0);
  const block: { text: string; y0: number; y1: number }[] = [];
  let nextTop = hit.y0;
  for (let index = lines.length - 1; index >= 0; index--) {
    const line = lines[index];
    if (isSalutationLine(line.text)) {
      continue;
    }
    if (isAddressBlockStop(line.text)) {
      break;
    }
    const lineHeight = Math.max(12, line.y1 - line.y0);
    if (block.length > 0 && nextTop - line.y1 > lineHeight * 2.8) {
      break;
    }
    block.unshift(line);
    nextTop = line.y0;
    if (block.length >= 12) {
      break;
    }
  }
  return block;
}

function completeAddresseeLine(
  page: IOcrPageResult,
  line: { text: string; y0: number; y1: number }
): string {
  const midY = (line.y0 + line.y1) / 2;
  const lineHeight = Math.max(10, line.y1 - line.y0);
  const onLine = (page.words || []).filter((word) => {
    const wordMid = (word.y0 + word.y1) / 2;
    return Math.abs(wordMid - midY) <= lineHeight * 0.7;
  }).sort((left, right) => left.x0 - right.x0);

  if (onLine.length === 0) {
    return (line.text || '').replace(/\s+/g, ' ').trim();
  }

  const leftCluster: IOcrWord[] = [];
  const gapLimit = Math.max(40, page.width * 0.065);
  for (let index = 0; index < onLine.length; index++) {
    const word = onLine[index];
    if (isRightColumnLabel(onLine, index)) {
      break;
    }
    if (leftCluster.length > 0) {
      const previous = leftCluster[leftCluster.length - 1];
      if (word.x0 - previous.x1 > gapLimit) {
        break;
      }
    }
    leftCluster.push(word);
  }

  const completed = joinOcrWords(leftCluster).replace(/\s+/g, ' ').trim();
  const original = (line.text || '').replace(/\s+/g, ' ').trim();
  return completed.length >= original.length ? completed : original;
}

function isRightColumnLabel(words: IOcrWord[], index: number): boolean {
  const key = wordKey(words[index].text || '');
  if (key === 'date' || key === 'tel' || key === 'fax' || key === 'email' || key === 'page') {
    return true;
  }
  const next = words[index + 1];
  const nextKey = next ? wordKey(next.text || '') : '';
  if ((key === 'our' || key === 'your' || key === 'you' || key === 'yr') &&
    (/^r+e+fs?(no|number)?$/.test(nextKey) || nextKey === 'reference')) {
    return true;
  }
  if (/^(our|your)?r+e+fs?(no|number)?$/.test(key) || key === 'reference' || key === 'ourreference' || key === 'yourreference') {
    const prev = words[index - 1];
    const prevKey = prev ? wordKey(prev.text || '') : '';
    if (key.indexOf('our') === 0 || key.indexOf('your') === 0) {
      return true;
    }
    return prevKey === 'our' || prevKey === 'your' || prevKey === 'you' || prevKey === 'yr';
  }
  const prev = words[index - 1];
  const prevKey = prev ? wordKey(prev.text || '') : '';
  return (key === 'ref' || key === 'reference') &&
    (prevKey === 'our' || prevKey === 'your' || prevKey === 'you' || prevKey === 'yr');
}

function consecutiveBlockFromText(text: string): string[] {
  const raw = (text || '').split(/\r?\n/).map((line) => line.trim());
  let dearIndex = -1;
  for (let index = 0; index < raw.length; index++) {
    if (isSalutationLine(raw[index])) {
      dearIndex = index;
      break;
    }
  }
  if (dearIndex < 0) {
    const salutation = findSalutationIndex(text || '');
    if (salutation.index < 0) {
      return [];
    }
    const before = (text || '').substring(0, salutation.index).split(/\r?\n/).map((line) => line.trim());
    return consecutiveLinesFromList(before);
  }
  return consecutiveLinesFromList(raw.slice(0, dearIndex));
}

function firstLineOfConsecutiveBlock(lines: string[]): string {
  const block = consecutiveLinesFromList(lines);
  return block.length > 0 ? block[0] : '';
}

function consecutiveLinesFromList(lines: string[]): string[] {
  const block: string[] = [];
  let sawGap = false;
  for (let index = lines.length - 1; index >= 0; index--) {
    const line = lines[index];
    if (!line) {
      if (block.length > 0) {
        break;
      }
      if (sawGap) {
        break;
      }
      sawGap = true;
      continue;
    }
    if (isSalutationLine(line)) {
      continue;
    }
    if (isAddressBlockStop(line)) {
      break;
    }
    block.unshift(line.replace(/\s+/g, ' ').trim());
    if (block.length >= 12) {
      break;
    }
  }
  return block;
}

function groupWordsIntoLines(words: IOcrWord[]): { text: string; x0: number; x1: number; y0: number; y1: number; words: IOcrWord[] }[] {
  const sorted = words.slice().sort((left, right) => {
    if (Math.abs(left.y0 - right.y0) > 8) {
      return left.y0 - right.y0;
    }
    return left.x0 - right.x0;
  });
  const groups: IOcrWord[][] = [];
  sorted.forEach((word) => {
    const last = groups[groups.length - 1];
    if (!last) {
      groups.push([word]);
      return;
    }
    const lastMid = (last[0].y0 + last[0].y1) / 2;
    const wordMid = (word.y0 + word.y1) / 2;
    const lineHeight = Math.max(last[0].y1 - last[0].y0, 10);
    if (Math.abs(wordMid - lastMid) <= lineHeight * 0.6) {
      last.push(word);
      return;
    }
    groups.push([word]);
  });
  return groups.map((group) => {
    let x0 = group[0].x0;
    let x1 = group[0].x1;
    let y0 = group[0].y0;
    let y1 = group[0].y1;
    group.forEach((word) => {
      x0 = Math.min(x0, word.x0);
      x1 = Math.max(x1, word.x1);
      y0 = Math.min(y0, word.y0);
      y1 = Math.max(y1, word.y1);
    });
    return {
      text: joinOcrWords(group).replace(/\s+/g, ' ').trim(),
      x0,
      x1,
      y0,
      y1,
      words: group
    };
  }).filter((line) => line.text.length > 0);
}

function lineHasUnderline(line: { words: IOcrWord[] }): boolean {
  for (let index = 0; index < line.words.length; index++) {
    if (line.words[index].underline) {
      return true;
    }
  }
  return false;
}

function subjectFromFirstUnderlinedLine(page: IOcrPageResult): string {
  const pageWords = page.words || [];
  const lines = groupWordsIntoLines(pageWords);
  const dear = findSalutationHit(pageWords);
  const closing = findClosingHit(pageWords);
  const block: IOcrWord[] = [];
  let collecting = false;
  for (let index = 0; index < lines.length; index++) {
    const line = lines[index];
    if (dear && line.y1 <= dear.y1) {
      continue;
    }
    if (closing && line.y0 >= closing.y0) {
      break;
    }
    if (isSalutationLine(line.text) || isClosingLine(line.text)) {
      if (collecting) {
        break;
      }
      continue;
    }
    if (!lineHasUnderline(line)) {
      if (collecting) {
        break;
      }
      continue;
    }
    collecting = true;
    for (let wordIndex = 0; wordIndex < line.words.length; wordIndex++) {
      block.push(line.words[wordIndex]);
    }
  }
  if (block.length === 0) {
    return '';
  }
  return stripSubjectLabel(joinOcrWords(block)).replace(/\s+/g, ' ').trim();
}

function subjectFromStyledWords(page: IOcrPageResult, style: 'underline' | 'bold'): string {
  const pageWords = page.words || [];
  const lines = groupWordsIntoLines(pageWords);
  const dear = findSalutationHit(pageWords);
  const closing = findClosingHit(pageWords);
  const hits: { words: IOcrWord[]; y0: number; y1: number }[] = [];
  for (let index = 0; index < lines.length; index++) {
    const line = lines[index];
    if (dear && line.y1 <= dear.y1) {
      continue;
    }
    if (closing && line.y0 >= closing.y0) {
      continue;
    }
    if (isSalutationLine(line.text) || isClosingLine(line.text) || isAddressBlockStop(line.text)) {
      continue;
    }
    const styled = style === 'underline'
      ? line.words.filter((word) => !!word.underline)
      : line.words.filter((word) => !!word.bold);
    if (styled.length === 0) {
      continue;
    }
    const used = style === 'bold' ? preferBoldWords(line.words) : styled;
    const text = stripSubjectLabel(joinOcrWords(used));
    if (style === 'bold') {
      if (!looksLikeBoldSubject(text)) {
        continue;
      }
    } else if (!looksLikeUnderlinedSubject(text)) {
      continue;
    }
    hits.push({ words: used, y0: line.y0, y1: line.y1 });
  }
  if (hits.length === 0) {
    return '';
  }
  const groupGap = Math.max(20, (page.height || 0) * 0.018);
  const groups: { words: IOcrWord[]; y1: number }[] = [];
  hits.forEach((item) => {
    const last = groups[groups.length - 1];
    if (last && item.y0 - last.y1 <= groupGap) {
      last.words = last.words.concat(item.words);
      last.y1 = item.y1;
      return;
    }
    groups.push({ words: item.words.slice(), y1: item.y1 });
  });
  return stripSubjectLabel(joinOcrWords(groups[0].words)).replace(/\s+/g, ' ').trim();
}

async function subjectFromUnderline(page: IOcrPageResult): Promise<string> {
  if (!page.imageUrl) {
    return '';
  }

  try {
    const image = await loadImage(page.imageUrl);
    const canvas = document.createElement('canvas');
    canvas.width = image.width;
    canvas.height = image.height;
    const context = canvas.getContext('2d', { willReadFrequently: true });
    if (!context) {
      return '';
    }
    context.drawImage(image, 0, 0);

    const band = subjectScanBand(page, image.width, image.height);
    const x0 = band.x0;
    const x1 = band.x1;
    const y0 = band.y0;
    const y1 = band.y1;
    const width = x1 - x0;
    const height = y1 - y0;
    if (width < 24 || height < 8) {
      canvas.width = 0;
      canvas.height = 0;
      return '';
    }

    const pixels = context.getImageData(x0, y0, width, height).data;
    canvas.width = 0;
    canvas.height = 0;

    const dear = findSalutationHit(page.words || []);
    const closing = findClosingHit(page.words || []);
    const ink: IInkBand = { pixels, width, height, x0, y0 };
    const dearY = dear ? dear.y1 : -1;
    const closingY = closing ? closing.y0 : -1;
    const thresholds = [110, 130, 150, 175, 195];
    for (let index = 0; index < thresholds.length; index++) {
      const found = headingsFromUnderlinePixels(
        page,
        pixels,
        width,
        height,
        x0,
        y0,
        image.width,
        image.height,
        thresholds[index],
        dearY,
        closingY,
        ink
      );
      if (found) {
        return found;
      }
    }
    return '';
  } catch {
    return '';
  }
}

function headingsFromUnderlinePixels(
  page: IOcrPageResult,
  pixels: Uint8ClampedArray,
  width: number,
  height: number,
  x0: number,
  y0: number,
  _pageWidth: number,
  _pageHeight: number,
  maxLum: number,
  dearY: number,
  closingY: number,
  _ink: IInkBand
): string {
  const pageWords = page.words || [];
  const lines = groupWordsIntoLines(pageWords);
  const block: IOcrWord[] = [];
  let collecting = false;
  for (let index = 0; index < lines.length; index++) {
    const line = lines[index];
    if (dearY >= 0 && line.y1 <= dearY) {
      continue;
    }
    if (closingY >= 0 && line.y0 >= closingY) {
      break;
    }
    if (isSalutationLine(line.text) || isClosingLine(line.text) || isAddressBlockStop(line.text)) {
      if (collecting) {
        break;
      }
      continue;
    }
    const span = underlineSpanUnderLine(line, pixels, x0, y0, width, height, maxLum);
    if (!span) {
      if (collecting) {
        break;
      }
      continue;
    }
    const text = stripSubjectLabel(line.text);
    if (!looksLikeUnderlinedSubject(text)) {
      if (collecting) {
        break;
      }
      continue;
    }
    collecting = true;
    line.words.forEach((word) => {
      const overlapLeft = Math.max(x0 + span.left, word.x0);
      const overlapRight = Math.min(x0 + span.right, word.x1);
      if (overlapRight - overlapLeft >= Math.max(4, (word.x1 - word.x0) * 0.25)) {
        word.underline = true;
      }
      block.push(word);
    });
  }
  if (block.length === 0) {
    return '';
  }
  return stripSubjectLabel(joinOcrWords(block)).replace(/\s+/g, ' ').trim();
}

function underlineSpanUnderLine(
  line: { x0: number; x1: number; y0: number; y1: number },
  pixels: Uint8ClampedArray,
  bandX0: number,
  bandY0: number,
  bandWidth: number,
  bandHeight: number,
  maxLum: number
): { left: number; right: number } | undefined {
  const lineHeight = Math.max(8, line.y1 - line.y0);
  const row0 = Math.floor(line.y1 - bandY0);
  const row1 = Math.floor(line.y1 - bandY0 + Math.max(4, lineHeight * 0.4));
  const col0 = Math.floor(line.x0 - bandX0);
  const col1 = Math.floor(line.x1 - bandX0);
  if (col1 - col0 < 8) {
    return undefined;
  }

  let best: { left: number; right: number; overlap: number } | undefined;
  const lineWidth = col1 - col0;
  for (let row = row0; row <= row1; row++) {
    if (row < 0 || row >= bandHeight) {
      continue;
    }
    const spans = darkSpansOnRow(pixels, bandWidth, row, maxLum);
    for (let spanIndex = 0; spanIndex < spans.length; spanIndex++) {
      const span = spans[spanIndex];
      const left = Math.max(col0 - 6, span.left);
      const right = Math.min(col1 + 6, span.right);
      const overlap = right - left;
      if (overlap < lineWidth * 0.35) {
        continue;
      }
      if (!best || overlap > best.overlap) {
        best = { left: span.left, right: span.right, overlap };
      }
    }
  }
  return best ? { left: best.left, right: best.right } : undefined;
}

function clusterUnderlineFragments(
  fragments: { y: number; left: number; right: number }[],
  maxUnderlineHeight: number
): { y0: number; y1: number; left: number; right: number }[] {
  const clustered: { y0: number; y1: number; left: number; right: number }[] = [];
  fragments.forEach((fragment) => {
    let matched: { y0: number; y1: number; left: number; right: number } | undefined;
    for (let index = clustered.length - 1; index >= 0; index--) {
      const item = clustered[index];
      if (fragment.y - item.y1 > 3) {
        break;
      }
      const gap = fragment.left > item.right
        ? fragment.left - item.right
        : item.left > fragment.right
          ? item.left - fragment.right
          : 0;
      if (gap <= 32 && fragment.y - item.y0 <= maxUnderlineHeight) {
        matched = item;
        break;
      }
    }
    if (!matched) {
      clustered.push({ y0: fragment.y, y1: fragment.y, left: fragment.left, right: fragment.right });
      return;
    }
    matched.y1 = fragment.y;
    matched.left = Math.min(matched.left, fragment.left);
    matched.right = Math.max(matched.right, fragment.right);
  });
  clustered.sort((left, right) => {
    if (Math.abs(left.y0 - right.y0) > 4) {
      return left.y0 - right.y0;
    }
    return left.left - right.left;
  });
  return clustered;
}

function subjectScanBand(page: IOcrPageResult, width: number, height: number): ISignatureRegion {
  const hit = findSalutationHit(page.words || []);
  const closing = findClosingHit(page.words || []);
  const linePad = Math.max(4, Math.round(height * 0.006));
  let y0 = height * 0.22;
  if (hit) {
    y0 = hit.y1 + linePad;
  }
  y0 = Math.max(0, Math.floor(y0));
  let y1 = Math.floor(height * 0.88);
  if (closing) {
    y1 = Math.min(y1, Math.floor(closing.y0 - linePad));
  }
  return {
    x0: Math.floor(width * 0.04),
    y0,
    x1: Math.floor(width * 0.96),
    y1: Math.max(y0 + 8, y1)
  };
}

function reLabelTop(page: IOcrPageResult): number {
  const line = chooseReLine(page);
  return line ? line.y0 : -1;
}

function chooseReLine(page: IOcrPageResult): { text: string; y0: number; y1: number } | undefined {
  const lines = groupWordsIntoLines(page.words || []);
  if (lines.length === 0) {
    return undefined;
  }
  const dear = findSalutationHit(page.words || []);
  const hits: number[] = [];
  for (let index = 0; index < lines.length; index++) {
    if (matchReLabel(lines[index].text) !== undefined) {
      hits.push(index);
    }
  }
  if (hits.length === 0) {
    return undefined;
  }
  let chosen = hits[0];
  if (dear) {
    const below = hits.filter((index) => lines[index].y0 >= dear.y1 - 6);
    if (below.length > 0) {
      chosen = below[0];
    } else {
      const above = hits.filter((index) => lines[index].y1 <= dear.y0 + 6);
      if (above.length > 0) {
        chosen = above[above.length - 1];
      }
    }
  }
  return lines[chosen];
}

function darkSpansOnRow(
  pixels: Uint8ClampedArray,
  width: number,
  row: number,
  maxLum: number
): { left: number; right: number }[] {
  const spans: { left: number; right: number }[] = [];
  const minRun = Math.max(8, Math.round(width * 0.015));
  let runLeft = -1;
  let darkInRun = 0;

  const finishRun = (end: number): void => {
    if (runLeft < 0) {
      return;
    }
    const runWidth = end - runLeft;
    const coverage = darkInRun / Math.max(1, runWidth);
    if (runWidth >= minRun && coverage > 0.16) {
      spans.push({ left: runLeft, right: end - 1 });
    }
    runLeft = -1;
    darkInRun = 0;
  };

  for (let col = 0; col < width; col++) {
    const index = (row * width + col) * 4;
    const lum = 0.299 * pixels[index] + 0.587 * pixels[index + 1] + 0.114 * pixels[index + 2];
    if (lum < maxLum) {
      if (runLeft < 0) {
        runLeft = col;
        darkInRun = 0;
      }
      darkInRun++;
    } else if (runLeft >= 0 && col - runLeft - darkInRun > 5) {
      finishRun(col);
    }
  }
  finishRun(width);
  return mergeCloseSpans(spans, 22);
}

function mergeCloseSpans(
  spans: { left: number; right: number }[],
  maxGap: number
): { left: number; right: number }[] {
  if (spans.length === 0) {
    return spans;
  }
  const merged: { left: number; right: number }[] = [spans[0]];
  for (let index = 1; index < spans.length; index++) {
    const last = merged[merged.length - 1];
    const next = spans[index];
    if (next.left - last.right <= maxGap) {
      last.right = next.right;
    } else {
      merged.push({ left: next.left, right: next.right });
    }
  }
  return merged;
}


function preferBoldWords(words: IOcrWord[], ink?: IInkBand): IOcrWord[] {
  if (!words || words.length === 0) {
    return words;
  }
  const marked = words.filter((word) => !!word.bold);
  if (marked.length > 0 && marked.length < words.length) {
    return marked;
  }
  if (marked.length === words.length) {
    return words;
  }
  if (!ink || words.length < 2) {
    return words;
  }
  const scored = words.map((word) => ({ word, density: wordInkDensity(ink, word) }));
  const densities = scored.map((item) => item.density).sort((left, right) => left - right);
  const median = densities[Math.floor(densities.length / 2)] || 0;
  const heaviest = densities[densities.length - 1] || 0;
  const lightest = densities[0] || 0;
  if (heaviest - lightest < 0.1) {
    return words;
  }
  const bold = scored.filter((item) => item.density >= Math.max(median, (median + heaviest) / 2));
  if (bold.length === 0 || bold.length === words.length) {
    return words;
  }
  return bold.map((item) => item.word);
}

function averageInkDensity(ink: IInkBand, words: IOcrWord[]): number {
  if (words.length === 0) {
    return 0;
  }
  let sum = 0;
  for (let index = 0; index < words.length; index++) {
    sum += wordInkDensity(ink, words[index]);
  }
  return sum / words.length;
}

function wordInkDensity(ink: IInkBand, word: IOcrWord): number {
  const left = Math.max(0, Math.floor(word.x0 - ink.x0));
  const top = Math.max(0, Math.floor(word.y0 - ink.y0));
  const right = Math.min(ink.width, Math.ceil(word.x1 - ink.x0));
  const bottom = Math.min(ink.height, Math.ceil(word.y1 - ink.y0));
  if (right - left < 2 || bottom - top < 2) {
    return 0;
  }
  let dark = 0;
  let total = 0;
  for (let row = top; row < bottom; row++) {
    for (let col = left; col < right; col++) {
      const index = (row * ink.width + col) * 4;
      const lum = 0.299 * ink.pixels[index] + 0.587 * ink.pixels[index + 1] + 0.114 * ink.pixels[index + 2];
      total++;
      if (lum < 145) {
        dark++;
      }
    }
  }
  return total > 0 ? dark / total : 0;
}

function looksLikeUnderlinedSubject(line: string): boolean {
  const trimmed = (line || '').trim();
  const letters = trimmed.replace(/[\s_\-.=]/g, '');
  if (letters.length < 2 || trimmed.length > 220) {
    return false;
  }
  if (/^[_.=-]{3,}$/.test(trimmed)) {
    return false;
  }
  if (isSalutationLine(trimmed) || isClosingLine(trimmed) || isAddressBlockStop(trimmed)) {
    return false;
  }
  return true;
}

function looksLikeBoldSubject(line: string): boolean {
  if (!looksLikeUnderlinedSubject(line)) {
    return false;
  }
  if (isBodyStart(line)) {
    return false;
  }
  const trimmed = (line || '').trim();
  if (trimmed.length > 100 && /^(i|we|please|thank)\b/i.test(trimmed)) {
    return false;
  }
  return true;
}

function looksLikeSubjectHeading(line: string): boolean {
  const trimmed = (line || '').trim();
  const letters = trimmed.replace(/[\s_\-.=]/g, '');
  if (letters.length < 2 || trimmed.length > 160) {
    return false;
  }
  if (/^[_.=-]{3,}$/.test(trimmed)) {
    return false;
  }
  if (isSalutationLine(trimmed) || isClosingLine(trimmed) || isBodyStart(trimmed) || isAddressBlockStop(trimmed)) {
    return false;
  }
  return true;
}

function stripSubjectLabel(line: string): string {
  return (line || '')
    .replace(/^(re|subject|ref)\s*[:.-]\s*/i, '')
    .replace(/\s+/g, ' ')
    .trim();
}

function subjectFromReBlock(page: IOcrPageResult): string {
  const fromWords = subjectFromReWords(page);
  if (fromWords) {
    return fromWords;
  }
  return subjectFromReText(page.text || '');
}

function subjectFromReWords(page: IOcrPageResult): string {
  const lines = groupWordsIntoLines(page.words || []);
  if (lines.length === 0) {
    return '';
  }
  const dear = findSalutationHit(page.words || []);
  const closing = findClosingHit(page.words || []);
  const hits: number[] = [];
  for (let index = 0; index < lines.length; index++) {
    if (matchReLabel(lines[index].text) === undefined) {
      continue;
    }
    const line = lines[index];
    if (dear && line.y1 <= dear.y1 - 2) {
      continue;
    }
    if (closing && line.y0 >= closing.y0 - 2) {
      continue;
    }
    hits.push(index);
  }
  if (hits.length === 0) {
    return '';
  }
  return collectReContinuation(lines, hits[0], dear ? dear.y0 : -1);
}

function subjectFromReText(text: string): string {
  const lines = (text || '').split(/\r?\n/).map((line) => line.trim());
  const compact: { text: string; sourceIndex: number }[] = [];
  for (let index = 0; index < lines.length; index++) {
    if (lines[index]) {
      compact.push({ text: lines[index], sourceIndex: index });
    }
  }
  let dearIndex = -1;
  let closingIndex = compact.length;
  for (let index = 0; index < compact.length; index++) {
    if (dearIndex < 0 && isSalutationLine(compact[index].text)) {
      dearIndex = index;
    }
    if (isClosingLine(compact[index].text)) {
      closingIndex = index;
      break;
    }
  }
  const hits: number[] = [];
  const start = dearIndex >= 0 ? dearIndex : 0;
  for (let index = start; index < closingIndex; index++) {
    if (matchReLabel(compact[index].text) !== undefined) {
      hits.push(index);
    }
  }
  if (hits.length === 0) {
    return '';
  }
  const mapped = compact.map((line) => ({
    text: line.text,
    y0: line.sourceIndex,
    y1: line.sourceIndex
  }));
  return collectReContinuation(mapped, hits[0], dearIndex);
}

function matchReLabel(line: string): { value: string } | undefined {
  const trimmed = (line || '').replace(/\s+/g, ' ').trim();
  if (!trimmed || /\b(?:our|your?|yr)\s+(?:re|ref)\b/i.test(trimmed)) {
    return undefined;
  }
  const match = trimmed.match(/^(?:re|subject)\s*[:;\uFF1A]\s*(.*)$/i);
  if (match) {
    return { value: stripSubjectLabel(match[1] || '') };
  }
  if (/^(?:re|subject)$/i.test(trimmed)) {
    return { value: '' };
  }
  return undefined;
}

function collectReContinuation(
  lines: { text: string; y0: number; y1: number }[],
  startIndex: number,
  dearY: number
): string {
  const start = lines[startIndex];
  if (!start) {
    return '';
  }
  const label = matchReLabel(start.text);
  const block: string[] = [];
  if (label && label.value) {
    block.push(cleanSubjectLine(label.value));
  }
  const reIsAboveDear = dearY >= 0 && start.y1 < dearY - 2;
  for (let index = startIndex + 1; index < lines.length && block.length < 12; index++) {
    const line = lines[index];
    const text = cleanSubjectLine(line.text);
    if (!text) {
      if (block.length > 0) {
        break;
      }
      continue;
    }
    if (reIsAboveDear && line.y0 >= dearY - 2) {
      break;
    }
    const previous = lines[index - 1];
    const lineHeight = Math.max(12, previous.y1 - previous.y0, 1);
    if (line.y0 - previous.y1 > lineHeight * 3.2) {
      break;
    }
    if (isSubjectContinueStop(text)) {
      break;
    }
    block.push(text);
  }
  return block.join(' ').replace(/\s+/g, ' ').trim();
}

function cleanSubjectLine(line: string): string {
  return stripSubjectLabel(line)
    .replace(/\b(date|tel|fax|email)\s*[:.].*$/i, '')
    .replace(/\s+/g, ' ')
    .trim();
}

function isSubjectContinueStop(line: string): boolean {
  return isSalutationLine(line) ||
    isClosingLine(line) ||
    isBodyStart(line) ||
    isAddressBlockStop(line) ||
    /^(?:attn|attention|cc)\b/i.test((line || '').trim());
}

function isBodyStart(line: string): boolean {
  const trimmed = (line || '').trim();
  return /^(i|we|please|thank|further|with reference)\b/i.test(trimmed);
}

function findSalutationIndex(text: string): { index: number } {
  const match = (text || '').match(/\bdear\s+(s[il1]rs?|madams?|mesdames)\b/i) ||
    (text || '').match(/\bdear\s+sir\s*[/\\]?\s*madam\b/i);
  return { index: match && match.index !== undefined ? match.index : -1 };
}

function isSalutationLine(line: string): boolean {
  const key = normalizeKey(line);
  return /\bdear\s+(s[il1]rs?|madams?|mesdames)\b/.test(key) ||
    /\bdear\s+sir\s*madam/.test(key);
}

function isAddressBlockStop(line: string): boolean {
  const key = normalizeKey(line);
  return key.indexOf('our ref') === 0 ||
    key.indexOf('your ref') === 0 ||
    key.indexOf('by fax') === 0 ||
    key.indexOf('by email') === 0 ||
    key.indexOf('by post') === 0 ||
    /^date\b/.test(key) ||
    /^tel\b/.test(key) ||
    /^fax\b/.test(key) ||
    /^email\b/.test(key);
}

function findSalutationHit(words: IOcrWord[]): { y0: number; y1: number } | undefined {
  const sorted = words.slice().sort((left, right) => {
    if (Math.abs(left.y0 - right.y0) > 10) {
      return left.y0 - right.y0;
    }
    return left.x0 - right.x0;
  });
  for (let index = 0; index < sorted.length; index++) {
    if (!isDearWord(sorted[index].text || '')) {
      continue;
    }
    const lineMid = (sorted[index].y0 + sorted[index].y1) / 2;
    const nearby = sorted.slice(index, index + 6).filter((word) =>
      Math.abs((word.y0 + word.y1) / 2 - lineMid) < 16
    );
    const phrase = nearby.map((word) => word.text || '').join(' ');
    if (!isSalutationLine(phrase) && !nearby.some((word) => isSirWord(word.text || ''))) {
      continue;
    }
    let y0 = sorted[index].y0;
    let y1 = sorted[index].y1;
    nearby.forEach((word) => {
      y0 = Math.min(y0, word.y0);
      y1 = Math.max(y1, word.y1);
    });
    return { y0, y1 };
  }
  return undefined;
}

function isDearWord(text: string): boolean {
  const key = normalizeKey(text);
  return key === 'dear' || key === 'deor' || key === 'dcar' || key === 'dear,';
}

function isSirWord(text: string): boolean {
  const key = normalizeKey(text);
  return key === 'sir' ||
    key === 'sirs' ||
    key === 'sit' ||
    key === 'slr' ||
    key === 'madam' ||
    key === 'madams' ||
    key === 'mesdames';
}

function findInkSignatures(image: HTMLImageElement, words: IOcrWord[]): ISignatureRegion[] {
  const canvas = document.createElement('canvas');
  canvas.width = image.width;
  canvas.height = image.height;
  const context = canvas.getContext('2d', { willReadFrequently: true });
  if (!context) {
    return [];
  }
  context.drawImage(image, 0, 0);

  const x0 = Math.floor(image.width * 0.40);
  const y0 = Math.floor(image.height * 0.40);
  const x1 = image.width;
  const y1 = Math.floor(image.height * 0.97);
  const width = x1 - x0;
  const height = y1 - y0;
  if (width < 16 || height < 16) {
    return [];
  }

  const pixels = context.getImageData(x0, y0, width, height).data;
  canvas.width = 0;
  canvas.height = 0;

  const ink: number[] = [];
  const printed: boolean[] = [];
  for (let row = 0; row < height; row++) {
    let dark = 0;
    for (let col = 0; col < width; col++) {
      const index = (row * width + col) * 4;
      const lum = 0.299 * pixels[index] + 0.587 * pixels[index + 1] + 0.114 * pixels[index + 2];
      if (lum < 135) {
        dark++;
      }
    }
    ink.push(dark / width);
    printed.push(rowHasPrintedText(y0 + row, words, x0, x1));
  }

  const smooth: number[] = ink.map((_, row) => {
    const from = Math.max(0, row - 1);
    const to = Math.min(ink.length - 1, row + 1);
    let sum = 0;
    for (let index = from; index <= to; index++) {
      sum += ink[index];
    }
    return sum / (to - from + 1);
  });

  const minHeight = Math.max(10, Math.round(image.height * 0.016));
  const maxHeight = Math.max(minHeight + 1, Math.round(image.height * 0.14));
  const footerStart = Math.floor(image.height * 0.975);
  const regions: ISignatureRegion[] = [];

  for (let start = 0; start < smooth.length; start++) {
    if (printed[start] || smooth[start] < 0.028) {
      continue;
    }
    let end = start;
    while (
      end < smooth.length &&
      end - start < maxHeight &&
      !printed[end] &&
      smooth[end] >= 0.018
    ) {
      end++;
    }
    const runHeight = end - start;
    if (runHeight >= minHeight) {
      const top = y0 + start;
      const bottom = y0 + end;
      if (bottom < footerStart) {
        regions.push({
          x0,
          y0: top,
          x1,
          y1: bottom
        });
      }
    }
    start = Math.max(start, end - 1);
  }

  return regions;
}

function rowHasPrintedText(y: number, words: IOcrWord[], x0: number, x1: number): boolean {
  for (let index = 0; index < words.length; index++) {
    const word = words[index];
    const text = (word.text || '').trim();
    if (text.length < 2 || !/[A-Za-z\u3400-\u9FFF]{2,}/.test(text)) {
      continue;
    }
    const overlapsY = word.y0 <= y && word.y1 >= y;
    const overlapsX = word.x0 < x1 && word.x1 > x0;
    if (overlapsY && overlapsX) {
      return true;
    }
  }
  return false;
}

function wordsBelowSignature(page: IOcrPageResult, region: ISignatureRegion): IOcrWord[] {
  const y0 = region.y1 - 2;
  const y1 = Math.min(page.height, region.y1 + Math.max(96, page.height * 0.1));
  const x0 = Math.max(0, region.x0 - page.width * 0.06);
  return (page.words || []).filter((word) => {
    const midX = (word.x0 + word.x1) / 2;
    const midY = (word.y0 + word.y1) / 2;
    return midX >= x0 && midY >= y0 && midY <= y1;
  });
}

function extractNameBetweenClosingAndTitle(page?: IOcrPageResult): string {
  if (!page) {
    return '';
  }
  const fromWords = nameBetweenClosingAndTitleFromWords(page);
  if (fromWords) {
    return fromWords;
  }
  return nameBetweenClosingAndTitleFromText(page.text || '');
}

function nameBetweenClosingAndTitleFromWords(page: IOcrPageResult): string {
  const closing = findClosingHit(page.words || []);
  if (!closing) {
    return '';
  }
  const below = (page.words || []).filter((word) => {
    const midX = (word.x0 + word.x1) / 2;
    const midY = (word.y0 + word.y1) / 2;
    return word.y0 >= closing.y1 - 10 &&
      midY <= closing.y1 + Math.max(260, page.height * 0.38) &&
      midX >= Math.max(0, closing.x0 - page.width * 0.4);
  });
  const lines = groupWordsIntoLines(below);
  return pickNameAboveTitle(lines.map((line) => ({ text: line.text, y0: line.y0 })));
}

function nameBetweenClosingAndTitleFromText(text: string): string {
  const lines = (text || '')
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter((line) => line.length > 0);
  for (let index = 0; index < lines.length; index++) {
    if (!isClosingLine(lines[index])) {
      continue;
    }
    const after = lines.slice(index + 1, index + 12).map((line, offset) => ({
      text: line,
      y0: offset
    }));
    return pickNameAboveTitle(after);
  }
  return '';
}

function pickNameAboveTitle(lines: { text: string; y0: number }[]): string {
  const found: { name: string; y0: number }[] = [];
  for (let index = 0; index < lines.length; index++) {
    const raw = (lines[index].text || '').replace(/\s+/g, ' ').trim();
    if (!raw || isClosingLine(raw) || isActingForLine(raw) || isIgnorableBelowClosing(raw)) {
      continue;
    }
    const cleaned = stripTrailingTitle(raw);
    if (!cleaned || isActingForLine(cleaned) || isIgnorableBelowClosing(cleaned)) {
      continue;
    }
    if (isJobTitleLine(raw) || isJobTitleLine(cleaned)) {
      const nameOnTitleLine = preferredNameFromLine(cleaned);
      if (nameOnTitleLine) {
        found.push({ name: nameOnTitleLine, y0: lines[index].y0 });
      }
      if (found.length > 0) {
        break;
      }
      continue;
    }
    const names = candidateNamesFromLine(cleaned);
    for (let nameIndex = 0; nameIndex < names.length; nameIndex++) {
      found.push({ name: names[nameIndex], y0: lines[index].y0 });
    }
  }

  if (found.length === 0) {
    return '';
  }
  found.sort((left, right) => left.y0 - right.y0);
  return found[found.length - 1].name;
}

function candidateNamesFromLine(line: string): string[] {
  const primary = preferredNameFromLine(line);
  if (primary) {
    return [primary];
  }
  const left = leftPersonName(line);
  return left ? [left] : [];
}

function preferredNameFromLine(line: string): string {
  return nameFromZoneLine(line) || loosePersonName(line);
}

function nameFromZoneLine(line: string): string {
  const inner = firstParenthesesContent(line);
  if (inner) {
    const fromInner = personNameFromLine(inner) || loosePersonName(inner);
    if (fromInner) {
      return fromInner;
    }
  }
  return personNameFromLine(line);
}

function loosePersonName(line: string): string {
  const text = (line || '')
    .replace(/^[(\uFF08]\s*/, '')
    .replace(/\s*[)\uFF09]$/, '')
    .replace(/\s+/g, ' ')
    .trim();
  if (!text || isJobTitleLine(text) || isNoiseLine(text) || isActingForLine(text)) {
    return '';
  }
  if (/工程師|總監|經理|主任|專員|顧問|秘書|署長|處長/.test(text)) {
    return '';
  }
  const cjk = text.match(/[\u3400-\u9FFF]/g);
  if (cjk && cjk.length >= 2 && cjk.length <= 4 && text.replace(/[\u3400-\u9FFF\s.·]/g, '').length === 0) {
    return text;
  }
  const withoutTitle = text.replace(/^(ir|engr|eng|dr|mr|mrs|ms|prof)\.?\s+/i, '');
  const tokens = withoutTitle.replace(/,/g, ' ').split(/\s+/).filter((token) => token.length > 0);
  if (tokens.length === 0 || tokens.length > 5) {
    return '';
  }
  const ok = tokens.every((token) =>
    /^[A-Za-z]\.?$/.test(token) ||
    /^[A-Za-z](?:\.[A-Za-z])+\.?$/.test(token) ||
    /^[A-Z][a-z]+(?:-[A-Z][a-z]+)?$/.test(token) ||
    /^[A-Z]{2,12}$/.test(token)
  );
  return ok ? text : '';
}

function leftPersonName(line: string): string {
  const withoutTitle = (line || '').replace(/^(ir|engr|eng|dr|mr|mrs|ms|prof)\.?\s+/i, '');
  const tokens = withoutTitle.replace(/,/g, ' ').split(/\s+/).filter((token) => token.length > 0);
  if (tokens.length < 4) {
    return '';
  }
  const firstThree = preferredNameFromLine(tokens.slice(0, 3).join(' '));
  if (firstThree) {
    return firstThree;
  }
  return preferredNameFromLine(tokens.slice(0, 2).join(' '));
}

function isActingForLine(line: string): boolean {
  const key = normalizeKey(line);
  return /^for\b/.test(key) || /\bon behalf\b/.test(key);
}

function isJobTitleLine(line: string): boolean {
  const key = normalizeKey(line);
  if (!key) {
    return false;
  }
  if (/工程師|總監|經理|主任|專員|顧問|秘書|署長|處長/.test(line)) {
    return true;
  }
  const titles = [
    'director', 'manager', 'engineer', 'associate', 'consultant', 'officer',
    'secretary', 'architect', 'planner', 'surveyor', 'partner', 'chief',
    'assistant', 'principal', 'coordinator', 'specialist', 'supervisor',
    'technician', 'inspector', 'executive', 'president', 'leader', 'head'
  ];
  const tokens = key.split(' ');
  for (let index = 0; index < titles.length; index++) {
    if (tokens.indexOf(titles[index]) >= 0) {
      return true;
    }
  }
  return false;
}

function stripTrailingTitle(line: string): string {
  let text = (line || '').replace(/\s+/g, ' ').trim();
  text = text.split(/,\s+(?=(?:ir|engr|eng|dr|mr|mrs|ms|prof)?\.?\s*(?:chief|director|manager|engineer|associate|consultant|officer|secretary|architect)\b)/i)[0] || text;
  text = text.replace(/\s+(?:(?:ir|engr|eng|dr|mr|mrs|ms|prof)\.?\s+)?(?:chief\s+)?(?:director|manager|engineer|associate|consultant|officer|secretary|architect|planner|surveyor|partner|coordinator|specialist|supervisor|technician|inspector|executive|president|leader)\b.*$/i, '');
  text = text.replace(/\s*(工程師|總監|經理|主任|專員|顧問|秘書|署長|處長)\s*$/, '');
  return text.replace(/\s+/g, ' ').trim();
}

function isIgnorableBelowClosing(line: string): boolean {
  const key = normalizeKey(line);
  return !key ||
    isIgnorableParen(key) ||
    key === 'signed' ||
    /^for and on behalf/.test(key) ||
    /^[-_.=]+$/.test(line);
}

function isClosingLine(line: string): boolean {
  const key = normalizeKey(line);
  return /(^|\s)yours\s+sincere/.test(' ' + key) ||
    /(^|\s)yours\s+faithful/.test(' ' + key) ||
    /(^|\s)yours\s+truly/.test(' ' + key);
}

function firstParenthesesContent(text: string): string {
  const matches = (text || '').match(/[(\uFF08]\s*([^)\uFF09]{1,80}?)\s*[)\uFF09]/g) || [];
  for (let index = 0; index < matches.length; index++) {
    const innerMatch = matches[index].match(/[(\uFF08]\s*([^)\uFF09]{1,80}?)\s*[)\uFF09]/);
    const inner = innerMatch ? innerMatch[1].replace(/\s+/g, ' ').trim() : '';
    if (inner && !isIgnorableParen(inner)) {
      return inner;
    }
  }
  return '';
}

function isIgnorableParen(value: string): boolean {
  const key = normalizeKey(value);
  return key === 'signed' ||
    key === 'signature' ||
    key === 'sgd' ||
    key === 'chop' ||
    key === 'seal';
}

function findClosingHit(words: IOcrWord[]): { x0: number; y0: number; y1: number } | undefined {
  const sorted = words.slice().sort((left, right) => {
    if (Math.abs(left.y0 - right.y0) > 10) {
      return left.y0 - right.y0;
    }
    return left.x0 - right.x0;
  });
  for (let index = 0; index < sorted.length; index++) {
    const window = sorted.slice(index, index + 4);
    const lineMid = (sorted[index].y0 + sorted[index].y1) / 2;
    const sameLine = window.filter((word) => Math.abs((word.y0 + word.y1) / 2 - lineMid) < 12);
    const phrase = sameLine.map((word) => word.text || '').join(' ');
    if (!isClosingLine(phrase)) {
      continue;
    }
    let x0 = sameLine[0].x0;
    let y0 = sameLine[0].y0;
    let y1 = sameLine[0].y1;
    sameLine.forEach((word) => {
      x0 = Math.min(x0, word.x0);
      y0 = Math.min(y0, word.y0);
      y1 = Math.max(y1, word.y1);
    });
    return { x0, y0, y1 };
  }
  return undefined;
}

function firstPersonName(text: string): string {
  const lines = (text || '')
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter((line) => line.length > 0);
  for (let index = 0; index < lines.length; index++) {
    const name = personNameFromLine(lines[index]);
    if (name) {
      return name;
    }
  }
  return '';
}

function personNameFromLine(line: string): string {
  const cleaned = (line || '')
    .replace(/^[\s(]+signed[\s)]+$/i, '')
    .replace(/^[-_.=]+$/, '')
    .trim();
  if (!cleaned || isJobTitleLine(cleaned) || isNoiseLine(cleaned) || !looksLikePersonName(cleaned)) {
    return '';
  }
  return cleaned.replace(/\s+/g, ' ').trim();
}

function isNoiseLine(line: string): boolean {
  const key = normalizeKey(line);
  if (!key || /@|\d{3,}/.test(line) || /^https?:/i.test(line)) {
    return true;
  }
  const noise = [
    'tel', 'fax', 'email', 'www', 'http', 'page', 'enclosure', 'attachment',
    'limited', 'ltd', 'company', 'department', 'division', 'office',
    'director', 'manager', 'engineer', 'associate', 'consultant', 'officer',
    'secretary', 'architect', 'planner', 'surveyor', 'partner', 'chief',
    'yours', 'faithfully', 'sincerely', 'behalf'
  ];
  const tokens = key.split(' ');
  for (let index = 0; index < noise.length; index++) {
    if (tokens.indexOf(noise[index]) >= 0 && !hasHonorific(line)) {
      return true;
    }
  }
  return false;
}

function hasHonorific(line: string): boolean {
  return /^(ir|engr|eng|dr|mr|mrs|ms|prof)\b/i.test(line.trim());
}

function looksLikePersonName(line: string): boolean {
  const trimmed = line.trim();
  if (/工程師|總監|經理|主任|專員|顧問|秘書|署長|處長/.test(trimmed)) {
    return false;
  }
  const cjk = trimmed.match(/[\u3400-\u9FFF]/g);
  if (cjk && cjk.length >= 2 && cjk.length <= 4) {
    return trimmed.replace(/[\u3400-\u9FFF\s.·]/g, '').length === 0;
  }

  const withoutTitle = trimmed.replace(/^(ir|engr|eng|dr|mr|mrs|ms|prof)\.?\s+/i, '');
  const tokens = withoutTitle.replace(/,/g, ' ').split(/\s+/).filter((token) => token.length > 0);
  if (tokens.length < 2 || tokens.length > 5) {
    return false;
  }

  return tokens.every((token) => {
    if (/^[A-Za-z]\.?$/.test(token) || /^[A-Za-z](?:\.[A-Za-z])+\.?$/.test(token)) {
      return true;
    }
    if (/^[A-Z][a-z]+(?:-[A-Z][a-z]+)?$/.test(token) || /^[A-Z]{2,12}$/.test(token)) {
      return true;
    }
    return false;
  });
}

function normalizeKey(text: string): string {
  return text.toLowerCase().replace(/[^a-z0-9]+/g, ' ').replace(/\s+/g, ' ').trim();
}

function loadImage(url: string): Promise<HTMLImageElement> {
  return new Promise((resolve, reject) => {
    const image = new Image();
    image.onload = () => resolve(image);
    image.onerror = () => reject(new Error('Unable to load page image.'));
    image.src = url;
  });
}
