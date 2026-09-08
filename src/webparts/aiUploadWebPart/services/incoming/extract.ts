// Incoming OCR locators only. Outgoing lives in services/outgoing/.
import { formatIssueDate, parseIssueDate } from '../../constants/issueDate';
import { projectNumberFromRef } from '../../constants/projectNumber';
import { extractYourRefNo } from '../fieldExtractor';
import { IOcrPageResult, IOcrWord } from '../IPdfOcr';
import { formatOcrTextWithStyles, joinOcrWords } from '../ocrSelection';

const DATE_LABELS: string[] = ['Date', '日期'];
const MEMO_FROM_LABELS: string[] = ['From', '發件人', '寄件人'];
const MEMO_TO_LABELS: string[] = ['To', '收件人', '致'];

interface IIncomingClosingHit {
  x0: number;
  x1: number;
  y0: number;
  y1: number;
}

interface ILine {
  text: string;
  x0: number;
  x1: number;
  y0: number;
  y1: number;
  words?: IOcrWord[];
}

export type IncomingLetterType = 'government' | 'consultant' | 'chinese' | 'memo' | 'email' | 'unknown';

export interface IIncomingClassification {
  letterType: IncomingLetterType;
  language: 'en' | 'zh' | 'bilingual';
  hasOurRef: boolean;
  hasYourRef: boolean;
}

export function classifyIncomingLetter(page?: IOcrPageResult): IIncomingClassification {
  const text = pageText(page);
  const hasChinese = /[\u4e00-\u9fff]/.test(text);
  const hasEnglish = /[A-Za-z]{4,}/.test(text);
  const language: IIncomingClassification['language'] = hasChinese && hasEnglish
    ? 'bilingual'
    : hasChinese
      ? 'zh'
      : 'en';
  const hasOurRef = /\bour\s+r+e+f/i.test(text) || /本[處处署局]檔號|本函編號/.test(text);
  const hasYourRef = /\byour?\s+r+e+f/i.test(text) || /貴[處处署局]檔號|來函編號/.test(text);
  const letterType = detectLetterType(text, hasOurRef);
  return { letterType, language, hasOurRef, hasYourRef };
}

export function incomingProjectNumber(pages: IOcrPageResult[]): string {
  const yourRef = extractYourRefNo(pages || []);
  return projectNumberFromRef(yourRef) || eightDigitRun(yourRef);
}

export function extractIncomingOrganization(page?: IOcrPageResult): string {
  if (!page) {
    return '';
  }
  const fromAddress = organizationAboveAddressFloor(page);
  if (fromAddress) {
    return fromAddress;
  }
  const fromWords = organizationFromLetterhead(page);
  if (fromWords) {
    return fromWords;
  }
  return organizationFromLetterheadText(page.text || '');
}

export function extractIncomingIssueDate(pages: IOcrPageResult[]): string {
  const list = pages || [];
  for (let index = 0; index < list.length; index++) {
    const page = list[index];
    const fromLabel = dateFromLabeledLines(page);
    if (fromLabel) {
      return fromLabel;
    }
    const fromHeader = dateFromHeaderArea(page);
    if (fromHeader) {
      return fromHeader;
    }
  }
  return '';
}

export function extractIncomingMemoSender(page?: IOcrPageResult): string {
  return labeledValue(page, MEMO_FROM_LABELS);
}

export function extractIncomingMemoReceiver(page?: IOcrPageResult): string {
  return labeledValue(page, MEMO_TO_LABELS);
}

export function extractIncomingSender(pages: IOcrPageResult[]): string {
  const list = pages || [];
  for (let index = list.length - 1; index >= 0; index--) {
    if (!pageHasIncomingYoursClosing(list[index])) {
      continue;
    }
    const name = signatureParenthesesOnPage(list[index]);
    if (name) {
      return name;
    }
  }
  for (let index = list.length - 1; index >= 0; index--) {
    const name = signatureParenthesesOnPage(list[index]);
    if (name) {
      return name;
    }
  }
  return '';
}

function signatureParenthesesOnPage(page?: IOcrPageResult): string {
  if (!page) {
    return '';
  }
  const fromClosing = parentheticalNameBelowClosing(page);
  if (fromClosing) {
    return fromClosing;
  }
  return parentheticalNameAfter署名Label(page);
}

function pageHasIncomingYoursClosing(page?: IOcrPageResult): boolean {
  if (!page) {
    return false;
  }
  const lines = groupWordsIntoLines(page.words || []);
  for (let index = 0; index < lines.length; index++) {
    const combined = lines[index + 1] ? (lines[index].text + ' ' + lines[index + 1].text) : lines[index].text;
    if (isIncomingYoursClosingPhrase(lines[index].text) || isIncomingYoursClosingPhrase(combined)) {
      return true;
    }
  }
  return isIncomingYoursClosingPhrase(stripIncomingMarkup(page.text || ''));
}

export function incomingSenderName(value: string): string {
  const text = (value || '').replace(/\s+/g, ' ').trim();
  if (!text) {
    return '';
  }
  const inner = signatureParenthesesInner(text);
  if (inner) {
    return inner;
  }
  if (isIncomingJobOrDeptLine(text) || isIncomingIgnorableParen(text) || isIncomingAttnParenInner(text)) {
    return '';
  }
  if (/[[(\uFF08【]/.test(text) && /[)\uFF09\]】]/.test(text)) {
    return '';
  }
  return cleanIncomingSenderName(text);
}

export function incomingSignatureParenName(value: string): string {
  return signatureParenthesesInner(value);
}

export function extractIncomingReceiver(pages?: IOcrPageResult[] | IOcrPageResult): string {
  const list = !pages ? [] : Array.isArray(pages) ? pages : [pages];
  for (let index = list.length - 1; index >= 0; index--) {
    if (!pageHasIncomingYoursClosing(list[index])) {
      continue;
    }
    const value = receiverDirectlyBelowYours(list[index]);
    if (value) {
      return value;
    }
  }
  for (let index = list.length - 1; index >= 0; index--) {
    const value = receiverDirectlyBelowYours(list[index]);
    if (value) {
      return value;
    }
  }
  const firstPage = list[0];
  if (!firstPage) {
    return '';
  }
  const fromAttn = extractIncomingAttn(firstPage);
  if (fromAttn) {
    return fromAttn;
  }
  return firstAddressLineAboveDear(firstPage);
}

function receiverDirectlyBelowYours(page?: IOcrPageResult): string {
  if (!page) {
    return '';
  }
  const fromWords = receiverDirectlyBelowYoursFromWords(page);
  if (fromWords) {
    return fromWords;
  }
  return receiverDirectlyBelowYoursFromText(page.text || '');
}

function receiverDirectlyBelowYoursFromWords(page: IOcrPageResult): string {
  const closing = findIncomingYoursClosingHit(page);
  if (!closing) {
    return '';
  }
  const lineHeight = Math.max(closing.y1 - closing.y0, 12);
  const maxY = closing.y1 + Math.max(lineHeight * 3.4, (page.height || 0) * 0.10);
  const colLeft = closing.x0 - Math.max(12, (page.width || 0) * 0.015);
  const colRight = Math.max(
    closing.x1 + 12,
    closing.x0 + Math.max(closing.x1 - closing.x0, (page.width || 0) * 0.40)
  );
  const under = (page.words || []).filter((word) => {
    const midX = (word.x0 + word.x1) / 2;
    const midY = (word.y0 + word.y1) / 2;
    return midY >= closing.y1 - 2 &&
      word.y0 <= maxY &&
      midX >= colLeft &&
      midX <= colRight;
  });
  const lines = groupWordsIntoLines(under);
  for (let index = 0; index < lines.length; index++) {
    const text = stripIncomingMarkup(lines[index].text);
    if (isIncomingCcLine(text)) {
      break;
    }
    if (isIncomingReceiverSkipLine(text)) {
      continue;
    }
    return text.replace(/\s+/g, ' ').trim();
  }
  return '';
}

function receiverDirectlyBelowYoursFromText(text: string): string {
  const lines = stripIncomingMarkup(text || '').split(/\r?\n/).map((line) => line.replace(/\s+/g, ' ').trim()).filter((line) => line.length > 0);
  let closingIndex = -1;
  for (let index = 0; index < lines.length; index++) {
    const combined = lines[index + 1] ? (lines[index] + ' ' + lines[index + 1]) : lines[index];
    if (isIncomingYoursClosingPhrase(lines[index]) || isIncomingYoursClosingPhrase(combined) || isIncomingReceiverClosingPhrase(lines[index])) {
      closingIndex = index;
    }
  }
  if (closingIndex < 0) {
    return '';
  }
  const remainder = textAfterIncomingClosing(lines[closingIndex]);
  const after = remainder ? [remainder] : [];
  const limit = Math.min(lines.length, closingIndex + 5);
  for (let index = closingIndex + 1; index < limit; index++) {
    after.push(lines[index]);
  }
  for (let index = 0; index < after.length; index++) {
    const line = after[index];
    if (isIncomingCcLine(line)) {
      break;
    }
    if (isIncomingReceiverSkipLine(line)) {
      continue;
    }
    return line.replace(/\s+/g, ' ').trim();
  }
  return '';
}

function isIncomingReceiverClosingPhrase(line: string): boolean {
  const plain = stripIncomingMarkup(line);
  return /^此致/.test(plain) && plain.length <= 12;
}

function isIncomingCcLine(text: string): boolean {
  const key = stripIncomingMarkup(text).toLowerCase().replace(/\s+/g, ' ').trim();
  if (!key) {
    return false;
  }
  return /^(c\.?\s*c\.?|cc|copy\s+to|copied\s+to|副本|抄送|副本送|副本抄送)\b/.test(key) ||
    /^(c\.?\s*c\.?|cc)\s*[:：]/.test(key);
}

function isIncomingReceiverSkipLine(text: string): boolean {
  const plain = stripIncomingMarkup(text);
  if (!plain) {
    return true;
  }
  if (isIncomingYoursClosingPhrase(plain) || isIncomingReceiverClosingPhrase(plain) || isIncomingOtherClosingPhrase(plain)) {
    return true;
  }
  if (isIncomingCcLine(plain) || isIncomingIgnorableParen(plain)) {
    return true;
  }
  if (isIncomingSignatureParenLine(plain) || isIncomingJobOrDeptLine(plain)) {
    return true;
  }
  if (/^for and on behalf\b/i.test(plain) || /^for (?:the )?director of\b/i.test(plain)) {
    return true;
  }
  if (/^(encl|enc|encls|enclosure|enclosures|附件|隨函)\b/i.test(plain)) {
    return true;
  }
  return false;
}

function isIncomingSignatureParenLine(text: string): boolean {
  const plain = stripIncomingMarkup(text);
  if (!plain) {
    return false;
  }
  if (/^[(\uFF08]\s*(signed|signature|sgd)\s*[)\uFF09]$/i.test(plain)) {
    return true;
  }
  const inner = signatureParenthesesInner(plain);
  if (!inner) {
    return false;
  }
  const without = plain.replace(/[[(\uFF08【][^)\uFF09\]】]*[)\uFF09\]】]/g, '').replace(/\s+/g, ' ').trim();
  return without.length === 0 || isIncomingJobOrDeptLine(without);
}

function extractIncomingAttn(page: IOcrPageResult): string {
  const dearY = incomingSalutationY(page);
  const fromWords = attnFromOcrWords(page, dearY);
  if (fromWords) {
    return fromWords;
  }
  const fromLines = attnFromGroupedLines(page, dearY);
  if (fromLines) {
    return fromLines;
  }
  return attnFromPageText(page.text || '');
}

function incomingSalutationY(page: IOcrPageResult): number {
  const lines = groupWordsIntoLines(page.words || []);
  const dearIndex = findIncomingSalutationIndex(lines.map((line) => line.text));
  return dearIndex >= 0 ? lines[dearIndex].y0 : -1;
}

function attnFromOcrWords(page: IOcrPageResult, dearY: number): string {
  const words = (page.words || []).slice().sort((left, right) => {
    if (Math.abs(left.y0 - right.y0) > 8) {
      return left.y0 - right.y0;
    }
    return left.x0 - right.x0;
  });
  for (let index = 0; index < words.length; index++) {
    if (!isIncomingAttnWord(words[index].text || '')) {
      continue;
    }
    if (dearY >= 0 && words[index].y0 >= dearY - 2) {
      continue;
    }
    if ((words[index].x0 + words[index].x1) / 2 > incomingAddressRight(page)) {
      continue;
    }
    const value = valueAfterIncomingAttn(words, index, incomingAddressRight(page));
    if (value) {
      return value;
    }
  }
  return '';
}

function isIncomingAttnWord(text: string): boolean {
  const key = (text || '').toLowerCase().replace(/[^a-z0-9]+/g, '');
  return key === 'attn' ||
    key === 'atin' ||
    key === 'attm' ||
    key === 'atln' ||
    key === 'attention';
}

function valueAfterIncomingAttn(words: IOcrWord[], index: number, addressRight: number): string {
  let start = index + 1;
  while (words[start]) {
    const raw = (words[start].text || '').trim();
    const key = (raw || '').toLowerCase().replace(/[^a-z0-9]+/g, '');
    if (/^[:.：-]+$/.test(raw) || key === 'of' || key === 'to') {
      start++;
      continue;
    }
    break;
  }
  const label = words[index];
  const lineMid = (label.y0 + label.y1) / 2;
  const lineHeight = Math.max(label.y1 - label.y0, 1);
  const pageWidth = addressRight < Number.POSITIVE_INFINITY ? addressRight / 0.55 : 800;
  const gapLimit = Math.max(36, pageWidth * 0.06);
  const sameLine: IOcrWord[] = [];
  for (let cursor = start; cursor < words.length; cursor++) {
    const word = words[cursor];
    const wordMid = (word.y0 + word.y1) / 2;
    if (Math.abs(wordMid - lineMid) > lineHeight * 0.85) {
      break;
    }
    if (word.x0 < label.x1 - 6) {
      continue;
    }
    if ((word.x0 + word.x1) / 2 > addressRight) {
      break;
    }
    if (sameLine.length > 0 && word.x0 - sameLine[sameLine.length - 1].x1 > gapLimit) {
      break;
    }
    if (isIncomingAttnValueStop(word.text || '')) {
      break;
    }
    sameLine.push(word);
  }
  const same = joinOcrWords(sameLine).replace(/^[:.：\s-]+/, '').trim();
  if (same) {
    return same;
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
    if ((word.x0 + word.x1) / 2 > addressRight) {
      continue;
    }
    if (isIncomingAttnValueStop(word.text || '') || isIncomingSalutation(word.text || '')) {
      break;
    }
    below.push(word);
  }
  return joinOcrWords(below).replace(/^[:.：\s-]+/, '').trim();
}

function isIncomingAttnValueStop(text: string): boolean {
  const key = (text || '').toLowerCase().replace(/[^a-z0-9\u4e00-\u9fff]+/g, '');
  return key === 'date' ||
    key === 'tel' ||
    key === 'fax' ||
    key === 'email' ||
    key === 'dear' ||
    isIncomingAttnWord(text) ||
    /^(by|via)(post|hand|fax|email)?$/.test(key);
}

function attnFromGroupedLines(page: IOcrPageResult, dearY: number): string {
  const lines = leftAddressLines(page);
  for (let index = 0; index < lines.length; index++) {
    if (dearY >= 0 && lines[index].y0 >= dearY - 2) {
      break;
    }
    const match = matchAttnLine(lines[index].text);
    if (!match) {
      continue;
    }
    if (match.value) {
      return stripRightColumnNoise(match.value);
    }
    const next = lines[index + 1];
    if (next && next.text && !matchAttnLine(next.text) && !isIncomingSalutation(next.text)) {
      return stripRightColumnNoise(next.text.replace(/\s+/g, ' ').trim());
    }
  }
  return '';
}

function attnFromPageText(text: string): string {
  const lines = (text || '').split(/\r?\n/).map((line) => line.replace(/\s+/g, ' ').trim()).filter((line) => line.length > 0);
  const dearIndex = findIncomingSalutationIndex(lines);
  const limit = dearIndex >= 0 ? dearIndex : lines.length;
  for (let index = 0; index < limit; index++) {
    const match = matchAttnLine(lines[index]);
    if (!match) {
      continue;
    }
    if (match.value) {
      return stripRightColumnNoise(match.value);
    }
    const next = lines[index + 1];
    if (next && !matchAttnLine(next) && !isIncomingSalutation(next)) {
      return stripRightColumnNoise(next);
    }
  }
  return '';
}

function matchAttnLine(line: string): { value: string } | undefined {
  const trimmed = (line || '').replace(/\s+/g, ' ').trim();
  if (!trimmed) {
    return undefined;
  }
  const labeled = trimmed.match(/\b(?:for\s+the\s+)?(?:attn|atin|attm|atln|attention)(?:\s+of)?\s*[:：.]\s*(.+)$/i);
  if (labeled) {
    return { value: stripTrailingAttnNoise((labeled[1] || '').replace(/\s+/g, ' ').trim()) };
  }
  if (/\b(?:attn|atin|attm|atln|attention)$/i.test(trimmed)) {
    return { value: '' };
  }
  return undefined;
}

function stripTrailingAttnNoise(value: string): string {
  return (value || '')
    .replace(/\s+\b(?:tel|fax|email|date|our\s+ref|your\s+ref)\b.*$/i, '')
    .trim();
}

function firstAddressLineAboveDear(page: IOcrPageResult): string {
  const cluster = addresseeLinesAboveDear(page);
  const picked = dropLetterheadFromAddress(cluster, page);
  if (picked) {
    return picked;
  }
  return firstAddressLineAboveDearFromText(page.text || '');
}

function organizationAboveAddressFloor(page: IOcrPageResult): string {
  const cluster = addresseeLinesAboveDear(page);
  const fromCluster = organizationAboveFloorInLines(cluster);
  if (fromCluster) {
    return fromCluster;
  }
  return organizationAboveFloorInLines(addresseeLinesAboveDearFromText(page.text || ''));
}

// Address sits on the left. A full OCR row can also pick up Our Ref / Date on the right, so
// Receiver and Organization only use the left-hand cluster, never the whole visual line.
function addresseeLinesAboveDear(page: IOcrPageResult): string[] {
  const lines = leftAddressLines(page);
  if (lines.length === 0) {
    return addresseeLinesAboveDearFromText(page.text || '');
  }
  const dearIndex = findIncomingSalutationIndex(lines.map((line) => line.text));
  if (dearIndex < 0) {
    return addresseeLinesAboveDearFromText(page.text || '');
  }
  const cutoffY = leftAddressTopY(lines);
  const dearHeight = Math.max(lines[dearIndex].y1 - lines[dearIndex].y0, 10);
  const cluster: string[] = [];
  let lastAcceptedY0 = lines[dearIndex].y0;
  for (let index = dearIndex - 1; index >= 0; index--) {
    const line = lines[index];
    const text = stripRightColumnNoise((line.text || '').replace(/\s+/g, ' ').trim());
    if (!text || isIncomingSalutation(text) || matchAttnLine(text) || isIncomingDeliveryLine(text)) {
      continue;
    }
    if (isIncomingMetaHeader(text) || isHeaderStopLine(text)) {
      break;
    }
    if (cutoffY > 0 && line.y1 <= cutoffY) {
      break;
    }
    const gap = lastAcceptedY0 - line.y1;
    const lineHeight = Math.max(line.y1 - line.y0, 10);
    if (cluster.length > 0 && gap > lineHeight * 1.8) {
      break;
    }
    if (cluster.length > 0 && lineHeight > dearHeight * 1.7) {
      break;
    }
    cluster.unshift(text);
    lastAcceptedY0 = line.y0;
    if (cluster.length >= 6) {
      break;
    }
  }
  if (cluster.length > 0) {
    return cluster;
  }
  return addresseeLinesAboveDearFromText(page.text || '');
}

function leftAddressLines(page: IOcrPageResult): ILine[] {
  const right = incomingAddressRight(page);
  const leftWords = (page.words || []).filter((word) => ((word.x0 + word.x1) / 2) <= right);
  return groupWordsIntoLines(leftWords).map((line) => leftClusterOfLine(line, page.width || 0));
}

function incomingAddressRight(page: IOcrPageResult): number {
  const width = page.width || 0;
  if (width <= 0) {
    return Number.POSITIVE_INFINITY;
  }
  return width * 0.55;
}

function leftClusterOfLine(line: ILine, pageWidth: number): ILine {
  const words = (line.words || []).slice().sort((left, right) => left.x0 - right.x0);
  if (words.length === 0) {
    return line;
  }
  const gapLimit = Math.max(36, pageWidth * 0.06);
  const left: IOcrWord[] = [words[0]];
  for (let index = 1; index < words.length; index++) {
    const previous = left[left.length - 1];
    if (words[index].x0 - previous.x1 > gapLimit) {
      break;
    }
    left.push(words[index]);
  }
  let x0 = left[0].x0;
  let x1 = left[0].x1;
  let y0 = left[0].y0;
  let y1 = left[0].y1;
  left.forEach((word) => {
    x0 = Math.min(x0, word.x0);
    x1 = Math.max(x1, word.x1);
    y0 = Math.min(y0, word.y0);
    y1 = Math.max(y1, word.y1);
  });
  return {
    text: joinOcrWords(left).replace(/\s+/g, ' ').trim(),
    x0,
    x1,
    y0,
    y1,
    words: left
  };
}

function leftAddressTopY(lines: ILine[]): number {
  for (let index = 0; index < lines.length; index++) {
    if (isHeaderStopLine(lines[index].text) || isIncomingMetaHeader(lines[index].text)) {
      return lines[index].y0 - 2;
    }
  }
  return 0;
}

function stripRightColumnNoise(text: string): string {
  return (text || '')
    .replace(/\s+\b(?:our\s+r+e+f|your?\s+r+e+f|date|tel|fax|email)\b.*$/i, '')
    .replace(/\s+(?:本[處处署局]檔號|貴[處处署局]檔號|日期)\s*[:：].*$/g, '')
    .trim();
}

function addresseeLinesAboveDearFromText(text: string): string[] {
  const lines = (text || '').split(/\r?\n/).map((line) => line.replace(/\s+/g, ' ').trim()).filter((line) => line.length > 0);
  const dearIndex = findIncomingSalutationIndex(lines);
  if (dearIndex < 0) {
    return [];
  }
  const cluster: string[] = [];
  for (let index = dearIndex - 1; index >= 0; index--) {
    const line = lines[index];
    if (!line || isIncomingSalutation(line) || matchAttnLine(line) || isIncomingDeliveryLine(line)) {
      continue;
    }
    if (isIncomingMetaHeader(line) || isHeaderStopLine(line)) {
      break;
    }
    if (/\b(?:tel|fax|email|www|http)\b/i.test(line)) {
      break;
    }
    cluster.unshift(stripRightColumnNoise(line));
    if (cluster.length >= 6) {
      break;
    }
  }
  return cluster.filter((line) => line.length > 0);
}

function organizationAboveFloorInLines(lines: string[]): string {
  let anchorIndex = lastIncomingAnchorIndex(lines, isIncomingFloorLine);
  if (anchorIndex < 0) {
    anchorIndex = lastIncomingAnchorIndex(lines, isIncomingRoadLine);
  }
  if (anchorIndex < 0) {
    return '';
  }
  if (isIncomingFloorLine(lines[anchorIndex])) {
    const sameLine = textBeforeFloor(lines[anchorIndex]);
    if (sameLine) {
      return sameLine;
    }
  }
  for (let previous = anchorIndex - 1; previous >= 0; previous--) {
    const candidate = (lines[previous] || '').replace(/\s+/g, ' ').trim();
    if (!candidate || isIncomingAddressAnchorLine(candidate) || matchAttnLine(candidate) || isIncomingDeliveryLine(candidate)) {
      continue;
    }
    return cleanOrganization(candidate);
  }
  return '';
}

function lastIncomingAnchorIndex(lines: string[], match: (line: string) => boolean): number {
  let found = -1;
  for (let index = 0; index < lines.length; index++) {
    if (match(lines[index])) {
      found = index;
    }
  }
  return found;
}

function isIncomingAddressAnchorLine(line: string): boolean {
  return isIncomingFloorLine(line) || isIncomingRoadLine(line);
}

function isIncomingFloorLine(line: string): boolean {
  return /\b(?:\d{1,3}|g|ug|lg|m)\s*[/\\\uFF0F]\s*f\b/i.test((line || '').replace(/\s+/g, ' ').trim());
}

function isIncomingRoadLine(line: string): boolean {
  return /\b\S+\s+(?:Road|Rd\.?)\b/i.test((line || '').replace(/\s+/g, ' ').trim());
}

function textBeforeFloor(line: string): string {
  const trimmed = (line || '').replace(/\s+/g, ' ').trim();
  const match = trimmed.match(/^(.*?)\b(?:\d{1,3}|g|ug|lg|m)\s*[/\\\uFF0F]\s*f\b/i);
  const before = cleanOrganization((match && match[1]) || '');
  return before.length >= 4 ? before : '';
}

function firstAddressLineAboveDearFromText(text: string): string {
  return addresseeLinesAboveDearFromText(text)[0] || '';
}

function dropLetterheadFromAddress(cluster: string[], page: IOcrPageResult): string {
  const org = (organizationFromLetterhead(page) || '').replace(/\s+/g, ' ').trim().toLowerCase();
  const lines = cluster.slice();
  while (lines.length > 0 && org) {
    const top = (lines[0] || '').replace(/\s+/g, ' ').trim().toLowerCase();
    if (top === org || (top.length > 8 && org.indexOf(top) >= 0) || (org.length > 8 && top.indexOf(org) >= 0)) {
      lines.shift();
      continue;
    }
    break;
  }
  return lines[0] || '';
}

function findIncomingSalutationIndex(lines: string[]): number {
  for (let index = 0; index < lines.length; index++) {
    if (isIncomingSalutation(lines[index])) {
      return index;
    }
  }
  return -1;
}

function isIncomingSalutation(line: string): boolean {
  const trimmed = (line || '').replace(/\s+/g, ' ').trim();
  if (!trimmed) {
    return false;
  }
  if (/敬啟者|敬启者|鈞鑒|台鑒/.test(trimmed)) {
    return true;
  }
  return /\bdear\s+(s[il1]rs?|madams?|mesdames|sir\s*[/\\]?\s*madam)\b/i.test(trimmed) ||
    (/^dear\b/i.test(trimmed) && trimmed.length <= 90);
}

function isIncomingDeliveryLine(line: string): boolean {
  const key = (line || '').toLowerCase().replace(/[^a-z0-9\u4e00-\u9fff]+/g, ' ').trim();
  return /^(by|via)\s+(post|hand|fax|email|e mail|courier)/.test(key) ||
    /^registered\s+(post|mail)/.test(key);
}

function isIncomingMetaHeader(line: string): boolean {
  const key = (line || '').toLowerCase().replace(/[^a-z0-9\u4e00-\u9fff]+/g, ' ').trim();
  return /^our\s+r/.test(key) ||
    /^your?\s+r/.test(key) ||
    /^date\b/.test(key) ||
    /^tel\b/.test(key) ||
    /^fax\b/.test(key) ||
    /^email\b/.test(key) ||
    /本[處处署局]檔號|貴[處处署局]檔號/.test(line || '') ||
    (/檔號/.test(line || '') && /[:：]/.test(line || '')) ||
    (/日期/.test(line || '') && /[:：]/.test(line || ''));
}

function detectLetterType(text: string, hasOurRef: boolean): IncomingLetterType {
  if (/\bfrom\s*:/i.test(text) && /\bto\s*:/i.test(text) && /\bsubject\s*:/i.test(text)) {
    return 'email';
  }
  if (/\b(?:memorandum|circular|fax\s+cover)\b/i.test(text) || /內部通告|通函/.test(text)) {
    return 'memo';
  }
  if (isGovernmentText(text)) {
    return 'government';
  }
  if (/敬啟者|敬启者|此致|主旨|事由/.test(text) && !/\bdear\s+(sir|madam)/i.test(text)) {
    return 'chinese';
  }
  if (hasOurRef) {
    return 'consultant';
  }
  return 'unknown';
}

function isGovernmentText(text: string): boolean {
  return /hong kong special administrative region|\bhksar\b|the government of/i.test(text) ||
    /civil engineering and development|highways department|drainage services|water supplies|architectural services|environmental protection department|lands department/i.test(text) ||
    /土木工程拓展署|路政署|渠務署|水務署|建築署|環境保護署|地政總署|香港特別行政區政府/.test(text);
}

function organizationFromLetterhead(page: IOcrPageResult): string {
  const lines = groupWordsIntoLines(page.words || []);
  if (lines.length === 0) {
    return '';
  }
  const cutoff = letterheadCutoff(lines, page.height);
  const scored: { text: string; score: number }[] = [];
  for (let index = 0; index < lines.length; index++) {
    const line = lines[index];
    if (line.y0 > cutoff) {
      break;
    }
    const text = (line.text || '').replace(/\s+/g, ' ').trim();
    const score = organizationScore(text);
    if (score > 0) {
      scored.push({ text: cleanOrganization(text), score });
    }
  }
  scored.sort((left, right) => right.score - left.score);
  return scored.length > 0 ? scored[0].text : '';
}

function organizationFromLetterheadText(text: string): string {
  const lines = (text || '').split(/\r?\n/).map((line) => line.replace(/\s+/g, ' ').trim()).filter((line) => line.length > 0);
  const cutoff = Math.min(lines.length, 8);
  const scored: { text: string; score: number }[] = [];
  for (let index = 0; index < cutoff; index++) {
    if (isHeaderStopLine(lines[index])) {
      break;
    }
    const score = organizationScore(lines[index]);
    if (score > 0) {
      scored.push({ text: cleanOrganization(lines[index]), score });
    }
  }
  scored.sort((left, right) => right.score - left.score);
  return scored.length > 0 ? scored[0].text : '';
}

function letterheadCutoff(lines: ILine[], pageHeight: number): number {
  for (let index = 0; index < lines.length; index++) {
    if (isHeaderStopLine(lines[index].text)) {
      return lines[index].y0 - 2;
    }
  }
  return Math.max(40, pageHeight * 0.22);
}

function organizationScore(line: string): number {
  const text = (line || '').replace(/\s+/g, ' ').trim();
  if (!text || text.length < 4 || text.length > 90) {
    return 0;
  }
  if (isSkippedLetterheadLine(text)) {
    return 0;
  }
  let score = 1;
  if (/\b(?:department|bureau|authority|office|commission|council)\b/i.test(text) ||
    /署|處|局|委員會|有限公司/.test(text)) {
    score += 8;
  }
  if (/\b(?:limited|ltd|company|corporation|consultants?|engineers?|contractors?)\b/i.test(text)) {
    score += 6;
  }
  if (/^[A-Z][A-Za-z].{8,}/.test(text) || /[\u4e00-\u9fff]{4,}/.test(text)) {
    score += 2;
  }
  return score;
}

function isSkippedLetterheadLine(line: string): boolean {
  if (isAecomLine(line) || isHeaderStopLine(line)) {
    return true;
  }
  if (/the government of the hong kong special administrative region/i.test(line) ||
    /香港特別行政區政府/.test(line)) {
    return true;
  }
  if (/\b(?:tel|fax|email|www|http|confidential|restricted)\b/i.test(line)) {
    return true;
  }
  if (/\b\d{1,3}\s*[/\\\uFF0F]\s*f\b/i.test(line) ||
    /\b(?:road|street|avenue|drive|hong kong|kowloon)\b/i.test(line)) {
    return true;
  }
  return /^[\d\s./:-]+$/.test(line);
}

function isHeaderStopLine(line: string): boolean {
  const key = (line || '').toLowerCase().replace(/[^a-z0-9\u4e00-\u9fff]+/g, ' ').trim();
  return /^our\s+r/.test(key) ||
    /^your?\s+r/.test(key) ||
    /^date\b/.test(key) ||
    /^by\s+(post|hand|fax|email)/.test(key) ||
    /本[處处署局]檔號|貴[處处署局]檔號|檔號/.test(line || '') ||
    /日期/.test(line || '') && /[:：]/.test(line || '');
}

function isAecomLine(line: string): boolean {
  return /\baecom\b/i.test(line || '') || /艾奕康/.test(line || '');
}

function cleanOrganization(value: string): string {
  return (value || '')
    .replace(/<[^>]+>/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function dateFromLabeledLines(page?: IOcrPageResult): string {
  const labeled = labeledValue(page, DATE_LABELS);
  const parsed = parseIssueDate(labeled);
  return parsed ? formatIssueDate(parsed) : '';
}

function dateFromHeaderArea(page?: IOcrPageResult): string {
  if (!page) {
    return '';
  }
  const lines = groupWordsIntoLines(page.words || []);
  const cutoff = Math.max(40, page.height * 0.32);
  for (let index = 0; index < lines.length; index++) {
    if (lines[index].y0 > cutoff) {
      break;
    }
    if (isHeaderStopLine(lines[index].text) && !/^date\b/i.test(lines[index].text || '') && (lines[index].text || '').indexOf('日期') < 0) {
      continue;
    }
    const parsed = parseIssueDate(lines[index].text || '');
    if (parsed) {
      return formatIssueDate(parsed);
    }
  }
  const textLines = (page.text || '').split(/\r?\n/);
  const limit = Math.min(textLines.length, 12);
  for (let index = 0; index < limit; index++) {
    const parsed = parseIssueDate(textLines[index] || '');
    if (parsed) {
      return formatIssueDate(parsed);
    }
  }
  return '';
}

function labeledValue(page: IOcrPageResult | undefined, labels: string[]): string {
  if (!page) {
    return '';
  }
  return labeledValueFromWords(page, labels) || labeledValueFromText(page.text || '', labels);
}

function labeledValueFromWords(page: IOcrPageResult, labels: string[]): string {
  const lines = groupWordsIntoLines(page.words || []);
  for (let index = 0; index < lines.length; index++) {
    const match = matchLabeledLine(lines[index].text, labels);
    if (!match) {
      continue;
    }
    if (match.value) {
      return match.value;
    }
    const next = lines[index + 1];
    if (next && next.text && !matchLabeledLine(next.text, labels)) {
      return next.text.replace(/\s+/g, ' ').trim();
    }
  }
  return '';
}

function labeledValueFromText(text: string, labels: string[]): string {
  const lines = (text || '').split(/\r?\n/).map((line) => line.replace(/\s+/g, ' ').trim()).filter((line) => line.length > 0);
  for (let index = 0; index < lines.length; index++) {
    const match = matchLabeledLine(lines[index], labels);
    if (!match) {
      continue;
    }
    if (match.value) {
      return match.value;
    }
    const next = lines[index + 1];
    if (next && !matchLabeledLine(next, labels)) {
      return next;
    }
  }
  return '';
}

function matchLabeledLine(line: string, labels: string[]): { value: string } | undefined {
  const trimmed = (line || '').replace(/\s+/g, ' ').trim();
  if (!trimmed) {
    return undefined;
  }
  for (let index = 0; index < labels.length; index++) {
    const label = labels[index];
    const escaped = label.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const match = trimmed.match(new RegExp('^' + escaped + '(?:\\s*[:：.\\-]|\\s+|$)\\s*(.*)$', 'i'));
    if (match) {
      return { value: (match[1] || '').replace(/\s+/g, ' ').trim() };
    }
  }
  return undefined;
}

function parentheticalNameBelowClosing(page?: IOcrPageResult): string {
  if (!page) {
    return '';
  }
  const fromWords = parentheticalNameBelowClosingFromWords(page);
  if (fromWords) {
    return fromWords;
  }
  const fromLines = parentheticalNameBelowClosingFromLines(page);
  if (fromLines) {
    return fromLines;
  }
  return parentheticalNameBelowClosingFromText(page.text || '');
}

function parentheticalNameBelowClosingFromWords(page: IOcrPageResult): string {
  const closing = findIncomingClosingHit(page);
  if (!closing) {
    return '';
  }
  const maxY = closing.y1 + Math.max(220, (page.height || 0) * 0.22);
  const below = (page.words || []).filter((word) => {
    return word.y0 >= closing.y1 - 6 && word.y0 <= maxY;
  }).slice().sort((left, right) => {
    if (Math.abs(left.y0 - right.y0) > 8) {
      return left.y0 - right.y0;
    }
    return left.x0 - right.x0;
  });
  const sameLineRemainder = (page.words || []).filter((word) => {
    const midY = (word.y0 + word.y1) / 2;
    return midY >= closing.y0 - 4 &&
      midY <= closing.y1 + 4 &&
      word.x0 >= closing.x0 + 8;
  });
  const styled = [
    formatOcrTextWithStyles(sameLineRemainder),
    formatOcrTextWithStyles(below) || joinOcrWords(below)
  ].filter((item) => item && item.trim()).join('\n');
  const fromStyled = signatureParenthesesInner(styled);
  if (fromStyled) {
    return fromStyled;
  }
  let index = 0;
  const search = sameLineRemainder.concat(below);
  while (index < search.length) {
    const open = nextIncomingOpenParenIndex(search, index);
    if (open < 0) {
      return '';
    }
    const close = nextIncomingCloseParenIndex(search, open);
    const group = close >= 0
      ? search.slice(open, close + 1)
      : search.slice(open, Math.min(search.length, open + 12));
    if (group.length > 24) {
      index = open + 1;
      continue;
    }
    const joined = joinOcrWords(group);
    index = close >= 0 ? close + 1 : open + 1;
    if (matchAttnLine(stripIncomingMarkup(joined)) || isIncomingAttnParenInner(stripIncomingMarkup(joined))) {
      continue;
    }
    const inner = signatureParenthesesInner(joined);
    if (inner) {
      return inner;
    }
  }
  return '';
}

function findIncomingClosingHit(page: IOcrPageResult): IIncomingClosingHit | undefined {
  const dearY = incomingSalutationY(page);
  const minY = dearY >= 0 ? dearY + 6 : Math.max(40, (page.height || 0) * 0.40);
  const yours = findIncomingClosingHitByKind(page, minY, 'yours');
  if (yours) {
    return yours;
  }
  return findIncomingClosingHitByKind(page, minY, 'other');
}

function findIncomingYoursClosingHit(page: IOcrPageResult): IIncomingClosingHit | undefined {
  const dearY = incomingSalutationY(page);
  const minY = dearY >= 0 ? dearY + 6 : Math.max(40, (page.height || 0) * 0.40);
  const yours = findIncomingClosingHitByKind(page, minY, 'yours');
  if (yours) {
    return yours;
  }
  const lines = groupWordsIntoLines(page.words || []);
  for (let index = lines.length - 1; index >= 0; index--) {
    if (lines[index].y0 < minY || !isIncomingReceiverClosingPhrase(lines[index].text)) {
      continue;
    }
    return {
      x0: lines[index].x0,
      x1: lines[index].x1,
      y0: lines[index].y0,
      y1: lines[index].y1
    };
  }
  return undefined;
}

function findIncomingClosingHitByKind(
  page: IOcrPageResult,
  minY: number,
  kind: 'yours' | 'other'
): IIncomingClosingHit | undefined {
  const match = kind === 'yours' ? isIncomingYoursClosingPhrase : isIncomingOtherClosingPhrase;
  const lines = groupWordsIntoLines(page.words || []);
  let hit: IIncomingClosingHit | undefined;
  for (let index = 0; index < lines.length; index++) {
    const line = lines[index];
    if (line.y0 < minY) {
      continue;
    }
    const next = lines[index + 1];
    const combined = next && (next.y0 - line.y1) < Math.max(line.y1 - line.y0, 10) * 2.2
      ? (line.text + ' ' + next.text)
      : line.text;
    if (!match(line.text) && !match(combined)) {
      continue;
    }
    const extendToNext = !match(line.text) && !!next;
    hit = {
      x0: line.x0,
      x1: extendToNext && next ? Math.max(line.x1, next.x1) : line.x1,
      y0: line.y0,
      y1: extendToNext && next ? Math.max(line.y1, next.y1) : line.y1
    };
  }
  if (hit) {
    return hit;
  }
  return findIncomingClosingHitFromWords(page.words || [], minY, match);
}

function findIncomingClosingHitFromWords(
  words: IOcrWord[],
  minY: number,
  match: (line: string) => boolean
): IIncomingClosingHit | undefined {
  const sorted = (words || []).slice().sort((left, right) => {
    if (Math.abs(left.y0 - right.y0) > 10) {
      return left.y0 - right.y0;
    }
    return left.x0 - right.x0;
  });
  let hit: IIncomingClosingHit | undefined;
  for (let index = 0; index < sorted.length; index++) {
    if (sorted[index].y0 < minY) {
      continue;
    }
    const window = sorted.slice(index, index + 6);
    const lineMid = (sorted[index].y0 + sorted[index].y1) / 2;
    const sameLine = window.filter((word) => Math.abs((word.y0 + word.y1) / 2 - lineMid) < 14);
    const nearby = window.filter((word) => word.y0 <= sorted[index].y1 + 28);
    const phrase = nearby.map((word) => word.text || '').join(' ');
    if (!match(phrase) && !match(joinOcrWords(sameLine))) {
      continue;
    }
    const used = match(joinOcrWords(sameLine)) ? sameLine : nearby;
    let x0 = used[0].x0;
    let x1 = used[0].x1;
    let y0 = used[0].y0;
    let y1 = used[0].y1;
    used.forEach((word) => {
      x0 = Math.min(x0, word.x0);
      x1 = Math.max(x1, word.x1);
      y0 = Math.min(y0, word.y0);
      y1 = Math.max(y1, word.y1);
    });
    hit = { x0, x1, y0, y1 };
  }
  return hit;
}

function nextIncomingOpenParenIndex(words: IOcrWord[], start: number): number {
  for (let index = start; index < words.length; index++) {
    if (/[[(\uFF08【]/.test(words[index].text || '')) {
      return index;
    }
  }
  return -1;
}

function nextIncomingCloseParenIndex(words: IOcrWord[], start: number): number {
  for (let index = start; index < words.length; index++) {
    const text = words[index].text || '';
    if (index === start && /[[(\uFF08【]/.test(text) && !/[)\uFF09\]】]/.test(text)) {
      continue;
    }
    if (/[)\uFF09\]】]/.test(text)) {
      return index;
    }
  }
  return -1;
}

function parentheticalNameBelowClosingFromLines(page: IOcrPageResult): string {
  const lines = groupWordsIntoLines(page.words || []);
  const dearY = incomingSalutationY(page);
  const minY = dearY >= 0 ? dearY + 6 : Math.max(40, (page.height || 0) * 0.40);
  const closingIndex = lastIncomingClosingIndex(lines.map((line) => line.text), minY, lines);
  if (closingIndex < 0) {
    return '';
  }
  const after: string[] = [];
  const remainder = textAfterIncomingClosing(lines[closingIndex].text);
  if (remainder) {
    after.push(remainder);
  }
  for (let index = closingIndex + 1; index < lines.length && after.length < 12; index++) {
    if (lines[index].y0 < lines[closingIndex].y1 - 2) {
      continue;
    }
    const plain = stripIncomingMarkup(lines[index].text);
    if (after.length > 0 && isIncomingJobOrDeptLine(plain) && !hasUnclosedIncomingParen(after.join(' '))) {
      break;
    }
    after.push(lines[index].text);
  }
  let buffer = '';
  for (let index = 0; index < after.length; index++) {
    buffer = (buffer + ' ' + after[index]).replace(/\s+/g, ' ').trim();
    const inner = signatureParenthesesInner(buffer);
    if (inner) {
      return inner;
    }
  }
  return '';
}

function parentheticalNameBelowClosingFromText(text: string): string {
  const lines = stripIncomingMarkup(text || '').split(/\r?\n/).map((line) => line.replace(/\s+/g, ' ').trim()).filter((line) => line.length > 0);
  const closingIndex = lastIncomingClosingIndex(lines, 0);
  if (closingIndex < 0) {
    return '';
  }
  let buffer = textAfterIncomingClosing(lines[closingIndex]);
  const fromRemainder = signatureParenthesesInner(buffer);
  if (fromRemainder) {
    return fromRemainder;
  }
  const limit = Math.min(lines.length, closingIndex + 8);
  for (let index = closingIndex + 1; index < limit; index++) {
    if (isIncomingJobOrDeptLine(lines[index]) && !hasUnclosedIncomingParen(buffer)) {
      break;
    }
    buffer = (buffer + ' ' + lines[index]).replace(/\s+/g, ' ').trim();
    const inner = signatureParenthesesInner(buffer);
    if (inner) {
      return inner;
    }
  }
  return '';
}

function parentheticalNameAfter署名Label(page: IOcrPageResult): string {
  const text = stripIncomingMarkup((page.text || '') + '\n' + joinOcrWords(page.words || []));
  const labeled = text.match(/署\s*名\s*[:：]?\s*[[(\uFF08【]\s*([^)\uFF09\]】]{1,80})\s*[)\uFF09\]】]/);
  if (!labeled) {
    return '';
  }
  return cleanIncomingSenderName((labeled[1] || '').replace(/\s+/g, ' ').trim());
}

function hasUnclosedIncomingParen(text: string): boolean {
  const plain = stripIncomingMarkup(text);
  const lastOpen = Math.max(plain.lastIndexOf('('), plain.lastIndexOf('\uFF08'));
  const lastClose = Math.max(plain.lastIndexOf(')'), plain.lastIndexOf('\uFF09'));
  return lastOpen >= 0 && lastOpen > lastClose;
}

function lastIncomingClosingIndex(texts: string[], minY: number, lines?: ILine[]): number {
  let yours = -1;
  let other = -1;
  for (let index = 0; index < texts.length; index++) {
    if (lines && lines[index].y0 < minY) {
      continue;
    }
    const next = texts[index + 1];
    const combined = next ? (texts[index] + ' ' + next) : texts[index];
    if (isIncomingYoursClosingPhrase(texts[index]) || isIncomingYoursClosingPhrase(combined)) {
      yours = index;
    } else if (isIncomingOtherClosingPhrase(texts[index]) || isIncomingOtherClosingPhrase(combined)) {
      other = index;
    }
  }
  return yours >= 0 ? yours : other;
}

function isIncomingYoursClosingPhrase(line: string): boolean {
  const key = incomingClosingKey(line);
  if (!key) {
    return false;
  }
  return /(^|\s)yours\s+(sincere|faithful|truly)/.test(' ' + key) ||
    /(^|\s)(sincere|faithful)\s+yours/.test(' ' + key) ||
    /^(sincerely|faithfully|truly)$/.test(key);
}

function textAfterIncomingClosing(line: string): string {
  const plain = stripIncomingMarkup(line);
  const cut = plain.replace(/^[\s\S]*?\b(?:yours\s+(?:sincerely|faithfully|truly)|sincerely|faithfully|truly)\b[^\w\u4e00-\u9fff]*/i, '');
  if (!cut || cut === plain) {
    return '';
  }
  return cut.replace(/\s+/g, ' ').trim();
}

function isIncomingOtherClosingPhrase(line: string): boolean {
  const trimmed = (line || '').replace(/\s+/g, ' ').trim();
  if (/此致|順頌|顺颂|專此|专此|敬祝/.test(trimmed) && trimmed.length <= 40) {
    return true;
  }
  const key = incomingClosingKey(line);
  if (!key) {
    return false;
  }
  if (/^(best|kind|warm|with)?\s*(regards|regards as always)$/.test(key) ||
    /^(with\s+)?(kind|best|warm)\s+regards$/.test(key)) {
    return true;
  }
  return /^respectfully$/.test(key);
}

function incomingClosingKey(line: string): string {
  return stripIncomingMarkup(line || '').toLowerCase().replace(/[^a-z0-9]+/g, ' ').replace(/\s+/g, ' ').trim();
}

function stripIncomingMarkup(text: string): string {
  return (text || '')
    .replace(/<[^>]+>/g, ' ')
    .replace(/&nbsp;/gi, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function signatureParenthesesInner(text: string): string {
  const plain = stripIncomingMarkup(text);
  const inners: string[] = [];
  const complete = /[[(\uFF08【]\s*([^)\uFF09\]】]{1,120}?)\s*[)\uFF09\]】]/g;
  let match: RegExpExecArray | null = complete.exec(plain);
  while (match) {
    inners.push((match[1] || '').replace(/\s+/g, ' ').trim());
    match = complete.exec(plain);
  }
  if (inners.length === 0) {
    const open = plain.search(/[[(\uFF08【]/);
    if (open >= 0) {
      const rest = plain.slice(open + 1).replace(/[)\uFF09\]】][\s\S]*$/, '').replace(/\s+/g, ' ').trim();
      if (rest) {
        inners.push(rest);
      }
    }
  }
  for (let index = 0; index < inners.length; index++) {
    const inner = inners[index];
    if (!inner || isIncomingIgnorableParen(inner) || isIncomingAttnParenInner(inner) || isIncomingJobOrDeptLine(inner)) {
      continue;
    }
    const name = cleanIncomingSenderName(inner);
    if (name) {
      return name;
    }
  }
  return '';
}

function isIncomingAttnParenInner(value: string): boolean {
  const text = stripIncomingMarkup(value);
  if (!text) {
    return false;
  }
  if (matchAttnLine(text)) {
    return true;
  }
  return /^(?:for\s+the\s+)?(?:attn|atin|attm|atln|attention)(?:\s+of)?\s*[:：.]/i.test(text);
}

function isIncomingIgnorableParen(value: string): boolean {
  const key = (value || '').toLowerCase().replace(/[^a-z0-9]+/g, ' ').replace(/\s+/g, ' ').trim();
  return key === 'signed' ||
    key === 'signature' ||
    key === 'sgd' ||
    key === 'chop' ||
    key === 'seal' ||
    key === 'company chop';
}

function isIncomingJobOrDeptLine(value: string): boolean {
  const key = (value || '').toLowerCase().replace(/[^a-z0-9\u4e00-\u9fff]+/g, ' ').replace(/\s+/g, ' ').trim();
  if (!key) {
    return false;
  }
  if (/\bbuilding\s+services\b/.test(key) ||
    /\b(senior|junior|principal|chief|assistant|associate|acting)\s+(engineer|architect|surveyor|manager|director|officer|inspector|consultant|planner|building)\b/.test(key) ||
    /^(for\s+)?(director|chief engineer|project manager|building services)\b/.test(key) ||
    /\b(department|division|section|unit)\b/.test(key)) {
    return true;
  }
  return /屋宇裝備|高級.*工程|工程師|總監|經理|署長|處長|主任/.test(value || '');
}

function cleanIncomingSenderName(value: string): string {
  const text = (value || '')
    .replace(/<[^>]+>/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
  if (!text || text.length < 2 || text.length > 80 || isIncomingIgnorableParen(text) || isIncomingJobOrDeptLine(text)) {
    return '';
  }
  return text;
}

function eightDigitRun(value: string): string {
  const match = (value || '').match(/\d{8}/);
  return match ? match[0] : '';
}

function pageText(page?: IOcrPageResult): string {
  if (!page) {
    return '';
  }
  return ((page.text || '') + '\n' + joinOcrWords(page.words || [])).replace(/\s+/g, ' ');
}

function groupWordsIntoLines(words: IOcrWord[]): ILine[] {
  const sorted = (words || []).slice().sort((left, right) => {
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
    group.forEach((item) => {
      x0 = Math.min(x0, item.x0);
      x1 = Math.max(x1, item.x1);
      y0 = Math.min(y0, item.y0);
      y1 = Math.max(y1, item.y1);
    });
    return {
      text: joinOcrWords(group).replace(/\s+/g, ' ').trim(),
      x0,
      x1,
      y0,
      y1,
      words: group
    };
  });
}
