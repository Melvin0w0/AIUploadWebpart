// Incoming OCR locators only. Outgoing lives in services/outgoing/.
import { formatIssueDate, parseIssueDate } from '../../constants/issueDate';
import { projectNumberFromRef } from '../../constants/projectNumber';
import { extractYourRefNo } from '../fieldExtractor';
import { IOcrPageResult, IOcrWord } from '../IPdfOcr';
import { joinOcrWords } from '../ocrSelection';
import { extractSubjectBelowDearSir } from '../signatureSender';

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
  const fromWords = organizationFromLetterhead(page);
  if (fromWords) {
    return fromWords;
  }
  return organizationFromLetterheadText(page.text || '');
}

export async function extractIncomingSubject(page?: IOcrPageResult): Promise<string> {
  if (!page) {
    return '';
  }
  const fromChinese = subjectFromChineseLabel(page);
  if (fromChinese) {
    return fromChinese;
  }
  const fromMemo = labeledValue(page, MEMO_SUBJECT_LABELS);
  if (fromMemo) {
    return fromMemo;
  }
  try {
    const fromOutgoingRules = await extractSubjectBelowDearSir(page);
    if (fromOutgoingRules) {
      return fromOutgoingRules;
    }
  } catch {
    return '';
  }
  return '';
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

function subjectFromChineseLabel(page: IOcrPageResult): string {
  const fromWords = labeledValueFromWords(page, CHINESE_SUBJECT_LABELS);
  if (fromWords) {
    return fromWords;
  }
  return labeledValueFromText(page.text || '', CHINESE_SUBJECT_LABELS);
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
      y1
    };
  });
}

interface ILine {
  text: string;
  x0: number;
  x1: number;
  y0: number;
  y1: number;
}

const CHINESE_SUBJECT_LABELS: string[] = ['主旨', '事由', '關於', '关于'];
const DATE_LABELS: string[] = ['Date', '日期'];
const MEMO_SUBJECT_LABELS: string[] = ['Subject', 'Re'];
const MEMO_FROM_LABELS: string[] = ['From', '發件人', '寄件人'];
const MEMO_TO_LABELS: string[] = ['To', '收件人', '致'];
