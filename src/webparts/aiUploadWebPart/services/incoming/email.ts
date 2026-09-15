// Incoming email printout rules only. Letter locators stay in extract.ts.
import { formatIssueDate, parseIssueDate } from '../../constants/issueDate';
import { IOcrPageResult, OcrPageDecision } from '../IPdfOcr';
import { joinOcrWords } from '../ocrSelection';

export interface IIncomingEmailHeaders {
  from: string;
  to: string;
  cc: string;
  subject: string;
  sent: string;
  agreementNo: string;
  ourRef: string;
  yourRef: string;
  text: string;
}

export function looksLikeIncomingEmail(text: string): boolean {
  const source = text || '';
  const from = /\bfrom\s*:/i.test(source) || /發件人|寄件人/.test(source);
  const to = /\bto\s*:/i.test(source) || /收件人|收件者/.test(source);
  const subject = /\bsubject\s*:/i.test(source) || /主旨\s*[:：]/.test(source);
  const sent = /\bsent\s*:/i.test(source);
  const attachments = /\battachments?\s*:/i.test(source);
  return (from && to && (subject || sent)) || (from && sent && (subject || to || attachments));
}

export function isIncomingEmailPage(page?: IOcrPageResult): boolean {
  if (!page) {
    return false;
  }
  return looksLikeIncomingEmail(incomingEmailPlainText(page) || page.text || '');
}

export function incomingPagesLookLikeEmail(pages?: IOcrPageResult[]): boolean {
  const list = pages || [];
  for (let index = 0; index < list.length; index++) {
    if (isIncomingEmailPage(list[index])) {
      return true;
    }
    const next = list[index + 1];
    if (!next) {
      continue;
    }
    if (looksLikeIncomingEmail(incomingJoinedPageText(list[index], next))) {
      return true;
    }
  }
  return false;
}

export function createIncomingOcrDecider(): (
  previous: IOcrPageResult | undefined,
  current: IOcrPageResult
) => OcrPageDecision {
  let emailMode = false;
  let letterAfterEmail = false;
  return (previous, current) => {
    const isEmailPage = incomingPageIsEmailSegment(previous, current);
    const continuation = emailMode && !letterAfterEmail && incomingPageLooksLikeEmailContinuation(previous, current);
    if (!letterAfterEmail && (isEmailPage || continuation)) {
      emailMode = true;
      return 'keep-continue';
    }
    if (emailMode && !letterAfterEmail) {
      if (incomingPageLooksLikeFollowingLetter(current) || incomingPageHasOurRef(current)) {
        letterAfterEmail = true;
        return (incomingPageHasOurRef(current) || incomingPageHasClosing(current.text))
          ? 'keep-stop'
          : 'keep-continue';
      }
      return 'drop-stop';
    }
    if (letterAfterEmail) {
      return (incomingPageHasOurRef(current) || incomingPageHasClosing(current.text))
        ? 'keep-stop'
        : 'keep-continue';
    }
    if (incomingPageHasClosing(current.text) && !incomingPageBottomLooksLikeEmail(current)) {
      return 'keep-stop';
    }
    return 'keep-continue';
  };
}

export function keepIncomingEmailPreviewPages(pages?: IOcrPageResult[]): IOcrPageResult[] {
  const list = pages || [];
  if (list.length === 0 || !incomingPagesLookLikeEmail(list)) {
    return list;
  }
  const kept: IOcrPageResult[] = [];
  for (let index = 0; index < list.length; index++) {
    const previous = index > 0 ? list[index - 1] : undefined;
    const current = list[index];
    if (kept.length > 0 && incomingPageLooksLikeFollowingLetter(current)) {
      break;
    }
    if (incomingPageIsEmailSegment(previous, current) ||
      (kept.length > 0 && incomingPageLooksLikeEmailContinuation(previous, current))) {
      kept.push(current);
      continue;
    }
    if (kept.length > 0) {
      break;
    }
  }
  return kept.length > 0 ? kept : list;
}

export function incomingLastEmailText(pages?: IOcrPageResult[]): string {
  return sliceLastIncomingEmail(incomingEmailPagesPlainText(pages || []));
}

export function sliceLastIncomingEmail(text: string): string {
  const slices = incomingEmailSlices(text);
  return slices.length > 0 ? slices[slices.length - 1] : (text || '').replace(/\s+/g, ' ').trim();
}

export function extractIncomingEmailCcFromLastPage(pages?: IOcrPageResult[]): string {
  return extractHeaderFromLastEmailPage(pages, 'cc');
}

export function extractIncomingEmailToFromLastPage(pages?: IOcrPageResult[]): string {
  return extractHeaderFromLastEmailPage(pages, 'to');
}

export function extractIncomingEmailSubjectFromLastPage(pages?: IOcrPageResult[]): string {
  const list = pages || [];
  const lastIndex = lastIncomingEmailFormatPageIndex(list);
  if (lastIndex < 0) {
    return subjectValueUntilNextTitle(incomingLastEmailText(list));
  }
  const pageText = incomingEmailPlainText(list[lastIndex]);
  const fromPage = subjectValueUntilNextTitle(sliceLastIncomingEmail(clipEmailBeforeLetter(pageText)));
  if (fromPage) {
    return fromPage;
  }
  if (lastIndex > 0) {
    const joined = (incomingEmailPlainText(list[lastIndex - 1]) + ' ' + pageText).replace(/\s+/g, ' ').trim();
    const fromJoined = subjectValueUntilNextTitle(sliceLastIncomingEmail(clipEmailBeforeLetter(joined)));
    if (fromJoined) {
      return fromJoined;
    }
  }
  return subjectValueUntilNextTitle(incomingLastEmailText(list));
}

export function extractIncomingEmailOurRef(pages?: IOcrPageResult[]): string {
  const list = pages || [];
  const lastEmailIndex = lastIncomingEmailFormatPageIndex(list);
  const emailPages = lastEmailIndex >= 0 ? list.slice(0, lastEmailIndex + 1) : list;
  const letterPages = lastEmailIndex >= 0 ? list.slice(lastEmailIndex + 1) : [];
  const emailText = incomingEmailPagesPlainText(emailPages).replace(/\s+/g, ' ').trim();
  const starts = findIncomingEmailStarts(emailText);
  const lastStart = starts.length > 0 ? starts[starts.length - 1] : 0;
  const fromLastEmail = ourRefValueFromEmailText(emailText.substring(lastStart));
  if (fromLastEmail) {
    return fromLastEmail;
  }
  if (lastEmailIndex >= 0) {
    const fromLastPage = ourRefValueFromEmailText(incomingEmailPlainText(list[lastEmailIndex]));
    if (fromLastPage) {
      return fromLastPage;
    }
  }
  for (let index = 0; index < letterPages.length; index++) {
    const value = ourRefValueFromEmailText(incomingEmailPlainText(letterPages[index]));
    if (value) {
      return value;
    }
  }
  return '';
}

export function extractIncomingEmailHeaders(pages?: IOcrPageResult[]): IIncomingEmailHeaders {
  const text = incomingLastEmailText(pages);
  const headers = incomingEmailHeaderBlock(text);
  return {
    from: incomingEmailPersonName(extractIncomingEmailHeaderValue(headers, 'from')),
    to: extractIncomingEmailToFromLastPage(pages) || incomingEmailPersonName(extractIncomingEmailHeaderValue(headers, 'to')),
    cc: extractIncomingEmailCcFromLastPage(pages) || incomingEmailCcSenderName(extractIncomingEmailHeaderValue(headers, 'cc')),
    subject: extractIncomingEmailSubjectFromLastPage(pages),
    sent: parseIncomingEmailSentDate(
      extractIncomingEmailHeaderValue(headers, 'sent') || extractIncomingEmailHeaderValue(headers, 'date')
    ),
    agreementNo: extractIncomingEmailAgreementNo(text),
    ourRef: extractIncomingEmailOurRef(pages) || extractIncomingEmailHeaderValue(text, 'ourRef'),
    yourRef: extractIncomingEmailHeaderValue(text, 'yourRef'),
    text
  };
}

export function incomingLastEmailPages(pages?: IOcrPageResult[]): IOcrPageResult[] {
  const list = pages || [];
  if (list.length <= 1) {
    return list;
  }
  const lastText = incomingLastEmailText(list);
  if (!lastText) {
    return list;
  }
  const snippet = lastText.substring(0, Math.min(120, lastText.length));
  let start = 0;
  for (let index = 0; index < list.length; index++) {
    if (incomingEmailPagesPlainText(list.slice(index)).indexOf(snippet) >= 0) {
      start = index;
    }
  }
  return list.slice(start);
}

function lastIncomingEmailFormatPageIndex(pages: IOcrPageResult[]): number {
  let last = -1;
  for (let index = 0; index < pages.length; index++) {
    const previous = index > 0 ? pages[index - 1] : undefined;
    const current = pages[index];
    if (last >= 0 && incomingPageLooksLikeFollowingLetter(current)) {
      break;
    }
    if (incomingPageIsEmailSegment(previous, current) ||
      (last >= 0 && incomingPageLooksLikeEmailContinuation(previous, current))) {
      last = index;
    }
  }
  return last;
}

function extractHeaderFromLastEmailPage(pages: IOcrPageResult[] | undefined, kind: 'to' | 'cc'): string {
  const list = pages || [];
  const lastIndex = lastIncomingEmailFormatPageIndex(list);
  if (lastIndex < 0) {
    return headerFromLastEmailInText(incomingEmailPagesPlainText(list), kind);
  }
  const pageText = incomingEmailPlainText(list[lastIndex]);
  const fromPage = headerFromLastEmailInText(pageText, kind);
  if (fromPage) {
    return fromPage;
  }
  if (lastIndex === 0) {
    return '';
  }
  const joined = (incomingEmailPlainText(list[lastIndex - 1]) + ' ' + pageText).replace(/\s+/g, ' ').trim();
  return headerFromLastEmailInText(joined, kind);
}

function headerFromLastEmailInText(text: string, kind: 'to' | 'cc'): string {
  const lastSlice = sliceLastIncomingEmail(clipEmailBeforeLetter(text));
  const headers = incomingEmailHeaderBlock(lastSlice);
  const value = extractIncomingEmailHeaderValue(headers, kind);
  return kind === 'cc' ? incomingEmailCcSenderName(value) : incomingEmailPersonName(value);
}

function ourRefValueFromEmailText(text: string): string {
  const source = (text || '').replace(/\s+/g, ' ').trim();
  if (!source) {
    return '';
  }
  const match = source.match(/\bour\s+r+e+f(?:erence)?(?:\s*no(?:\.)?)?\s*[:：.-]+\s*/i) ||
    source.match(/本[處处署局]檔號\s*[:：.-]+\s*/) ||
    source.match(/本函編號\s*[:：.-]+\s*/);
  if (!match || match.index === undefined) {
    return '';
  }
  let value = source.substring(match.index + match[0].length).replace(/\s+/g, ' ').trim();
  const stop = value.search(/(?:^|\s)(?:your?\s+r+e+f|date|from|sent|to|cc|subject|dear|tel|fax|email|yours|貴[處处署局]檔號|來函編號|日期|敬啟者|敬启者)\b/i);
  if (stop >= 0) {
    value = value.substring(0, stop);
  }
  value = value.replace(/^[:.\s-]+/, '').replace(/[;；]+$/g, '').replace(/\s+/g, ' ').trim();
  return value.length >= 2 && value.length <= 120 ? value : '';
}

function incomingEmailSlices(text: string): string[] {
  const source = (text || '').replace(/\s+/g, ' ').trim();
  if (!source) {
    return [];
  }
  const starts = findIncomingEmailStarts(source);
  if (starts.length === 0) {
    const clipped = clipEmailBeforeLetter(source);
    return clipped ? [clipped] : [];
  }
  const slices: string[] = [];
  for (let index = 0; index < starts.length; index++) {
    const rawEnd = index + 1 < starts.length ? starts[index + 1] : source.length;
    const slice = clipEmailBeforeLetter(source.substring(starts[index], rawEnd));
    if (slice) {
      slices.push(slice);
    }
  }
  return slices;
}

function clipEmailBeforeLetter(emailText: string): string {
  const source = (emailText || '').replace(/\s+/g, ' ').trim();
  if (!source) {
    return '';
  }
  const headerEnd = incomingEmailHeaderBlockEnd(source);
  const rest = source.substring(headerEnd);
  const letter = rest.search(/\bour\s+r+e+f(?:erence)?\s*[:：.]|本[處处署局]檔號|本函編號|敬啟者|敬启者|\byours\s+(?:faithfully|sincerely|truly)\b/i);
  if (letter >= 0) {
    return source.substring(0, headerEnd + letter).replace(/\s+/g, ' ').trim();
  }
  return source;
}

function incomingEmailHeaderBlock(text: string): string {
  const source = (text || '').replace(/\s+/g, ' ').trim();
  if (!source) {
    return '';
  }
  return source.substring(0, incomingEmailHeaderBlockEnd(source)).replace(/\s+/g, ' ').trim();
}

function incomingEmailHeaderBlockEnd(text: string): number {
  const headerLabel = /(?:^|\s)(?:from|sent|to|cc|bcc|subject|re|date|attachments?|importance|reply-to|發件人|寄件人|收件人|主旨|附件|抄送)\s*[:：]/ig;
  const positions: number[] = [];
  let match = headerLabel.exec(text);
  while (match) {
    if (positions.length > 0 && match.index - positions[positions.length - 1] > 320) {
      break;
    }
    positions.push(match.index);
    if (positions.length >= 12) {
      break;
    }
    match = headerLabel.exec(text);
  }
  if (positions.length === 0) {
    return Math.min(text.length, 900);
  }
  const last = positions[positions.length - 1];
  const after = text.substring(last);
  const rest = after.replace(/^(?:^|\s)(?:from|sent|to|cc|bcc|subject|re|date|attachments?|importance|reply-to|發件人|寄件人|收件人|主旨|附件|抄送)\s*[:：]\s*/i, '');
  const valueEndRel = rest.search(/(?:^|\s)(?:from|sent|to|cc|bcc|subject|re|date|attachments?|importance|reply-to|dear|yours|our\s+r+e+f|發件人|寄件人|收件人|主旨|附件|抄送|敬啟者)\s*[:：]?/i);
  const valueLen = valueEndRel >= 0 ? valueEndRel : Math.min(rest.length, 400);
  return Math.min(text.length, last + (after.length - rest.length) + valueLen);
}

function findIncomingEmailStarts(text: string): number[] {
  const starts: number[] = [];
  const seen: { [index: number]: boolean } = {};
  const add = (index: number): void => {
    if (index < 0 || index >= text.length || seen[index]) {
      return;
    }
    seen[index] = true;
    starts.push(index);
  };

  add(text.search(/\b(?:from|sent|to|subject)\s*:|發件人|寄件人|主旨\s*[:：]/i));

  const marker = /(?:-----+\s*)?(?:original\s+message|forwarded\s+message|begin\s+forwarded\s+message|轉發的郵件|转发的邮件|原始郵件|原始邮件)\b/ig;
  let match = marker.exec(text);
  while (match) {
    const after = text.substring(match.index);
    const header = after.search(/\b(?:from|sent|to|subject)\s*:|發件人|寄件人|主旨\s*[:：]/i);
    add(header >= 0 ? match.index + header : match.index);
    match = marker.exec(text);
  }

  starts.sort((left, right) => left - right);
  for (let index = 0; index < starts.length && starts.length < 20; index++) {
    const nextRel = findSecondIncomingEmailHeader(text.substring(starts[index]));
    if (nextRel > 0) {
      add(starts[index] + nextRel);
      starts.sort((left, right) => left - right);
    }
  }
  return starts;
}

function findSecondIncomingEmailHeader(text: string): number {
  const headerLabel = /(?:^|\s)(?:from|sent|to|cc|bcc|subject|re|date|attachments?|importance|reply-to|發件人|寄件人|收件人|主旨|附件|抄送)\s*[:：]/ig;
  const positions: number[] = [];
  let match = headerLabel.exec(text);
  while (match) {
    if (positions.length > 0 && match.index - positions[positions.length - 1] > 320) {
      return match.index;
    }
    positions.push(match.index);
    if (positions.length >= 12) {
      break;
    }
    match = headerLabel.exec(text);
  }
  if (positions.length === 0) {
    return -1;
  }
  const restStart = positions[positions.length - 1];
  const rest = text.substring(restStart);
  const skipped = rest.replace(/^(?:^|\s)(?:from|sent|to|cc|bcc|subject|re|date|attachments?|importance|reply-to|發件人|寄件人|收件人|主旨|附件|抄送)\s*[:：]\s*/i, '');
  const afterFirstLastHeader = rest.length - skipped.length;
  const tail = rest.substring(afterFirstLastHeader);
  const second = tail.search(/\bfrom\s*:[\s\S]{0,240}?\b(?:sent|to|subject)\s*:/i);
  if (second >= 0) {
    return restStart + afterFirstLastHeader + second;
  }
  const secondZh = tail.search(/發件人\s*[:：][\s\S]{0,240}?(?:收件人|主旨)\s*[:：]/);
  if (secondZh >= 0) {
    return restStart + afterFirstLastHeader + secondZh;
  }
  return -1;
}

function extractIncomingEmailSubjectFromText(text: string): string {
  return subjectValueUntilNextTitle(text);
}

function subjectValueUntilNextTitle(text: string): string {
  const source = (text || '').replace(/\s+/g, ' ').trim();
  if (!source) {
    return '';
  }
  const label = /(?:^|\s)(?:subject|主旨|事由)\s*[:：]\s*/ig;
  let last: RegExpExecArray | null = null;
  let match = label.exec(source);
  while (match) {
    last = match;
    match = label.exec(source);
  }
  if (!last || last.index === undefined) {
    return '';
  }
  let value = source.substring(last.index + last[0].length);
  const nextTitle = value.search(NEXT_EMAIL_HEADER_TITLE);
  if (nextTitle >= 0) {
    value = value.substring(0, nextTitle);
  }
  return value.replace(/\s+/g, ' ').trim();
}

const NEXT_EMAIL_HEADER_TITLE = /(?:^|\s)(?:from|sent|to|cc|bcc|date|subject|attachments?|importance|reply-to|sensitivity|thread-topic|thread-index|發件人|寄件人|收件人|主旨|附件|抄送|發送時間)\s*[:：]|[A-Za-z](?:Sent|From|To|Cc|Bcc|Date|Subject|Attachments?|Importance|Reply-To)\s*[:：]|\b(?:Sent|From|To|Cc|Bcc|Date|Subject|Attachments?|Importance|Reply-To)\s*[:：]/;

function extractIncomingEmailHeaderValue(
  text: string,
  kind: 'from' | 'to' | 'cc' | 'sent' | 'date' | 'subject' | 'ourRef' | 'yourRef'
): string {
  const patterns: { [key: string]: RegExp } = {
    from: /(?:^|\s)(?:from|發件人|寄件人)\s*[:：]\s*/i,
    to: /(?:^|\s)(?:to|收件人|收件者)\s*[:：]\s*/i,
    cc: /(?:^|\s)(?:cc|c\.c\.|抄送)\s*[:：]\s*/i,
    sent: /(?:^|\s)(?:sent|發送時間|传送时间)\s*[:：]\s*/i,
    date: /(?:^|\s)(?:date|日期)\s*[:：]\s*/i,
    subject: /(?:^|\s)(?:subject|re|主旨|事由)\s*[:：]\s*/i,
    ourRef: /(?:^|\s)(?:our\s+r+e+f|本[處处署局]檔號|本函編號)\s*[:：]?\s*/i,
    yourRef: /(?:^|\s)(?:your?\s+r+e+f|貴[處处署局]檔號|來函編號)\s*[:：]?\s*/i
  };
  return valueUntilNextEmailTitle(text, patterns[kind]);
}

function extractIncomingEmailAgreementNo(text: string): string {
  const match = (text || '').match(
    /(?:agreement\s*(?:no\.?|number)|contract\s*(?:no\.?|number)|agmt\.?\s*no\.?|agt\.?\s*no\.?|合約編號|合同編號|協議編號|协议编号|合約號碼|合同号码)\s*[:.\-\uFF1A]?\s*/i
  );
  if (!match || match.index === undefined) {
    return '';
  }
  let value = text.substring(match.index + match[0].length).replace(/\s+/g, ' ').trim();
  const stop = value.search(/(?:^|\s)(?:subject|re|dear|yours|from|sent|to|cc|主旨|敬啟者|敬启者)\b/i);
  if (stop > 0) {
    value = value.substring(0, stop);
  }
  value = value.replace(/[;；]+$/g, '').replace(/\s+/g, ' ').trim();
  if (!value || value.length < 4 || !/\d/.test(value)) {
    return '';
  }
  return value.length > 80 ? value.substring(0, 80).trim() : value;
}

function incomingEmailPersonName(value: string): string {
  const first = incomingEmailFirstRecipient(value);
  return incomingEmailDisplayName(first);
}

function incomingEmailCcSenderName(value: string): string {
  return incomingEmailSurnameGivenName(incomingEmailPersonName(value));
}

function incomingEmailFirstRecipient(value: string): string {
  const source = (value || '').replace(/\s+/g, ' ').trim();
  if (!source) {
    return '';
  }
  const bySemi = source.split(/\s*;\s*/).map((part) => part.trim()).filter((part) => !!part);
  const first = bySemi[0] || source;
  if (bySemi.length > 1) {
    return first;
  }
  const comma = first.split(/\s*,\s*/).map((part) => part.trim()).filter((part) => !!part);
  if (comma.length >= 2 && incomingEmailWordCount(comma[0]) >= 2) {
    return comma[0];
  }
  return first;
}

function incomingEmailWordCount(value: string): number {
  return (value || '').replace(/\s+/g, ' ').trim().split(' ').filter((part) => !!part).length;
}

function incomingEmailSurnameGivenName(value: string): string {
  const text = (value || '').replace(/\s+/g, ' ').trim();
  if (!text) {
    return '';
  }
  if (/,/.test(text)) {
    return text.replace(/\s*,\s*/, ', ').replace(/\s+/g, ' ').trim();
  }
  if (/[\u4e00-\u9fff]/.test(text) && !/[A-Za-z]{2,}/.test(text)) {
    return text;
  }
  const parts = text.split(/\s+/);
  if (parts.length < 2) {
    return text;
  }
  return parts[parts.length - 1] + ', ' + parts.slice(0, -1).join(' ');
}

function incomingEmailDisplayName(value: string): string {
  let text = (value || '').replace(/\s+/g, ' ').trim();
  if (!text) {
    return '';
  }
  text = text.replace(/^["'“”‘’]+|["'“”‘’]+$/g, '');
  const mailto = text.match(/^(.*?)\s*\[mailto:[^\]]+\]$/i);
  if (mailto) {
    text = (mailto[1] || '').trim();
  }
  const angle = text.match(/^(.*?)\s*<[^>]+>$/);
  if (angle) {
    text = (angle[1] || '').trim();
  }
  text = text.replace(/^["'“”‘’]+|["'“”‘’]+$/g, '').replace(/\s+/g, ' ').trim();
  if (/^[^@\s]+@[^@\s]+$/.test(text)) {
    return text.split('@')[0].replace(/[._]+/g, ' ').trim();
  }
  return text;
}

function parseIncomingEmailSentDate(value: string): string {
  const cleaned = (value || '')
    .replace(/\s+/g, ' ')
    .trim()
    .replace(/^(?:monday|tuesday|wednesday|thursday|friday|saturday|sunday)\s*,?\s*/i, '')
    .replace(/^(?:星期[一二三四五六日天])\s*,?\s*/, '')
    .replace(/\s+\d{1,2}:\d{2}(?::\d{2})?(?:\s*[AP]M)?(?:\s+\S+)?$/i, '');
  const parsed = parseIssueDate(cleaned) || parseIssueDate(value);
  return parsed ? formatIssueDate(parsed) : '';
}

function firstNonHeaderEmailTitleValue(text: string): string {
  const label = /(?:^|\s)([A-Za-z][A-Za-z0-9 /&-]{0,40}|[\u4e00-\u9fff]{1,12})\s*[:：]\s*/g;
  let match = label.exec(text);
  while (match) {
    const title = (match[1] || '').replace(/\s+/g, ' ').trim();
    if (!isIncomingEmailHeaderTitle(title)) {
      const rest = text.substring(match.index + match[0].length);
      const next = rest.search(/(?:^|\s)(?:[A-Za-z][A-Za-z0-9 /&-]{0,40}|[\u4e00-\u9fff]{1,12}|sent|from|to|cc|bcc|date|subject|re|attachments?|importance|reply-to|發件人|寄件人|收件人|主旨|附件)\s*[:：]/i);
      const value = (next >= 0 ? rest.substring(0, next) : rest).replace(/\s+/g, ' ').trim();
      if (value.length >= 2 && value.length <= 400) {
        return value;
      }
    }
    match = label.exec(text);
  }
  return '';
}

function valueUntilNextEmailTitle(text: string, label: RegExp): string {
  const match = text.match(label);
  if (!match || match.index === undefined) {
    return '';
  }
  const rest = text.substring(match.index + match[0].length);
  const next = rest.search(/(?:^|\s)(?:sent|from|to|cc|bcc|date|subject|re|attachments?|importance|reply-to|發件人|寄件人|收件人|主旨|附件|抄送)\s*[:：]/i);
  const value = (next >= 0 ? rest.substring(0, next) : rest).replace(/\s+/g, ' ').trim();
  return value.length >= 2 && value.length <= 400 ? value : '';
}

function isIncomingEmailHeaderTitle(title: string): boolean {
  const key = (title || '').toLowerCase().replace(/\s+/g, ' ').trim();
  return key === 'from' ||
    key === 'to' ||
    key === 'cc' ||
    key === 'bcc' ||
    key === 'sent' ||
    key === 'date' ||
    key === 'subject' ||
    key === 're' ||
    key === 'attachment' ||
    key === 'attachments' ||
    key === 'importance' ||
    key === 'reply-to' ||
    key === 'reply to' ||
    key === '發件人' ||
    key === '寄件人' ||
    key === '收件人' ||
    key === '主旨' ||
    key === '附件' ||
    key === '抄送';
}

function incomingEmailPlainText(page?: IOcrPageResult): string {
  if (!page) {
    return '';
  }
  const fromWords = joinOcrWords(page.words || []).replace(/\s+/g, ' ').trim();
  if (fromWords) {
    return stripEmailMarkup(fromWords);
  }
  return stripEmailMarkup((page.text || '').replace(/\s+/g, ' ').trim());
}

function incomingEmailPagesPlainText(pages: IOcrPageResult[]): string {
  return (pages || []).map((page) => incomingEmailPlainText(page)).filter((text) => !!text).join(' ').replace(/\s+/g, ' ').trim();
}

function incomingJoinedPageText(page: IOcrPageResult, next: IOcrPageResult): string {
  const full = (incomingEmailPlainText(page) + ' ' + incomingEmailPlainText(next)).replace(/\s+/g, ' ').trim();
  const split = (incomingPageBandText(page, 0.4, 1) + ' ' + incomingPageBandText(next, 0, 0.6)).replace(/\s+/g, ' ').trim();
  return (full + ' ' + split).replace(/\s+/g, ' ').trim();
}

function incomingPageBandText(page: IOcrPageResult, fromRatio: number, toRatio: number): string {
  const height = page.height || 0;
  const top = height * fromRatio;
  const bottom = height * toRatio;
  const words = (page.words || []).filter((word) => {
    const y = word.y0;
    return y >= top && y <= bottom;
  });
  if (words.length > 0) {
    return stripEmailMarkup(joinOcrWords(words).replace(/\s+/g, ' ').trim());
  }
  return '';
}

function incomingPageIsEmailSegment(previous: IOcrPageResult | undefined, current: IOcrPageResult): boolean {
  if (isIncomingEmailPage(current) || incomingPageBottomLooksLikeEmail(current)) {
    return true;
  }
  return !!(previous && looksLikeIncomingEmail(incomingJoinedPageText(previous, current)));
}

function incomingPageLooksLikeEmailContinuation(previous: IOcrPageResult | undefined, current: IOcrPageResult): boolean {
  if (incomingPageLooksLikeFollowingLetter(current)) {
    return false;
  }
  if (incomingPageHasEmailHeaderTokens(current)) {
    return true;
  }
  if (previous && looksLikeIncomingEmail(incomingJoinedPageText(previous, current))) {
    return true;
  }
  return !!(previous && incomingPageBottomLooksLikeEmail(previous) && incomingPageHasEmailHeaderTokens(previous));
}

function incomingPageLooksLikeStandaloneDocument(page: IOcrPageResult): boolean {
  const top = incomingPageBandText(page, 0, 0.35);
  if (!top || incomingPageHasEmailHeaderTokensInText(top)) {
    return false;
  }
  return isGovernmentEmailText(top) ||
    /\bour\s+r+e+f/i.test(top) ||
    /本[處处署局]檔號/.test(top);
}

function incomingPageLooksLikeFollowingLetter(page: IOcrPageResult): boolean {
  return incomingPageLooksLikeNewLetter(page) || incomingPageLooksLikeStandaloneDocument(page);
}

function incomingPageHasOurRef(page: IOcrPageResult): boolean {
  const text = incomingEmailPlainText(page) || page.text || '';
  return /\bour\s+r+e+f(?:erence)?(?:\s*no(?:\.)?)?\s*[:：.-]/i.test(text) ||
    /本[處处署局]檔號\s*[:：.-]/.test(text) ||
    /本函編號\s*[:：.-]/.test(text);
}

function incomingPageLooksLikeNewLetter(page: IOcrPageResult): boolean {
  const top = incomingPageBandText(page, 0, 0.4);
  if (!top) {
    return false;
  }
  if (incomingPageHasEmailHeaderTokensInText(top)) {
    return false;
  }
  return /\bdear\s+(sir|madam|sirs|colleagues?)\b/i.test(top) ||
    /敬啟者|敬启者/.test(top) ||
    (/\bour\s+r+e+f/i.test(top) && /\bdate\b/i.test(top));
}

function incomingPageBottomLooksLikeEmail(page: IOcrPageResult): boolean {
  const bottom = incomingPageBandText(page, 0.4, 1);
  return incomingPageHasEmailHeaderTokensInText(bottom) || looksLikeIncomingEmail(bottom);
}

function incomingPageHasEmailHeaderTokens(page: IOcrPageResult): boolean {
  return incomingPageHasEmailHeaderTokensInText(incomingEmailPlainText(page) || page.text || '');
}

function incomingPageHasEmailHeaderTokensInText(text: string): boolean {
  const source = text || '';
  return /\bfrom\s*:/i.test(source) ||
    /\bsent\s*:/i.test(source) ||
    /\bsubject\s*:/i.test(source) ||
    /發件人|寄件人/.test(source) ||
    /主旨\s*[:：]/.test(source);
}

function incomingPageHasClosing(text: string): boolean {
  const key = ' ' + (text || '').toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim() + ' ';
  return key.indexOf(' yours sincere') >= 0 ||
    key.indexOf(' yours faithful') >= 0 ||
    key.indexOf(' yours falth') >= 0 ||
    key.indexOf(' yours truly') >= 0 ||
    key.indexOf(' your sincere') >= 0 ||
    key.indexOf(' your faithful') >= 0 ||
    key.indexOf(' your falth') >= 0 ||
    key.indexOf(' your truly') >= 0;
}

function isGovernmentEmailText(text: string): boolean {
  return /hong kong special administrative region|\bhksar\b|the government of/i.test(text) ||
    /civil engineering and development|highways department|drainage services|water supplies|architectural services|environmental protection department|lands department/i.test(text) ||
    /土木工程拓展署|路政署|渠務署|水務署|建築署|環境保護署|地政總署|香港特別行政區政府/.test(text);
}

function stripEmailMarkup(text: string): string {
  return (text || '')
    .replace(/<[^>]+>/g, ' ')
    .replace(/&nbsp;/gi, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}
