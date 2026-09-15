// Incoming Convert rules only. Outgoing lives in services/outgoing/.
import { isOrganizationField, isReceiverField, isRefNoField, isSenderField, isSubjectField } from '../../constants/defaultFormFields';
import { isIssueDateField } from '../../constants/issueDate';
import { isProjectNumberField } from '../../constants/projectNumber';
import { IAiExtractionHints, IDetectedFields, IPickedValue, emptyDetectedFields, picked, tryPicked, tryText, tryTextAsync } from '../correspondenceTypes';
import { extractOurRefNo, extractOurRefOnly, extractYourRefNo } from '../fieldExtractor';
import { IOcrPageResult } from '../IPdfOcr';
import { analyzeDocumentSignature, extractSubjectBelowDearSir, subjectAppearsInPage } from '../signatureSender';
import {
  extractIncomingEmailCcFromLastPage,
  extractIncomingEmailHeaders,
  extractIncomingEmailOurRef,
  extractIncomingEmailSubjectFromLastPage,
  extractIncomingEmailToFromLastPage,
  incomingLastEmailPages,
  incomingPagesLookLikeEmail
} from './email';
import {
  classifyIncomingLetter,
  extractIncomingAgreementNo,
  extractIncomingIssueDate,
  extractIncomingMemoSender,
  extractIncomingOrganizationLocated,
  extractIncomingReceiverLocated,
  extractIncomingSenderLocated,
  extractIncomingSubject,
  incomingSenderName,
  incomingSignatureParenName
} from './extract';

export function isIncomingEmailFormat(pages?: IOcrPageResult[]): boolean {
  return incomingPagesLookLikeEmail(pages);
}

export async function detectIncomingFields(pages: IOcrPageResult[]): Promise<IDetectedFields> {
  if (isIncomingEmailFormat(pages)) {
    return detectIncomingEmailFields(pages);
  }
  return detectIncomingLetterFields(pages);
}

async function detectIncomingLetterFields(pages: IOcrPageResult[]): Promise<IDetectedFields> {
  const list = pages || [];
  const firstPage = list[0];
  const detected = emptyDetectedFields(firstPage);
  const classification = classifyIncomingLetter(firstPage);
  detected.letterType = classification.letterType;
  detected.signature = await analyzeDocumentSignature(list);
  const receiver = tryPicked(() => extractIncomingReceiverLocated(list));
  detected.receiverName = receiver.value;
  detected.sources.receiver = receiver.source;
  const organization = tryPicked(() => extractIncomingOrganizationLocated(firstPage));
  detected.organization = organization.value;
  detected.sources.organization = organization.source;
  detected.subjectText = tryText(() => extractIncomingSubject(firstPage));
  if (!detected.subjectText) {
    detected.subjectText = await tryTextAsync(() => extractSubjectBelowDearSir(firstPage));
  }
  if (detected.subjectText) {
    detected.sources.subject = 'Dear 後粗體+底線';
  }
  detected.refNo = tryText(() => extractOurRefNo(list));
  if (detected.refNo) {
    detected.sources.refNo = 'Our Ref';
  }
  detected.yourRef = tryText(() => extractYourRefNo(list));
  if (detected.yourRef) {
    detected.sources.yourRef = 'Your Ref';
  }
  detected.agreementNo = tryText(() => extractIncomingAgreementNo(firstPage));
  if (detected.agreementNo) {
    detected.sources.agreementNo = 'Agreement / Contract No.';
  }
  detected.issueDate = tryText(() => extractIncomingIssueDate(list));
  if (detected.issueDate) {
    detected.sources.issueDate = 'Date';
  }
  detected.memoSender = tryText(() => extractIncomingMemoSender(firstPage));
  const closingSender = tryPicked(() => extractIncomingSenderLocated(list));
  const inkParenSender = incomingSignatureParenName(detected.signature.textBelow);
  const memoSender = incomingSenderName(detected.memoSender);
  let senderName = '';
  let senderSource = '';
  if (closingSender.value) {
    senderName = closingSender.value;
    senderSource = closingSender.source;
  } else if (inkParenSender) {
    senderName = inkParenSender;
    senderSource = '簽署旁括號';
  } else if (memoSender) {
    senderName = memoSender;
    senderSource = 'Memo From';
  }
  detected.signature = {
    ...detected.signature,
    senderName
  };
  detected.sources.sender = senderSource;
  return detected;
}

async function detectIncomingEmailFields(pages: IOcrPageResult[]): Promise<IDetectedFields> {
  const list = pages || [];
  const workPages = incomingLastEmailPages(list);
  const firstPage = workPages[0] || list[0];
  const detected = emptyDetectedFields(firstPage);
  detected.letterType = 'email';
  const email = extractIncomingEmailHeaders(list);
  detected.signature = await analyzeDocumentSignature(workPages);
  const emailTo = extractIncomingEmailToFromLastPage(list);
  detected.receiverName = emailTo;
  detected.sources.receiver = emailTo ? 'Email To' : '';
  const organization = tryPicked(() => extractIncomingOrganizationLocated(firstPage));
  detected.organization = organization.value;
  detected.sources.organization = organization.source;
  const emailSubject = extractIncomingEmailSubjectFromLastPage(list);
  detected.subjectText = emailSubject;
  detected.sources.subject = emailSubject ? 'Email Subject' : '';
  detected.refNo = extractIncomingEmailOurRef(list) || tryText(() => extractOurRefOnly(list));
  if (detected.refNo) {
    detected.sources.refNo = 'Our Ref';
  }
  detected.yourRef = tryText(() => extractYourRefNo(workPages));
  if (detected.yourRef) {
    detected.sources.yourRef = 'Your Ref';
  }
  detected.agreementNo = email.agreementNo || tryText(() => extractIncomingAgreementNo(firstPage));
  if (detected.agreementNo) {
    detected.sources.agreementNo = 'Agreement / Contract No.';
  }
  detected.issueDate = email.sent || tryText(() => extractIncomingIssueDate(workPages));
  if (detected.issueDate) {
    detected.sources.issueDate = email.sent ? 'Email Sent' : 'Date';
  }
  const ccSender = extractIncomingEmailCcFromLastPage(list);
  detected.signature = {
    ...detected.signature,
    senderName: ccSender
  };
  detected.sources.sender = ccSender ? 'Email CC' : '';
  return detected;
}

export function pickIncomingFieldValue(
  label: string,
  detected: IDetectedFields,
  aiValue: string,
  keywordValue: string
): IPickedValue {
  if (detected.letterType === 'email') {
    return pickIncomingEmailFieldValue(label, detected, aiValue, keywordValue);
  }
  return pickIncomingLetterFieldValue(label, detected, aiValue, keywordValue);
}

function pickIncomingLetterFieldValue(
  label: string,
  detected: IDetectedFields,
  aiValue: string,
  keywordValue: string
): IPickedValue {
  if (isSenderField(label)) {
    return picked(incomingSenderName(detected.signature.senderName), detected.sources.sender);
  }
  if (isReceiverField(label)) {
    return picked(detected.receiverName, detected.sources.receiver);
  }
  if (isSubjectField(label)) {
    if (detected.subjectText) {
      return picked(detected.subjectText, detected.sources.subject || 'Dear 後粗體+底線');
    }
    return picked(groundedSubject(detected.firstPage, aiValue), 'AI');
  }
  if (isRefNoField(label)) {
    if (detected.refNo) {
      return picked(detected.refNo, detected.sources.refNo || 'Our Ref');
    }
    return picked((aiValue || '').trim(), 'AI');
  }
  if (isProjectNumberField(label)) {
    return picked('', '');
  }
  if (isOrganizationField(label)) {
    return picked(detected.organization, detected.sources.organization);
  }
  if (isIssueDateField(label)) {
    if (detected.issueDate) {
      return picked(detected.issueDate, detected.sources.issueDate || 'Date');
    }
    if (aiValue && aiValue.trim()) {
      return picked(aiValue.trim(), 'AI');
    }
    return picked(keywordValue || '', '關鍵字');
  }
  if (aiValue && aiValue.trim()) {
    return picked(aiValue.trim(), 'AI');
  }
  return picked(keywordValue || '', '關鍵字');
}

function pickIncomingEmailFieldValue(
  label: string,
  detected: IDetectedFields,
  aiValue: string,
  keywordValue: string
): IPickedValue {
  if (isSenderField(label)) {
    if (detected.signature.senderName) {
      return picked(detected.signature.senderName, detected.sources.sender);
    }
    if (aiValue && aiValue.trim()) {
      return picked(aiValue.trim(), 'AI');
    }
    return picked(keywordValue || '', keywordValue ? '關鍵字' : '');
  }
  if (isReceiverField(label)) {
    if (detected.receiverName) {
      return picked(detected.receiverName, detected.sources.receiver);
    }
    if (aiValue && aiValue.trim()) {
      return picked(aiValue.trim(), 'AI');
    }
    return picked(keywordValue || '', keywordValue ? '關鍵字' : '');
  }
  if (isSubjectField(label)) {
    if (detected.subjectText) {
      return picked(detected.subjectText, detected.sources.subject || 'Email Subject');
    }
    if (aiValue && aiValue.trim()) {
      return picked(aiValue.trim(), 'AI');
    }
    return picked(keywordValue || '', keywordValue ? '關鍵字' : '');
  }
  if (isRefNoField(label)) {
    if (detected.refNo) {
      return picked(detected.refNo, detected.sources.refNo || 'Our Ref');
    }
    return picked((aiValue || '').trim(), 'AI');
  }
  if (isProjectNumberField(label)) {
    return picked('', '');
  }
  if (isOrganizationField(label)) {
    return picked(detected.organization, detected.sources.organization);
  }
  if (isIssueDateField(label)) {
    if (detected.issueDate) {
      return picked(detected.issueDate, detected.sources.issueDate || 'Email Sent');
    }
    if (aiValue && aiValue.trim()) {
      return picked(aiValue.trim(), 'AI');
    }
    return picked(keywordValue || '', '關鍵字');
  }
  if (aiValue && aiValue.trim()) {
    return picked(aiValue.trim(), 'AI');
  }
  return picked(keywordValue || '', '關鍵字');
}

export function incomingAiHints(detected: IDetectedFields): IAiExtractionHints {
  return {
    page: detected.firstPage,
    signature: detected.signature,
    receiverName: detected.receiverName,
    subjectText: detected.subjectText,
    refNo: detected.refNo,
    yourRef: detected.yourRef,
    agreementNo: detected.agreementNo,
    organization: detected.organization,
    kind: 'incoming',
    letterType: detected.letterType
  };
}

function groundedSubject(page: IOcrPageResult | undefined, aiValue: string): string {
  const aiSubject = (aiValue || '').trim();
  if (!page || !aiSubject || !subjectAppearsInPage(page, aiSubject)) {
    return '';
  }
  return aiSubject;
}
