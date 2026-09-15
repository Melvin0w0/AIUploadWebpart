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

export async function detectIncomingFields(pages: IOcrPageResult[]): Promise<IDetectedFields> {
  const list = pages || [];
  const looksEmail = incomingPagesLookLikeEmail(list);
  const workPages = looksEmail ? incomingLastEmailPages(list) : list;
  const firstPage = workPages[0] || list[0];
  const detected = emptyDetectedFields(firstPage);
  const classification = classifyIncomingLetter(firstPage);
  detected.letterType = looksEmail ? 'email' : classification.letterType;
  const email = detected.letterType === 'email' ? extractIncomingEmailHeaders(list) : undefined;
  detected.signature = await analyzeDocumentSignature(workPages);
  if (looksEmail) {
    const emailTo = extractIncomingEmailToFromLastPage(list);
    detected.receiverName = emailTo;
    detected.sources.receiver = emailTo ? 'Email To' : '';
  } else {
    const receiver = tryPicked(() => extractIncomingReceiverLocated(list));
    detected.receiverName = receiver.value;
    detected.sources.receiver = receiver.source;
  }
  const organization = tryPicked(() => extractIncomingOrganizationLocated(firstPage));
  detected.organization = organization.value;
  detected.sources.organization = organization.source;
  if (looksEmail) {
    const emailSubject = extractIncomingEmailSubjectFromLastPage(list);
    detected.subjectText = emailSubject;
    detected.sources.subject = emailSubject ? 'Email Subject' : '';
  } else {
    detected.subjectText = tryText(() => extractIncomingSubject(firstPage));
    if (!detected.subjectText) {
      detected.subjectText = await tryTextAsync(() => extractSubjectBelowDearSir(firstPage));
    }
    if (detected.subjectText) {
      detected.sources.subject = 'Dear 後粗體+底線';
    }
  }
  detected.refNo = looksEmail
    ? (extractIncomingEmailOurRef(list) || tryText(() => extractOurRefOnly(list)))
    : tryText(() => extractOurRefNo(list));
  if (detected.refNo) {
    detected.sources.refNo = 'Our Ref';
  }
  detected.yourRef = tryText(() => extractYourRefNo(looksEmail ? workPages : list));
  if (detected.yourRef) {
    detected.sources.yourRef = 'Your Ref';
  }
  detected.agreementNo = (email && email.agreementNo) || tryText(() => extractIncomingAgreementNo(firstPage));
  if (detected.agreementNo) {
    detected.sources.agreementNo = 'Agreement / Contract No.';
  }
  detected.issueDate = (email && email.sent) || tryText(() => extractIncomingIssueDate(looksEmail ? workPages : list));
  if (detected.issueDate) {
    detected.sources.issueDate = email && email.sent ? 'Email Sent' : 'Date';
  }
  detected.memoSender = tryText(() => extractIncomingMemoSender(firstPage));
  const closingSender = tryPicked(() => extractIncomingSenderLocated(looksEmail ? workPages : list));
  const inkParenSender = incomingSignatureParenName(detected.signature.textBelow);
  const memoSender = incomingSenderName(detected.memoSender);
  let senderName = '';
  let senderSource = '';
  if (looksEmail) {
    const ccSender = extractIncomingEmailCcFromLastPage(list);
    if (ccSender) {
      senderName = ccSender;
      senderSource = 'Email CC';
    }
  } else if (closingSender.value) {
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

export function pickIncomingFieldValue(
  label: string,
  detected: IDetectedFields,
  aiValue: string,
  keywordValue: string
): IPickedValue {
  if (isSenderField(label)) {
    if (detected.letterType === 'email') {
      if (detected.signature.senderName) {
        return picked(detected.signature.senderName, detected.sources.sender);
      }
      if (aiValue && aiValue.trim()) {
        return picked(aiValue.trim(), 'AI');
      }
      return picked(keywordValue || '', keywordValue ? '關鍵字' : '');
    }
    return picked(incomingSenderName(detected.signature.senderName), detected.sources.sender);
  }
  if (isReceiverField(label)) {
    if (detected.letterType === 'email') {
      if (detected.receiverName) {
        return picked(detected.receiverName, detected.sources.receiver);
      }
      if (aiValue && aiValue.trim()) {
        return picked(aiValue.trim(), 'AI');
      }
      return picked(keywordValue || '', keywordValue ? '關鍵字' : '');
    }
    return picked(detected.receiverName, detected.sources.receiver);
  }
  if (isSubjectField(label)) {
    if (detected.subjectText) {
      return picked(detected.subjectText, detected.sources.subject || 'Dear 後粗體+底線');
    }
    if (detected.letterType === 'email') {
      if (aiValue && aiValue.trim()) {
        return picked(aiValue.trim(), 'AI');
      }
      return picked(keywordValue || '', keywordValue ? '關鍵字' : '');
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
