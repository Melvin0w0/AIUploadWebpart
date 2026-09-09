// Incoming Convert rules only. Outgoing lives in services/outgoing/.
import { isOrganizationField, isReceiverField, isRefNoField, isSenderField, isSubjectField } from '../../constants/defaultFormFields';
import { isIssueDateField } from '../../constants/issueDate';
import { isProjectNumberField, sanitizeProjectNumber } from '../../constants/projectNumber';
import { IAiExtractionHints, IDetectedFields, IPickedValue, emptyDetectedFields, picked, tryPicked, tryText, tryTextAsync } from '../correspondenceTypes';
import { extractOurRefNo, extractYourRefNo } from '../fieldExtractor';
import { IOcrPageResult } from '../IPdfOcr';
import { analyzeDocumentSignature, extractSubjectBelowDearSir, subjectAppearsInPage } from '../signatureSender';
import {
  classifyIncomingLetter,
  extractIncomingIssueDate,
  extractIncomingMemoSender,
  extractIncomingOrganizationLocated,
  extractIncomingReceiverLocated,
  extractIncomingSenderLocated,
  extractIncomingSubject,
  incomingProjectNumber,
  incomingSenderName,
  incomingSignatureParenName
} from './extract';

export async function detectIncomingFields(pages: IOcrPageResult[]): Promise<IDetectedFields> {
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
  detected.projectNumber = tryText(() => incomingProjectNumber(list));
  if (detected.projectNumber) {
    detected.sources.projectNumber = 'Your Ref';
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

export function pickIncomingFieldValue(
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
    if (detected.projectNumber) {
      return picked(detected.projectNumber, detected.sources.projectNumber || 'Your Ref');
    }
    const fromAi = sanitizeProjectNumber(aiValue || '');
    if (fromAi) {
      return picked(fromAi, 'AI');
    }
    return picked(keywordValue || '', '關鍵字');
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
