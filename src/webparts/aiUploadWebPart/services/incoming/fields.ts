// Incoming Convert rules only. Outgoing lives in services/outgoing/.
import { isOrganizationField, isReceiverField, isRefNoField, isSenderField, isSubjectField } from '../../constants/defaultFormFields';
import { isIssueDateField } from '../../constants/issueDate';
import { isProjectNumberField, sanitizeProjectNumber } from '../../constants/projectNumber';
import { IAiExtractionHints, IDetectedFields, emptyDetectedFields, tryText, tryTextAsync } from '../correspondenceTypes';
import { extractOurRefNo, extractYourRefNo } from '../fieldExtractor';
import { IOcrPageResult } from '../IPdfOcr';
import { analyzeDocumentSignature, extractSubjectBelowDearSir, subjectAppearsInPage } from '../signatureSender';
import {
  classifyIncomingLetter,
  extractIncomingIssueDate,
  extractIncomingMemoSender,
  extractIncomingOrganization,
  extractIncomingReceiver,
  extractIncomingSender,
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
  detected.receiverName = tryText(() => extractIncomingReceiver(list));
  detected.organization = tryText(() => extractIncomingOrganization(firstPage));
  detected.subjectText = await tryTextAsync(() => extractSubjectBelowDearSir(firstPage));
  detected.refNo = tryText(() => extractOurRefNo(list));
  detected.yourRef = tryText(() => extractYourRefNo(list));
  detected.projectNumber = tryText(() => incomingProjectNumber(list));
  detected.issueDate = tryText(() => extractIncomingIssueDate(list));
  detected.memoSender = tryText(() => extractIncomingMemoSender(firstPage));
  const closingSender = tryText(() => extractIncomingSender(list));
  const inkParenSender = incomingSignatureParenName(detected.signature.textBelow);
  detected.signature = {
    ...detected.signature,
    senderName: closingSender || inkParenSender || incomingSenderName(detected.memoSender)
  };
  return detected;
}

export function pickIncomingFieldValue(
  label: string,
  detected: IDetectedFields,
  aiValue: string,
  keywordValue: string
): string {
  if (isSenderField(label)) {
    return incomingSenderName(detected.signature.senderName);
  }
  if (isReceiverField(label)) {
    return detected.receiverName;
  }
  if (isSubjectField(label)) {
    return detected.subjectText || groundedSubject(detected.firstPage, aiValue);
  }
  if (isRefNoField(label)) {
    return detected.refNo || (aiValue || '').trim();
  }
  if (isProjectNumberField(label)) {
    return detected.projectNumber || sanitizeProjectNumber(aiValue || '') || keywordValue || '';
  }
  if (isOrganizationField(label)) {
    return detected.organization;
  }
  if (isIssueDateField(label)) {
    return detected.issueDate || (aiValue && aiValue.trim()) || keywordValue || '';
  }
  return (aiValue && aiValue.trim()) || keywordValue || '';
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
