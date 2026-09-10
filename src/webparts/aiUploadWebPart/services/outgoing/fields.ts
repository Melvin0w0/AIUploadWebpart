// Outgoing Convert rules only. Incoming lives in services/incoming/.
import { isOrganizationField, isReceiverField, isRefNoField, isSenderField, isSubjectField } from '../../constants/defaultFormFields';
import { isProjectNumberField, projectNumberFromRef, sanitizeProjectNumber } from '../../constants/projectNumber';
import { IAiExtractionHints, IDetectedFields, IPickedValue, emptyDetectedFields, picked, tryText, tryTextAsync } from '../correspondenceTypes';
import { extractOurRefNo, extractOurRefOnly } from '../fieldExtractor';
import { IOcrPageResult } from '../IPdfOcr';
import { analyzeDocumentSignature, extractOrganizationAboveAddressee, extractOutgoingSenderFromPages, extractReceiverAboveDearSir, extractSubjectBelowDearSir, subjectAppearsInPage } from '../signatureSender';

export async function detectOutgoingFields(pages: IOcrPageResult[]): Promise<IDetectedFields> {
  const list = pages || [];
  const firstPage = list[0];
  const closingPage = list.length > 0 ? list[list.length - 1] : undefined;
  const signature = await analyzeDocumentSignature(list);
  const subjectText = await tryTextAsync(() => extractSubjectBelowDearSir(firstPage));
  const outgoingSender = tryText(() => extractOutgoingSenderFromPages(list, signature.region, signature.regionPageNumber));
  const detected = emptyDetectedFields(firstPage);
  detected.closingPage = closingPage;
  detected.signature = {
    ...signature,
    senderName: outgoingSender
  };
  detected.receiverName = tryText(() => extractReceiverAboveDearSir(firstPage));
  if (detected.receiverName) {
    detected.sources.receiver = 'Attn / Dear 上方';
  }
  detected.organization = tryText(() => extractOrganizationAboveAddressee(firstPage));
  if (detected.organization) {
    detected.sources.organization = 'Department 行';
  }
  detected.subjectText = subjectText;
  if (subjectText) {
    detected.sources.subject = 'Dear 後粗體+底線';
  }
  detected.refNo = tryText(() => extractOurRefNo(list));
  if (detected.refNo) {
    detected.sources.refNo = 'Our Ref';
  }
  detected.projectNumber = tryText(() => projectNumberFromRef(extractOurRefOnly(list)));
  if (detected.projectNumber) {
    detected.sources.projectNumber = 'Our Ref';
  }
  if (outgoingSender) {
    detected.sources.sender = '職稱上一行';
  }
  return detected;
}

export function pickOutgoingFieldValue(
  label: string,
  detected: IDetectedFields,
  aiValue: string,
  keywordValue: string
): IPickedValue {
  if (isSenderField(label)) {
    if (detected.signature.senderName) {
      return picked(detected.signature.senderName, detected.sources.sender || '職稱上一行');
    }
    return picked(firstAiLine(aiValue), 'AI');
  }
  if (isReceiverField(label)) {
    if (detected.receiverName) {
      return picked(detected.receiverName, detected.sources.receiver || 'Attn / Dear 上方');
    }
    return picked(firstAiLine(aiValue), 'AI');
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
      return picked(detected.projectNumber, detected.sources.projectNumber || 'Our Ref');
    }
    const fromAi = sanitizeProjectNumber(aiValue || '');
    if (fromAi) {
      return picked(fromAi, 'AI');
    }
    return picked(keywordValue || '', '關鍵字');
  }
  if (isOrganizationField(label)) {
    const aiOrganization = (aiValue || '').trim().split(/\r?\n/)[0].trim();
    if (detected.organization) {
      return picked(detected.organization, detected.sources.organization || 'Department 行');
    }
    if (aiOrganization) {
      return picked(aiOrganization, 'AI');
    }
    return picked(keywordValue || '', '關鍵字');
  }
  if (aiValue && aiValue.trim()) {
    return picked(aiValue.trim(), 'AI');
  }
  return picked(keywordValue || '', '關鍵字');
}

export function outgoingAiHints(detected: IDetectedFields): IAiExtractionHints {
  return {
    page: detected.firstPage,
    closingPage: detected.closingPage,
    signature: detected.signature,
    receiverName: detected.receiverName,
    subjectText: detected.subjectText,
    refNo: detected.refNo,
    organization: detected.organization,
    kind: 'outgoing'
  };
}

function firstAiLine(aiValue: string): string {
  const aiReceiver = (aiValue || '').trim();
  if (!aiReceiver || /^dear\b/i.test(aiReceiver)) {
    return '';
  }
  return aiReceiver.split(/\r?\n/)[0].trim();
}

function groundedSubject(page: IOcrPageResult | undefined, aiValue: string): string {
  const aiSubject = (aiValue || '').trim();
  if (!page || !aiSubject || !subjectAppearsInPage(page, aiSubject)) {
    return '';
  }
  return aiSubject;
}
