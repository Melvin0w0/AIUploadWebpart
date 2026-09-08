// Outgoing Convert rules only. Incoming lives in services/incoming/.
import { isOrganizationField, isReceiverField, isRefNoField, isSenderField, isSubjectField } from '../../constants/defaultFormFields';
import { isProjectNumberField, projectNumberFromRef, sanitizeProjectNumber } from '../../constants/projectNumber';
import { IAiExtractionHints, IDetectedFields, emptyDetectedFields, tryText, tryTextAsync } from '../correspondenceTypes';
import { extractOurRefNo, extractOurRefOnly } from '../fieldExtractor';
import { IOcrPageResult } from '../IPdfOcr';
import { analyzeDocumentSignature, asPersonName, extractOrganizationAboveAddressee, extractReceiverAboveDearSir, extractSubjectBelowDearSir, subjectAppearsInPage } from '../signatureSender';

export async function detectOutgoingFields(pages: IOcrPageResult[]): Promise<IDetectedFields> {
  const list = pages || [];
  const firstPage = list[0];
  const detected = emptyDetectedFields(firstPage);
  detected.signature = await analyzeDocumentSignature(list);
  detected.receiverName = tryText(() => extractReceiverAboveDearSir(firstPage));
  detected.organization = tryText(() => extractOrganizationAboveAddressee(firstPage));
  detected.subjectText = await tryTextAsync(() => extractSubjectBelowDearSir(firstPage));
  detected.refNo = tryText(() => extractOurRefNo(list));
  detected.projectNumber = tryText(() => projectNumberFromRef(extractOurRefOnly(list)));
  return detected;
}

export function pickOutgoingFieldValue(
  label: string,
  detected: IDetectedFields,
  aiValue: string,
  keywordValue: string
): string {
  if (isSenderField(label)) {
    return detected.signature.senderName || asPersonName(aiValue || '');
  }
  if (isReceiverField(label)) {
    return detected.receiverName || firstAiLine(aiValue);
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
    const aiOrganization = (aiValue || '').trim().split(/\r?\n/)[0].trim();
    return detected.organization || aiOrganization || keywordValue || '';
  }
  return (aiValue && aiValue.trim()) || keywordValue || '';
}

export function outgoingAiHints(detected: IDetectedFields): IAiExtractionHints {
  return {
    page: detected.firstPage,
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
