// Incoming AI prompt only. Outgoing lives in services/outgoing/aiPrompt.ts.
import { IAiExtractionHints } from '../correspondenceTypes';
import { OUTGOING_SUBJECT_PROMPT, outgoingBodyImageLabel } from '../outgoing/aiPrompt';

export function incomingSystemPrompt(): string {
  return [
    'You extract metadata from incoming correspondence sent TO AECOM, written by another organisation.',
    'Use OCR text and the first-page image. Reply with JSON only. Use empty strings when a value is not clearly present.',
    'Copy original wording from text when it is visible. OCR text may include <u>underlined</u> and <b>bold</b> tags; never copy those tags into values.',
    'Organization is the left-hand addressee line immediately above a floor line such as 12/F, or if there is no xx/F, immediately above a line containing xxx Road. The address is only on the left of the page; ignore Our Ref, Date, and other text on the right of the same row. Do not use the letterhead.',
    'Sender is exactly the text inside the parentheses immediately below Yours faithfully / Yours sincerely / Yours truly / 署名, for example (Ben xXx. LXX) -> Ben xXx. LXX. That closing is often on the last page, not page 1. Skip (signed). Do not copy CC / c.c. / 副本 names below the signature. Do not copy parentheses from Attn or the address on page 1. Do not copy the job title under the parentheses.',
    'Receiver is the Attn value when Attn / Attention is present. If there is no Attn, Receiver is the first left-hand address line above Dear / 敬啟者.',
    'Ref No is Our Ref / 本處檔號 / 檔號 of the originating party, not Your Ref.',
    'Project Number is the 8 digits immediately before the slash in Your Ref / 貴處檔號 / 來函編號; if there is no slash, the 8 digits immediately before the hyphen. Do not take Project Number from Our Ref.',
    'Do not invent values.'
  ].join(' ') + '\nSubject:\n' + OUTGOING_SUBJECT_PROMPT;
}

export function incomingLetterheadImageLabel(): string {
  return 'Letterhead at the top of the page. Organization is not taken from here; it is the line above xx/F or xxx Road in the addressee address:';
}

export function incomingBodyImageLabel(): string {
  return outgoingBodyImageLabel();
}

export function incomingClosingImageLabel(): string {
  return 'Closing / signature block. Sender is the text INSIDE the signature parentheses:';
}

export function buildIncomingUserPrompt(
  ocrText: string,
  fieldLabels: string[],
  hasImage: boolean,
  hints?: IAiExtractionHints
): string {
  const signature = hints && hints.signature;
  const receiverName = hints && hints.receiverName;
  const subjectText = hints && hints.subjectText;
  const refNo = hints && hints.refNo;
  const yourRef = hints && hints.yourRef;
  const organization = hints && hints.organization;
  const letterType = hints && hints.letterType;

  const sourceLines = hasImage
    ? [
      'Extract these fields from an incoming letter sent TO AECOM.',
      'Images: full first page, letterhead at the top, the body after the salutation, then the signature block if detected.',
      'Organization is the LEFT-HAND line above xx/F, or if there is no floor line, the line above xxx Road. Not the whole row and not the letterhead. Receiver is the Attn value if present; if there is no Attn, it is the first left-hand address line above Dear. Ignore Our Ref / Date on the right. Subject: after Dear Sir/Madam, copy the entire first line that has both <b> and <u> on the same words. Do not copy a line that only sits above a table or divider. Sender is the text inside the signature parentheses.'
    ]
    : [
      'Extract these fields from the OCR text of an incoming letter sent TO AECOM.',
      'Organization is the LEFT-HAND line above xx/F, or if there is no floor line, the line above xxx Road, above Dear / 敬啟者. Not the whole row and not the letterhead. Ignore right-column Our Ref / Date. Sender is the text inside the signature parentheses. Receiver is the Attn value if present; if there is no Attn, the first left-hand address line above Dear. Subject: after Dear Sir/Madam, copy the entire first line that has both <b> and <u> on the same words. Do not copy a line that only sits above a table or divider.'
    ];

  const parts = [
    sourceLines.join(' '),
    'Return a JSON object whose keys are exactly:',
    fieldLabels.join(', '),
    '',
    'Field meanings:',
    incomingFieldHelp(),
    '',
    'OCR text (plain words plus <u>underline</u> and <b>bold</b> tags when detected):',
    ocrText || '(none)'
  ];
  if (letterType && letterType !== 'unknown') {
    parts.push('', 'Detected incoming letter type:', letterType);
  }
  if (signature && signature.senderName) {
    parts.push('', 'Detected Sender from the parentheses immediately below Yours sincerely / Yours faithfully, not from CC:', signature.senderName);
  } else if (signature && signature.textBelow) {
    parts.push('', 'OCR immediately below the signature. Sender is the text INSIDE parentheses immediately below Yours sincerely / Yours faithfully. Do not copy CC / c.c. / 副本:', signature.textBelow);
  } else {
    parts.push('', 'Copy the text inside the signature parentheses below Yours faithfully / Yours sincerely / Yours truly / 署名. Skip (signed). For a memo, use From:.');
  }
  if (receiverName && receiverName.trim()) {
    parts.push('', 'Detected Receiver from Attn, or if there is no Attn, from the first left-hand address line above Dear:', receiverName.trim());
  } else {
    parts.push('', 'Receiver is the value after Attn: / Attention: if present. If there is no Attn, copy the first left-hand address line above Dear / 敬啟者. Do not use the line below Yours faithfully.');
  }
  if (organization && organization.trim()) {
    parts.push('', 'Detected Organization from the LEFT-HAND line above xx/F or xxx Road in the addressee address (not the whole row, not the letterhead):', organization.trim());
  } else {
    parts.push('', 'Organization is the left-hand line immediately above xx/F, or if there is no xx/F, the line above xxx Road in the addressee address. Do not copy right-column text from that row, and do not use the letterhead.');
  }
  parts.push('', OUTGOING_SUBJECT_PROMPT);
  if (subjectText && subjectText.trim()) {
    parts.push('', 'Detected Subject from nearby lines after Dear that have both bold and underline. Use this text:', subjectText.trim());
  }
  if (refNo && refNo.trim()) {
    parts.push('', 'Detected Ref No from Our Ref: or standalone Ref:', refNo.trim());
  } else {
    parts.push('', 'Find Our Ref: / 本處檔號 / 檔號 first. Ref No is the value to the right of that label, not Your Ref.');
  }
  if (yourRef && yourRef.trim()) {
    parts.push('', 'Detected Your Ref. Project Number is the 8 digits before / or - in this value:', yourRef.trim());
  } else {
    parts.push('', 'Find Your Ref: / 貴處檔號 / 來函編號. Project Number is the 8 digits immediately before / or - in that value, not Our Ref.');
  }
  return parts.join('\n');
}

function incomingFieldHelp(): string {
  return [
    'Name: document name or identifier. Registration Number must use this same value.',
    'Registration Number: always copy Name exactly. Do not invent a different value.',
    'Leading BL: leave empty unless an AECOM business-line name is clearly shown. Must be one of: Architecture, Building Engineering, Environment, Geotechnical, Digital, Land Supply and Municipal, MEP, Project and Construction Management, Program, Cost and Consultancy, Transportation, Unclassified, Urbanism and Planning, Water.',
    'Project Number: 8 digits immediately before the slash in Your Ref: / Your Ref : / 貴處檔號 / 來函編號. Example Your Ref: 12345678/ABC -> 12345678. If there is no slash, take the 8 digits immediately before the hyphen. Digits only. Do not use Our Ref.',
    'Sub-Project Number: dropdown value None, or an integer from 1 to 99. Use None when it is not shown.',
    'Organization: in the LEFT addressee address above Dear / 敬啟者, find a floor line such as 12/F or G/F and copy only the left-hand line immediately above it. If there is no xx/F, find a line containing xxx Road and copy the left-hand line immediately above that. Do not include Our Ref, Your Ref, Date, or other text on the right of that row. Do not use letterhead.',
    'Sender: copy ONLY the text inside the parentheses immediately below Yours faithfully / Yours sincerely / Yours truly / 署名, e.g. (Ben xXx. LXX) -> Ben xXx. LXX. Look on the page that contains that closing, which is often the last page, not the first page. Skip (signed). Do not copy CC / c.c. / 副本 / copy to names, whether they sit below the signature. Do not copy parentheses from Attn: near the top of page 1. Do not copy the job title under the parentheses. For a memo or email with no signature parentheses, use From:.',
    'Receiver: if Attn: / Attn : / Attention: is present, copy only the value after that label. If there is no Attn, copy the first left-hand address line above Dear / 敬啟者. Do not use the line below Yours faithfully. Ignore Our Ref / Date on the right of the same row.',
    'Subject:\n' + OUTGOING_SUBJECT_PROMPT,
    'File No: file number if shown separately from Ref No.',
    'Ref No: copy the value to the right of Our Ref: / 本處檔號 / 檔號. OCR may read Ref as Rref or Reef. Do not use Your Ref.',
    'Issue Date: the document date in dd/MM/yyyy, for example 03/09/2026. Accept 8 September 2026 and 2026年9月8日.',
    'Attachment: Yes if the letter mentions attachments, enclosures, 附件, or 隨函, otherwise No. Reply only Yes or No.',
    'Scan: Yes if the document is a scan or has a scan mark, otherwise No. Reply only Yes or No.',
    'Remark: remarks or notes',
    'Location: office, site, or city',
    'cc to AECOM: Yes if the CC list includes AECOM people beyond the addressee, otherwise No. Reply only Yes or No.'
  ].join('\n');
}
