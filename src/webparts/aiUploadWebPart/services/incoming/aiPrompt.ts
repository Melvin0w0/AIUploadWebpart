// Incoming AI prompt only. Outgoing lives in services/outgoing/aiPrompt.ts.
import { IAiExtractionHints } from '../correspondenceTypes';

export function incomingSubjectPrompt(): string {
  return [
    'Extract Subject from the letter body after the salutation (Dear Sir / Dear Madam / 敬啟者) and before the closing (Yours faithfully / 此致).',
    'Prefer Re: / Subject: / 主旨 / 事由 / 關於 if present.',
    'If those labels are missing, copy a heading after the salutation. Bold+underline is helpful but not required.',
    'Do not include the salutation or the closing. Do not copy <u>, </u>, <b>, or </b> tags.'
  ].join('\n');
}

export function incomingSystemPrompt(): string {
  return [
    'You extract metadata from incoming correspondence sent TO AECOM, written by another organisation.',
    'Use OCR text and the first-page image. Reply with JSON only. Use empty strings when a value is not clearly present.',
    'Copy original wording from text when it is visible. OCR text may include <u>underlined</u> and <b>bold</b> tags; never copy those tags into values.',
    'Organization is the sender organisation on the letterhead at the top of the page, never AECOM in the addressee block.',
    'Sender is the printed person name immediately below the handwritten signature of the originating party, not the job title. Skip for Director of / Chief Engineer.',
    'Receiver is the AECOM person after Attn: / Attention: if present. Omit Mr., Ms., Mrs., Miss, and any parenthetical text.',
    'Ref No is Our Ref / 本處檔號 / 檔號 of the originating party, not Your Ref.',
    'Project Number is the 8 digits immediately before the slash in Your Ref / 貴處檔號 / 來函編號; if there is no slash, the 8 digits immediately before the hyphen. Do not take Project Number from Our Ref.',
    'Subject: after Dear / 敬啟者 and before the closing, copy Re: / Subject: / 主旨 / 事由, or a heading. Bold+underline if present, but it is not required.',
    'Do not invent values.'
  ].join(' ');
}

export function incomingLetterheadImageLabel(): string {
  return 'Letterhead at the top. Organization is the sender organisation printed here, not AECOM in the addressee block:';
}

export function incomingBodyImageLabel(): string {
  return 'Letter body AFTER the salutation and BEFORE the closing. Subject: Re:/Subject:/主旨/事由, or a heading. Bold+underline if present but not required:';
}

export function incomingClosingImageLabel(): string {
  return 'Closing block. Sender is the printed person name immediately below the handwritten signature:';
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
      'Organization is the sender on the letterhead, not AECOM. Receiver is the AECOM Attn. Subject may be Re:/主旨, not only bold+underline. Sender is the printed name below the signature.'
    ]
    : [
      'Extract these fields from the OCR text of an incoming letter sent TO AECOM.',
      'Organization is the letterhead sender, not AECOM. Sender is the person name below the signature. Receiver is Attn (AECOM). Subject: Re:/Subject:/主旨/事由, or a heading after Dear/敬啟者.'
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
    parts.push('', 'Detected Sender from the printed name below the signature:', signature.senderName);
  } else if (signature && signature.textBelow) {
    parts.push('', 'OCR immediately below the signature:', signature.textBelow);
  } else {
    parts.push('', 'Look below the handwritten signature, or From: on a memo/email, for the printed person name. That is Sender.');
  }
  if (receiverName && receiverName.trim()) {
    parts.push('', 'Detected Receiver:', receiverName.trim());
  } else {
    parts.push('', 'Receiver is the AECOM person after Attn:. For a memo/email, use To:. Omit Mr./Ms. and parenthetical text.');
  }
  if (organization && organization.trim()) {
    parts.push('', 'Detected Organization:', organization.trim());
  } else {
    parts.push('', 'Organization is the sender organisation on the letterhead above Our Ref. Do not use AECOM from the addressee block.');
  }
  parts.push('', incomingSubjectPrompt());
  if (subjectText && subjectText.trim()) {
    parts.push('', 'Detected Subject. Use this text unless the page clearly shows a better heading:', subjectText.trim());
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
    'Organization: the sender organisation printed in the letterhead ABOVE Our Ref. Government department, company, or consultant name. Never copy AECOM from the addressee block.',
    'Sender: printed person name immediately below the handwritten signature. Do not copy Chief Engineer, Director, Manager, or for Director of. Skip (signed). For a memo or email, use From:.',
    'Receiver: AECOM recipient. If Attn: / Attention: is present, copy only the person name after that label. For a memo or email, use To:. Omit Mr., Ms., Mrs., Miss, Dr., or Ir. Delete parentheses and the text inside them.',
    'Subject:\n' + incomingSubjectPrompt(),
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
