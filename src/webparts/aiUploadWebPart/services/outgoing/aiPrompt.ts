// Outgoing AI prompt only. Incoming lives in services/incoming/aiPrompt.ts.
import { IAiExtractionHints } from '../correspondenceTypes';

export const OUTGOING_SUBJECT_PROMPT: string = [
  'Strictly extract content from the letter body ONLY, which is the text that appears AFTER the salutation (Dear Sir / Dear Madam / Dear Sir or Madam / etc.) and BEFORE the closing (Yours faithfully / Yours sincerely / Yours truly).',
  '',
  'Primary instruction (must follow first):',
  '- After the salutation, find lines that contain both <b> bold and <u> underlined text.',
  '- Copy each matching line in full, including words on the same line that are not inside the tags.',
  '- If the next matching line is a wrap of the same heading, join it with a single space. Do not join lines that are far apart.',
  '',
  'If no line has both bold and underline:',
  '- Then copy the first underlined <u> line only.',
  '- Then copy the bold heading in the same body range. Prefer text wrapped in <b>...</b> tags.',
  '- Do not copy a random bold word inside a normal sentence.',
  '',
  'Fallback (only if there is absolutely no underline or bold heading anywhere in the body):',
  '- Then extract the Re: or Subject: line that appears in the same body range.',
  '',
  'Do not include the salutation or the closing in the result.',
  'Do not copy <u>, </u>, <b>, or </b> tags into the field value.'
].join('\n');

export function outgoingSystemPrompt(): string {
  return 'You extract metadata from AECOM project correspondence. Use OCR text and the first-page image. Reply with JSON only. Use empty strings when a value is not clearly present. Copy original wording from text when it is visible. OCR text may include <u>underlined</u> and <b>bold</b> tags; never copy those tags into values. Organization is every full line containing Department that appears below Our Ref and above Dear. Sender is the printed person name immediately below the handwritten signature, not the job title. Receiver is the value after Attn: if present; if there is no Attn, it is the person name above xx/F, skipping Department or Director lines. Omit Mr., Ms., Mrs., Miss, and any parenthetical text. Subject:\n' + OUTGOING_SUBJECT_PROMPT + '\nRef No is the value to the right of Our Ref:, or if that is missing, the value to the right of a standalone Ref:. Project Number is the 8 digits immediately before the slash in Our Ref; if there is no slash, the 8 digits immediately before the hyphen. Do not invent values.';
}

export function outgoingBodyImageLabel(): string {
  return 'Letter body AFTER the salutation and BEFORE the closing. Subject: copy the entire first line that has both <b> and <u>. If none, the first underlined line, then bold heading. Fallback: Re: or Subject: line:';
}

export function outgoingClosingImageLabel(): string {
  return 'Closing block. Sender is the printed person name immediately below the handwritten signature:';
}

export function buildOutgoingUserPrompt(
  ocrText: string,
  fieldLabels: string[],
  hasImage: boolean,
  hints?: IAiExtractionHints
): string {
  const signature = hints && hints.signature;
  const receiverName = hints && hints.receiverName;
  const subjectText = hints && hints.subjectText;
  const refNo = hints && hints.refNo;
  const organization = hints && hints.organization;

  const sourceLines = hasImage
    ? [
      'Extract these fields from the first page of a scanned document.',
      'Images: full first page, the addressee area for By Post / By Hand / Attn, the heading after Dear Sir, then the signature block if detected.',
      'Organization is every full line containing Department that sits below Our Ref and above Dear. Receiver is Attn if present, otherwise the person name above xx/F (skip Department/Director lines and keep going up). Omit Mr./Ms. and parenthetical text. Subject: after Dear Sir/Madam, copy the entire first line that has both <b> and <u>. Sender is the printed name immediately below the signature.'
    ]
    : [
      'Extract these fields from the OCR text of a document.',
      'Organization is every full line containing Department that sits below Our Ref and above Dear. Sender is the person name below the signature. Receiver is Attn if present, otherwise the person name above xx/F (skip Department/Director and keep going up). Subject: after Dear Sir/Madam, copy the entire first line that has both <b> and <u>.'
    ];

  const parts = [
    sourceLines.join(' '),
    'Return a JSON object whose keys are exactly:',
    fieldLabels.join(', '),
    '',
    'Field meanings:',
    outgoingFieldHelp(),
    '',
    'OCR text (plain words plus <u>underline</u> and <b>bold</b> tags when detected):',
    ocrText || '(none)'
  ];
  if (signature && signature.senderName) {
    parts.push('', 'Detected Sender from the printed name below the signature:', signature.senderName);
  } else if (signature && signature.textBelow) {
    parts.push('', 'OCR immediately below the signature:', signature.textBelow);
  } else {
    parts.push('', 'Look below the handwritten signature for the printed person name. That is Sender.');
  }
  if (receiverName && receiverName.trim()) {
    parts.push('', 'Detected Receiver:', receiverName.trim());
  } else {
    parts.push('', 'If Attn: is present, Receiver is the value after Attn:. If not, Receiver is the person name above xx/F. If that line is a Department, use the name above it. Omit Mr./Ms. and parenthetical text.');
  }
  if (organization && organization.trim()) {
    parts.push('', 'Detected Organization:', organization.trim());
  } else {
    parts.push('', 'Organization is every full line containing Department below Our Ref and above Dear. If none, use the line above xx/F without symbols.');
  }
  parts.push('', OUTGOING_SUBJECT_PROMPT);
  if (subjectText && subjectText.trim()) {
    parts.push('', 'Detected Subject from nearby lines after Dear that have both bold and underline. Use this text:', subjectText.trim());
  }
  if (refNo && refNo.trim()) {
    parts.push('', 'Detected Ref No from Our Ref: or standalone Ref:', refNo.trim());
  } else {
    parts.push('', 'Find Our Ref: first. If it is missing, use a standalone Ref: / Ref :. Ref No is the value to the right of that label, not Your Ref.');
  }
  return parts.join('\n');
}

function outgoingFieldHelp(): string {
  return [
    'Name: document name or identifier. Registration Number must use this same value.',
    'Registration Number: always copy Name exactly. Do not invent a different value.',
    'Leading BL: leading business line. Must be one of: Architecture, Building Engineering, Environment, Geotechnical, Digital, Land Supply and Municipal, MEP, Project and Construction Management, Program, Cost and Consultancy, Transportation, Unclassified, Urbanism and Planning, Water. Abbreviations such as ARC, BEG, ENV, GEO, ISD, LSM, MEP, PCM, PCC, TRA, UNC, UAP, WAT are also accepted. If an AECOM business-line logo or name is visible, select that listed value.',
    'Project Number: first take the 8 digits immediately before the slash in Our Ref: / Our Ref :. For example Our Ref: 12345678/ABC -> 12345678. If there is no slash, take the 8 digits immediately before the hyphen. For example Our Ref: 12345678-ABC -> 12345678. Digits only. Do not use Your Ref.',
    'Sub-Project Number: dropdown value None, or an integer from 1 to 99. Use None when it is not shown.',
    'Organization: ABOVE the Dear line and BELOW Our Ref:, copy every full line that contains the word Department. If there are several such lines, join them with a space. If none, copy the entire line immediately above a floor line such as 12/F, and strip all punctuation and symbols. Do not use letterhead above Our Ref. Do not use lines after Dear.',
    'Sender: copy the printed person name immediately below the handwritten signature. If two names appear under Yours faithfully, use the lower name that sits just above the job title. Do not copy Chief Engineer, Director, Manager, or similar titles. Skip (signed).',
    'Receiver: if Attn: / Attn : / Attention: is present, copy only the value after that label. If there is no Attn, find a floor line such as 12/F and copy the person name above it. If the line above 12/F is a Department or Director line, keep going up until the person name, which usually starts with Mr / Ms. Do not include Mr., Ms., Mrs., Miss, Dr., or Ir. Delete any parentheses and the text inside them. Do not copy Department lines.',
    'Subject:\n' + OUTGOING_SUBJECT_PROMPT,
    'File No: file number',
    'Ref No: copy the value to the right of Our Ref: / Our Ref :. If there is no Our Ref, copy the value to the right of a standalone Ref: / Ref :. OCR may read Ref as Rref or Reef. Do not use Your Ref.',
    'Issue Date: the document date in dd/MM/yyyy, for example 03/09/2026.',
    'Attachment: Yes if the letter mentions attachments or enclosures, otherwise No. Reply only Yes or No.',
    'Scan: Yes if the document is a scan or has a scan mark, otherwise No. Reply only Yes or No.',
    'Remark: remarks or notes',
    'Location: office, site, or city',
    'cc to AECOM: Yes if AECOM is copied, otherwise No. Reply only Yes or No.'
  ].join('\n');
}
