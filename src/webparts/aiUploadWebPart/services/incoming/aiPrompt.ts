// Incoming AI prompt only. Outgoing lives in services/outgoing/aiPrompt.ts.
import { IAiExtractionHints } from '../correspondenceTypes';
import { OUTGOING_SUBJECT_PROMPT, outgoingBodyImageLabel } from '../outgoing/aiPrompt';

export function incomingSystemPrompt(letterType?: string): string {
  if (letterType === 'email') {
    return incomingEmailSystemPrompt();
  }
  return incomingLetterSystemPrompt();
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
  if (hints && hints.letterType === 'email') {
    return buildIncomingEmailUserPrompt(ocrText, fieldLabels, hasImage, hints);
  }
  return buildIncomingLetterUserPrompt(ocrText, fieldLabels, hasImage, hints);
}

function incomingLetterSystemPrompt(): string {
  return [
    'You extract metadata from incoming correspondence sent TO AECOM, written by another organisation.',
    'Use OCR text and the first-page image. Reply with JSON only. Use empty strings when a value is not clearly present.',
    'Copy original wording from text when it is visible. OCR text may include <u>underlined</u> and <b>bold</b> tags; never copy those tags into values.',
    'Organization is the left-hand addressee line immediately above a floor line such as 12/F, or if there is no xx/F, immediately above a line containing xxx Road. The address is only on the left of the page; ignore Our Ref, Date, and other text on the right of the same row. Do not use the letterhead.',
    'Sender is exactly the text inside the parentheses immediately below Yours faithfully / Yours sincerely / Yours truly / 署名, for example (Ben xXx. LXX) -> Ben xXx. LXX. That closing is often on the last page, not page 1. Skip (signed). Do not copy CC / c.c. / 副本 names below the signature. Do not copy parentheses from Attn or the address on page 1. Do not copy the job title under the parentheses.',
    'Receiver is the Attn value when Attn / Attention is present. If there is no Attn, Receiver is the first left-hand address line above Dear / 敬啟者.',
    'Ref No is Our Ref / 本處檔號 / 檔號 of the originating party, not Your Ref.',
    'Leave Project Number empty. Do not take it from Your Ref or Our Ref.',
    'Do not invent values.'
  ].join(' ') + '\nSubject:\n' + OUTGOING_SUBJECT_PROMPT + '\nIncoming extra: the Subject heading may start on the latter part of a line after Dear (right of Dear Sir, or after unstyled words on the left) and continue on the next line. Copy from that latter part through the following line. Do not copy the unstyled left-hand prefix.';
}

function incomingEmailSystemPrompt(): string {
  return [
    'You extract metadata from an incoming email printout sent TO AECOM.',
    'Use OCR text and the first-page image. Reply with JSON only. Use empty strings when a value is not clearly present.',
    'Copy original wording from text when it is visible. OCR text may include <u>underlined</u> and <b>bold</b> tags; never copy those tags into values.',
    'Use ONLY the last email in the printout (the last From / Sent / To / Subject block, often after Original Message or forwarded content). Ignore earlier emails above it.',
    'Sender is the FIRST name in CC: / Cc: / 抄送 on the LAST email-format page only, rewritten as Surname, Given (AAA BBB -> BBB, AAA). Do not use From:. Do not use c.c. / 副本 in a following letter.',
    'Receiver: first take the name immediately below Regards / Best regards / Kind regards in the last email. If that name also appears in CC: of the last email, use that CC name. If it does not match CC, use the To: / 收件人 value of that last email. Do not use Attn or Dear from a following letter.',
    'Subject is ONLY the value after Subject: / 主旨: of that last email. STOP before the next header such as Sent: / From: / To: / Attachments:. Do not copy that next header or anything after it. Do not use letter titles after Dear.',
    'Issue Date is that last Sent: or Date:.',
    'Ref No is the value immediately after Our Ref: / 本處檔號. Look in the last email first; if it is not there, copy Our Ref: from the following letter. Do not use Your Ref.',
    'Organization is the left-hand addressee line immediately above a floor line such as 12/F, or if there is no xx/F, immediately above a line containing xxx Road. Do not use the letterhead.',
    'Leave Project Number empty in the JSON. It is filled later as {Code}-EOI from Root URL Mapping List after Leading BL is selected.',
    'Do not invent values.'
  ].join(' ');
}

function buildIncomingLetterUserPrompt(
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
  const agreementNo = hints && hints.agreementNo;
  const organization = hints && hints.organization;
  const letterType = hints && hints.letterType;

  const sourceLines = hasImage
    ? [
      'Extract these fields from an incoming letter sent TO AECOM.',
      'Images: full first page, letterhead at the top, the body after the salutation, then the signature block if detected.',
      'Organization is the LEFT-HAND line above xx/F, or if there is no floor line, the line above xxx Road. Not the whole row and not the letterhead. Receiver is the Attn value if present; if there is no Attn, it is the first left-hand address line above Dear. Ignore Our Ref / Date on the right. Subject: after Dear, copy bold+underline text even if it starts on the latter part of a line and continues on the next line. Do not copy the unstyled left prefix or a line that only sits above a table. Sender is the text inside the signature parentheses.'
    ]
    : [
      'Extract these fields from the OCR text of an incoming letter sent TO AECOM.',
      'Organization is the LEFT-HAND line above xx/F, or if there is no floor line, the line above xxx Road, above Dear / 敬啟者. Not the whole row and not the letterhead. Ignore right-column Our Ref / Date. Sender is the text inside the signature parentheses. Receiver is the Attn value if present; if there is no Attn, the first left-hand address line above Dear. Subject: after Dear, copy bold+underline text even if it starts on the latter part of a line and continues on the next line. Do not copy a line that only sits above a table or divider.'
    ];

  const parts = [
    sourceLines.join(' '),
    'Return a JSON object whose keys are exactly:',
    fieldLabels.join(', '),
    '',
    'Field meanings:',
    incomingLetterFieldHelp(),
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
    parts.push('', 'Detected Subject from bold+underline text after Dear. It may start mid-line and continue on the next line. Use this text:', subjectText.trim());
  }
  if (refNo && refNo.trim()) {
    parts.push('', 'Detected Ref No from Our Ref: or standalone Ref:', refNo.trim());
  } else {
    parts.push('', 'Find Our Ref: / 本處檔號 / 檔號 first. Ref No is the value to the right of that label, not Your Ref.');
  }
  if (yourRef && yourRef.trim()) {
    parts.push('', 'Detected Your Ref. Do not use this for Project Number:', yourRef.trim());
  }
  if (agreementNo && agreementNo.trim()) {
    parts.push('', 'Detected Agreement No. / Contract No. below Dear and above Subject. Leave Project Number empty; it is resolved from Notification Set-up:', agreementNo.trim());
  } else {
    parts.push('', 'Look below Dear and above Subject for Agreement No. or Contract No:. Copy the text after that label only as context. Leave Project Number empty.');
  }
  return parts.join('\n');
}

function buildIncomingEmailUserPrompt(
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

  const sourceLines = hasImage
    ? [
      'Extract these fields from an incoming email printout sent TO AECOM.',
      'Images: the last email-format page. Ignore any letter that follows the emails for Sender / Receiver / Subject.',
      'Use ONLY the last email (the last From / Sent / To / Subject block). Sender is the FIRST name in CC: on that last email-format page, rewritten as Surname, Given. Receiver: if the name below Regards also appears in CC:, use that CC name; otherwise use that last To:. Subject is ONLY the value after Subject: / 主旨:, and STOP before the next header. Issue Date is that last Sent: or Date:. Ref No is Our Ref: in the last email, or if missing, Our Ref: in the following letter.'
    ]
    : [
      'Extract these fields from the OCR text of an incoming email printout sent TO AECOM.',
      'Use ONLY the last email. Sender is the FIRST name in CC: on the last email-format page, rewritten as Surname, Given. Receiver: if the name below Regards also appears in CC:, use that CC name; otherwise use that last To:. Subject is ONLY the value after Subject: / 主旨:, and STOP before the next header. Issue Date is that last Sent: or Date:. Ref No is Our Ref: in the last email, or if missing, Our Ref: in the following letter.'
    ];

  const parts = [
    sourceLines.join(' '),
    'Return a JSON object whose keys are exactly:',
    fieldLabels.join(', '),
    '',
    'Field meanings:',
    incomingEmailFieldHelp(),
    '',
    'OCR text (plain words plus <u>underline</u> and <b>bold</b> tags when detected):',
    ocrText || '(none)',
    '',
    'Detected incoming format:',
    'email'
  ];
  if (signature && signature.senderName) {
    parts.push('', 'Detected Sender from CC: on the last email-format page only:', signature.senderName);
  } else {
    parts.push('', 'Sender is the FIRST name in CC: / Cc: / 抄送 on the LAST email-format page only, rewritten as Surname, Given (AAA BBB -> BBB, AAA). Do not use From:, earlier email pages, or c.c. / 副本 in a following letter.');
  }
  if (receiverName && receiverName.trim()) {
    parts.push('', 'Detected Receiver. Prefer the name below Regards when it also appears in CC:; otherwise To: of the last email:', receiverName.trim());
  } else {
    parts.push('', 'Receiver: if the name below Regards / Best regards in the LAST email also appears in CC:, use that CC name. Otherwise use To: / 收件人 of the LAST email only. Do not use earlier emails or Attn in a following letter.');
  }
  if (organization && organization.trim()) {
    parts.push('', 'Detected Organization from the LEFT-HAND line above xx/F or xxx Road in the addressee address (not the whole row, not the letterhead):', organization.trim());
  } else {
    parts.push('', 'Organization is the left-hand line immediately above xx/F, or if there is no xx/F, the line above xxx Road in the addressee address. Do not copy right-column text from that row, and do not use the letterhead.');
  }
  if (subjectText && subjectText.trim()) {
    parts.push('', 'Detected Subject from the last email Subject: only. Stop before the next title. Use this text:', subjectText.trim());
  } else {
    parts.push('', 'Subject is ONLY the value after Subject: / 主旨: of the last email. STOP before Sent: / From: / To: / Attachments:. Do not copy a letter title after Dear.');
  }
  if (refNo && refNo.trim()) {
    parts.push('', 'Detected Ref No from Our Ref: in the last email or the following letter:', refNo.trim());
  } else {
    parts.push('', 'Ref No is the value after Our Ref: / 本處檔號. Search the last email first; if it is missing, use Our Ref: in the following letter. Do not use Your Ref.');
  }
  if (yourRef && yourRef.trim()) {
    parts.push('', 'Detected Your Ref. Do not use this for Project Number:', yourRef.trim());
  }
  parts.push('', 'Leave Project Number empty. The form fills {Code}-EOI from Root URL Mapping List after Leading BL is selected.');
  return parts.join('\n');
}

function incomingLetterFieldHelp(): string {
  return [
    'Name: document name or identifier. Registration Number must use this same value.',
    'Registration Number: always copy Name exactly. Do not invent a different value.',
    'Leading BL: leave empty unless an AECOM business-line name is clearly shown. Must be one of: Architecture, Building Engineering, Environment, Geotechnical, Digital, Land Supply and Municipal, MEP, Project and Construction Management, Program, Cost and Consultancy, Transportation, Unclassified, Urbanism and Planning, Water.',
    'Project Number: leave empty. Do not copy Your Ref, Our Ref, Agreement No., or Contract No. into this field.',
    'Sub-Project Number: dropdown value None, or an integer from 1 to 99. Use None when it is not shown.',
    'Organization: in the LEFT addressee address above Dear / 敬啟者, find a floor line such as 12/F or G/F and copy only the left-hand line immediately above it. If there is no xx/F, find a line containing xxx Road and copy the left-hand line immediately above that. Do not include Our Ref, Your Ref, Date, or other text on the right of that row. Do not use letterhead.',
    'Sender: copy ONLY the text inside the parentheses immediately below Yours faithfully / Yours sincerely / Yours truly / 署名, e.g. (Ben xXx. LXX) -> Ben xXx. LXX. Look on the page that contains that closing, which is often the last page, not the first page. Skip (signed). Do not copy CC / c.c. / 副本 / copy to names, whether they sit below the signature. Do not copy parentheses from Attn: near the top of page 1. Do not copy the job title under the parentheses. For a memo, use From:.',
    'Receiver: if Attn: / Attn : / Attention: is present, copy only the value after that label. If there is no Attn, copy the first left-hand address line above Dear / 敬啟者. Do not use the line below Yours faithfully. Ignore Our Ref / Date on the right of the same row.',
    'Subject:\n' + OUTGOING_SUBJECT_PROMPT + '\nIncoming extra: the heading may start on the latter part of a line after Dear and continue on the next line. Copy from that latter part through the next line. Do not copy the unstyled left-hand prefix.',
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

function incomingEmailFieldHelp(): string {
  return [
    'Name: document name or identifier. Registration Number must use this same value.',
    'Registration Number: always copy Name exactly. Do not invent a different value.',
    'Leading BL: leave empty unless an AECOM business-line name is clearly shown. Must be one of: Architecture, Building Engineering, Environment, Geotechnical, Digital, Land Supply and Municipal, MEP, Project and Construction Management, Program, Cost and Consultancy, Transportation, Unclassified, Urbanism and Planning, Water.',
    'Project Number: leave empty here. The form fills {Code}-EOI from Root URL Mapping List after Leading BL is selected. Do not copy Your Ref, Our Ref, Agreement No., or Contract No.',
    'Sub-Project Number: dropdown value None, or an integer from 1 to 99. Use None when it is not shown.',
    'Organization: in the LEFT addressee address, find a floor line such as 12/F or G/F and copy only the left-hand line immediately above it. If there is no xx/F, find a line containing xxx Road and copy the left-hand line immediately above that. Do not use letterhead.',
    'Sender: the FIRST name in CC: / Cc: / 抄送 on the LAST email-format page only, rewritten as Surname, Given (AAA BBB -> BBB, AAA). Do not use From:. Do not use earlier email pages. Do not use c.c. / 副本 in a following letter.',
    'Receiver: first copy the name immediately below Regards / Best regards / Kind regards in the last email if that name also appears in CC:. If it does not match CC, copy the To: / 收件人 value of the LAST email only. Do not use Attn, Dear, or a following letter.',
    'Subject: copy ONLY the value after Subject: / 主旨: of the last email, and STOP before the next title such as Sent: / From: / To: / Attachments:. Do not copy that next title. Do not use a letter title after Dear.',
    'File No: file number if shown separately from Ref No.',
    'Ref No: copy the value after Our Ref: / 本處檔號 in the last email, or if it is not in the email, from the following letter. Do not use Your Ref.',
    'Issue Date: the last email Sent: or Date: in dd/MM/yyyy, for example 03/09/2026. Accept 8 September 2026 and 2026年9月8日.',
    'Attachment: Yes if the email mentions attachments, enclosures, 附件, or 隨函, otherwise No. Reply only Yes or No.',
    'Scan: Yes if the document is a scan or has a scan mark, otherwise No. Reply only Yes or No.',
    'Remark: remarks or notes',
    'Location: office, site, or city',
    'cc to AECOM: Yes if the CC list includes AECOM people beyond the addressee, otherwise No. Reply only Yes or No.'
  ].join('\n');
}
