import { IOcrPageResult } from './IPdfOcr';
import { dearSirBandRegion, ISignatureAnalysis } from './signatureSender';

export interface IAiExtractionConfig {
  endpoint: string;
  apiKey: string;
  deployment: string;
  apiVersion: string;
}

export function isAiExtractionConfigured(config: IAiExtractionConfig): boolean {
  return config.endpoint.trim().length > 0 && config.apiKey.trim().length > 0;
}

const SUBJECT_PROMPT: string = [
  'Strictly extract content from the letter body ONLY, which is the text that appears AFTER the salutation (Dear Sir / Dear Madam / Dear Sir or Madam / etc.) and BEFORE the closing (Yours faithfully / Yours sincerely / Yours truly).',
  '',
  'Primary instruction (must follow first):',
  '- Locate and copy the entire underlined passage in that body section as one continuous block.',
  '- Do not stop at the first line after the salutation. Scan the whole body for any underlined text.',
  '',
  'Fallback (only if there is absolutely no underline anywhere in the body):',
  '- Then extract the Re: or Subject: line that appears in the same body range.',
  '',
  'Do not include the salutation or the closing in the result.'
].join('\n');

export async function extractFieldsWithAi(
  ocrText: string,
  fieldLabels: string[],
  config: IAiExtractionConfig,
  page?: IOcrPageResult,
  signature?: ISignatureAnalysis,
  receiverName?: string,
  subjectText?: string,
  refNo?: string,
  organization?: string
): Promise<{ [label: string]: string }> {
  const text = (ocrText || '').trim();
  if (fieldLabels.length === 0) {
    return {};
  }

  const clipped = text.length > 14000 ? text.substring(0, 14000) : text;
  const images = await buildPageImages(page, signature);
  if (!clipped && images.length === 0) {
    return {};
  }

  try {
    return await requestExtraction(clipped, fieldLabels, config, images, signature, receiverName, subjectText, refNo, organization);
  } catch {
    if (images.length === 0) {
      return {};
    }
    try {
      return await requestExtraction(clipped, fieldLabels, config, [], signature, receiverName, subjectText, refNo, organization);
    } catch {
      return {};
    }
  }
}

async function requestExtraction(
  ocrText: string,
  fieldLabels: string[],
  config: IAiExtractionConfig,
  images: IChatImage[],
  signature?: ISignatureAnalysis,
  receiverName?: string,
  subjectText?: string,
  refNo?: string,
  organization?: string
): Promise<{ [label: string]: string }> {
  const payload = buildRequest(ocrText, fieldLabels, config, images, signature, receiverName, subjectText, refNo, organization);
  let response: Response;
  try {
    response = await fetch(payload.url, {
      method: 'POST',
      headers: payload.headers,
      body: payload.body
    });
  } catch {
    throw new Error('CORS');
  }

  if (!response.ok) {
    throw new Error(`AI extraction failed with status ${response.status}.`);
  }

  const json = await response.json() as IChatCompletionResponse;
  const content = json.choices && json.choices[0] && json.choices[0].message
    ? (json.choices[0].message.content || '')
    : '';
  return parseFieldJson(content, fieldLabels);
}

function buildRequest(
  ocrText: string,
  fieldLabels: string[],
  config: IAiExtractionConfig,
  images: IChatImage[],
  signature?: ISignatureAnalysis,
  receiverName?: string,
  subjectText?: string,
  refNo?: string,
  organization?: string
): { url: string; headers: { [key: string]: string }; body: string } {
  const endpoint = config.endpoint.replace(/\/+$/, '');
  const isOpenAi = endpoint.indexOf('api.openai.com') >= 0;
  const apiVersion = config.apiVersion.trim() || '2024-08-01-preview';
  const url = isOpenAi
    ? `${endpoint}/v1/chat/completions`
    : `${endpoint}/openai/deployments/${encodeURIComponent(config.deployment.trim())}/chat/completions?api-version=${encodeURIComponent(apiVersion)}`;

  const headers: { [key: string]: string } = {
    'Content-Type': 'application/json'
  };
  if (isOpenAi) {
    headers.Authorization = `Bearer ${config.apiKey.trim()}`;
  } else {
    headers['api-key'] = config.apiKey.trim();
  }

  const prompt = buildUserPrompt(ocrText, fieldLabels, images.length > 0, signature, receiverName, subjectText, refNo, organization);
  const userContent: string | IChatPart[] = images.length > 0
    ? buildImagePromptParts(prompt, images)
    : prompt;

  const requestBody: {
    messages: { role: string; content: string | IChatPart[] }[];
    temperature: number;
    max_tokens: number;
    model?: string;
  } = {
    messages: [
      {
        role: 'system',
        content: 'You extract metadata from AECOM project correspondence. Use OCR text and the first-page image. Reply with JSON only. Use empty strings when a value is not clearly present. Copy original wording from text when it is visible. Organization is the line ending with Department that appears below By Post. Sender is the printed person name immediately below the handwritten signature, not the job title. Receiver is the value after Attn: if present; if there is no Attn, it is the person name below By Hand. Omit Mr., Ms., Mrs., Miss, and any parenthetical text. Subject:\n' + SUBJECT_PROMPT + '\nRef No is the value to the right of Our Ref:, or if that is missing, the value to the right of a standalone Ref:. Project Number is the 8 digits immediately before the slash in Our Ref. Do not invent values.'
      },
      {
        role: 'user',
        content: userContent
      }
    ],
    temperature: 0,
    max_tokens: 1200
  };
  if (isOpenAi) {
    requestBody.model = config.deployment.trim() || 'gpt-4o-mini';
  }
  const body = JSON.stringify(requestBody);

  return { url, headers, body };
}

function buildUserPrompt(
  ocrText: string,
  fieldLabels: string[],
  hasImage: boolean,
  signature?: ISignatureAnalysis,
  receiverName?: string,
  subjectText?: string,
  refNo?: string,
  organization?: string
): string {
  const fieldHelp = [
    'Name: document name or identifier. Registration Number must use this same value.',
    'Registration Number: always copy Name exactly. Do not invent a different value.',
    'Leading BL: leading business line. Must be one of: Architecture, Building Engineering, Environment, Geotechnical, Digital, Land Supply and Municipal, MEP, Project and Construction Management, Program, Cost and Consultancy, Transportation, Unclassified, Urbanism and Planning, Water. Abbreviations such as ARC, BEG, ENV, GEO, ISD, LSM, MEP, PCM, PCC, TRA, UNC, UAP, WAT are also accepted. If an AECOM business-line logo or name is visible, select that listed value.',
    'Project Number: the 8 digits immediately before the slash in Our Ref: / Our Ref :. For example Our Ref: 12345678/ABC -> 12345678. Digits only. Do not use Your Ref.',
    'Sub-Project Number: dropdown value None, or an integer from 1 to 99. Use None when it is not shown.',
    'Organization: copy the XXX Department line below By Post / By Post:. For example Highways Department. Do not copy Director of ... unless that line itself is the Department name. Do not use the letterhead.',
    'Sender: copy the printed person name immediately below the handwritten signature. If two names appear under Yours faithfully, use the lower name that sits just above the job title. Do not copy Chief Engineer, Director, Manager, or similar titles. Skip (signed).',
    'Receiver: if Attn: / Attn : / Attention: is present, copy only the value after that label. If there is no Attn, copy the person name below By Hand. Do not include Mr., Ms., Mrs., Miss, Dr., or Ir. Delete any parentheses and the text inside them. Do not copy Department lines.',
    'Subject:\n' + SUBJECT_PROMPT,
    'File No: file number',
    'Ref No: copy the value to the right of Our Ref: / Our Ref :. If there is no Our Ref, copy the value to the right of a standalone Ref: / Ref :. OCR may read Ref as Rref or Reef. Do not use Your Ref.',
    'Issue Date: the document date in dd/MM/yyyy, for example 03/09/2026.',
    'Attachment: Yes if the letter mentions attachments or enclosures, otherwise No. Reply only Yes or No.',
    'Scan: Yes if the document is a scan or has a scan mark, otherwise No. Reply only Yes or No.',
    'Remark: remarks or notes',
    'Location: office, site, or city',
    'cc to AECOM: Yes if AECOM is copied, otherwise No. Reply only Yes or No.'
  ].join('\n');

  const sourceLines = hasImage
    ? [
      'Extract these fields from the first page of a scanned document.',
      'Images: full first page, the addressee area for By Post / By Hand / Attn, the heading after Dear Sir, then the signature block if detected.',
      'Organization is the Department line below By Post. Receiver is Attn if present, otherwise the person name below By Hand without Mr./Ms. or parenthetical text. Subject: scan the whole letter body after the salutation for the entire underlined block; do not stop at the first line. Sender is the printed name immediately below the signature.'
    ]
    : [
      'Extract these fields from the OCR text of a document.',
      'Organization is the Department line below By Post. Sender is the person name below the signature. Receiver is Attn if present, otherwise the person name below By Hand. Subject: scan the whole letter body after the salutation for the entire underlined block; do not stop at the first line.'
    ];

  const parts = [
    sourceLines.join(' '),
    'Return a JSON object whose keys are exactly:',
    fieldLabels.join(', '),
    '',
    'Field meanings:',
    fieldHelp,
    '',
    'OCR text:',
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
    parts.push('', 'If Attn: is present, Receiver is the value after Attn:. If not, Receiver is the person name below By Hand, without Mr./Ms. or parenthetical text.');
  }
  if (organization && organization.trim()) {
    parts.push('', 'Detected Organization:', organization.trim());
  } else {
    parts.push('', 'Organization is the XXX Department line below By Post.');
  }
  parts.push('', SUBJECT_PROMPT);
  if (subjectText && subjectText.trim()) {
    parts.push('', 'OCR underlined candidate only. Follow the Subject rules above; scan the whole body and do not stop at the first line after the salutation:', subjectText.trim());
  }
  if (refNo && refNo.trim()) {
    parts.push('', 'Detected Ref No from Our Ref: or standalone Ref:', refNo.trim());
  } else {
    parts.push('', 'Find Our Ref: first. If it is missing, use a standalone Ref: / Ref :. Ref No is the value to the right of that label, not Your Ref.');
  }
  return parts.join('\n');
}

function buildImagePromptParts(prompt: string, images: IChatImage[]): IChatPart[] {
  const parts: IChatPart[] = [{ type: 'text', text: prompt }];
  images.forEach((image) => {
    parts.push({ type: 'text', text: image.label });
    parts.push({
      type: 'image_url',
      image_url: {
        url: image.url,
        detail: image.detail
      }
    });
  });
  return parts;
}

async function buildPageImages(
  page?: IOcrPageResult,
  signature?: ISignatureAnalysis
): Promise<IChatImage[]> {
  const source = page && page.imageUrl ? page.imageUrl.trim() : '';
  if (!source) {
    return [];
  }

  try {
    const image = await loadImage(source);
    const images: IChatImage[] = [];
    const fullPage = encodeImageRegion(image, 0, 0, image.width, image.height, 1024, 0.62);
    if (fullPage) {
      images.push({ url: fullPage, detail: 'low', label: 'Full first page:' });
    }
    const band = dearSirBandRegion(page);
    if (band) {
      const crop = encodeImageRegion(
        image,
        band.x0,
        band.y0,
        Math.max(8, band.x1 - band.x0),
        Math.max(8, band.y1 - band.y0),
        1280,
        0.78
      );
      if (crop) {
        images.push({
          url: crop,
          detail: 'high',
          label: 'Letter body AFTER the salutation and BEFORE the closing. Subject: scan this whole body for the entire underlined block; do not stop at the first line. Fallback only if there is no underline: Re: or Subject: line:'
        });
      }
    }
    const region = signature && signature.region;
    if (region) {
      const below = Math.max(110, Math.round(image.height * 0.12));
      const sx = Math.max(0, region.x0 - 12);
      const sy = Math.max(0, region.y0 - 8);
      const crop = encodeImageRegion(
        image,
        sx,
        sy,
        image.width - sx,
        Math.min(image.height - sy, region.y1 - region.y0 + below + 8),
        1280,
        0.75
      );
      if (crop) {
        images.push({
          url: crop,
          detail: 'high',
          label: 'Closing block. Sender is the printed person name immediately below the handwritten signature:'
        });
      }
    }
    return images;
  } catch {
    return [];
  }
}

function loadImage(url: string): Promise<HTMLImageElement> {
  return new Promise((resolve, reject) => {
    const image = new Image();
    image.onload = () => resolve(image);
    image.onerror = () => reject(new Error('Unable to load page image.'));
    image.src = url;
  });
}

function encodeImageRegion(
  image: HTMLImageElement,
  sx: number,
  sy: number,
  sw: number,
  sh: number,
  maxWidth: number,
  quality: number
): string | undefined {
  if (sw < 8 || sh < 8) {
    return undefined;
  }
  const scale = sw > maxWidth ? maxWidth / sw : 1;
  const canvas = document.createElement('canvas');
  canvas.width = Math.max(1, Math.round(sw * scale));
  canvas.height = Math.max(1, Math.round(sh * scale));
  const context = canvas.getContext('2d');
  if (!context) {
    return undefined;
  }
  context.drawImage(image, sx, sy, sw, sh, 0, 0, canvas.width, canvas.height);
  const dataUrl = canvas.toDataURL('image/jpeg', quality);
  canvas.width = 0;
  canvas.height = 0;
  return dataUrl.indexOf('data:image/jpeg') === 0 ? dataUrl : undefined;
}

function parseFieldJson(content: string, fieldLabels: string[]): { [label: string]: string } {
  const values: { [label: string]: string } = {};
  fieldLabels.forEach((label) => {
    values[label] = '';
  });

  const raw = (content || '').trim();
  if (!raw) {
    return values;
  }

  let parsed: { [key: string]: unknown };
  try {
    parsed = JSON.parse(stripFence(raw)) as { [key: string]: unknown };
  } catch {
    return values;
  }

  fieldLabels.forEach((label) => {
    const direct = parsed[label];
    if (typeof direct === 'string' && direct.trim()) {
      values[label] = direct.trim();
      return;
    }
    const matchedKey = Object.keys(parsed).filter((key) => key.toLowerCase() === label.toLowerCase())[0];
    const matched = matchedKey ? parsed[matchedKey] : undefined;
    if (typeof matched === 'string' && matched.trim()) {
      values[label] = matched.trim();
    }
  });

  return values;
}

function stripFence(content: string): string {
  const trimmed = content.trim();
  if (trimmed.indexOf('```') !== 0) {
    return trimmed;
  }
  const withoutOpen = trimmed.replace(/^```json\s*/i, '').replace(/^```\s*/, '');
  const close = withoutOpen.lastIndexOf('```');
  return close >= 0 ? withoutOpen.substring(0, close).trim() : withoutOpen.trim();
}

interface IChatImage {
  url: string;
  detail: 'low' | 'high';
  label: string;
}

interface IChatPart {
  type: 'text' | 'image_url';
  text?: string;
  image_url?: {
    url: string;
    detail: 'low' | 'high';
  };
}

interface IChatCompletionResponse {
  choices?: {
    message?: {
      content?: string;
    };
  }[];
}
