import { CorrespondenceKind } from '../constants/incomingName';
import { IAiExtractionHints } from './correspondenceTypes';
import { IOcrPageResult } from './IPdfOcr';
import { buildIncomingUserPrompt, incomingBodyImageLabel, incomingClosingImageLabel, incomingLetterheadImageLabel, incomingSystemPrompt } from './incoming/aiPrompt';
import { buildOutgoingUserPrompt, outgoingBodyImageLabel, outgoingClosingImageLabel, outgoingSystemPrompt } from './outgoing/aiPrompt';
import { dearSirBandRegion, ISignatureAnalysis } from './signatureSender';

export type { IAiExtractionHints } from './correspondenceTypes';

export interface IAiExtractionConfig {
  endpoint: string;
  apiKey: string;
  deployment: string;
  apiVersion: string;
}

export function isAiExtractionConfigured(config: IAiExtractionConfig): boolean {
  return config.endpoint.trim().length > 0 && config.apiKey.trim().length > 0;
}

export async function extractFieldsWithAi(
  ocrText: string,
  fieldLabels: string[],
  config: IAiExtractionConfig,
  hints?: IAiExtractionHints
): Promise<{ [label: string]: string }> {
  const text = (ocrText || '').trim();
  if (fieldLabels.length === 0) {
    return {};
  }

  const clipped = text.length > 14000 ? text.substring(0, 14000) : text;
  const images = await buildPageImages(hints && hints.page, hints && hints.signature, hints && hints.kind);
  if (!clipped && images.length === 0) {
    return {};
  }

  try {
    return await requestExtraction(clipped, fieldLabels, config, images, hints);
  } catch {
    if (images.length === 0) {
      return {};
    }
    try {
      return await requestExtraction(clipped, fieldLabels, config, [], hints);
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
  hints?: IAiExtractionHints
): Promise<{ [label: string]: string }> {
  const payload = buildRequest(ocrText, fieldLabels, config, images, hints);
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
  hints?: IAiExtractionHints
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

  const incoming = hints && hints.kind === 'incoming';
  const prompt = incoming
    ? buildIncomingUserPrompt(ocrText, fieldLabels, images.length > 0, hints)
    : buildOutgoingUserPrompt(ocrText, fieldLabels, images.length > 0, hints);
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
        content: incoming ? incomingSystemPrompt() : outgoingSystemPrompt()
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
  signature?: ISignatureAnalysis,
  kind?: CorrespondenceKind
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
    if (kind === 'incoming') {
      const headHeight = Math.max(80, Math.round(image.height * 0.22));
      const letterhead = encodeImageRegion(image, 0, 0, image.width, headHeight, 1280, 0.78);
      if (letterhead) {
        images.push({
          url: letterhead,
          detail: 'high',
          label: incomingLetterheadImageLabel()
        });
      }
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
          label: kind === 'incoming' ? incomingBodyImageLabel() : outgoingBodyImageLabel()
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
          label: kind === 'incoming' ? incomingClosingImageLabel() : outgoingClosingImageLabel()
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
      values[label] = stripStyleTags(direct.trim());
      return;
    }
    const matchedKey = Object.keys(parsed).filter((key) => key.toLowerCase() === label.toLowerCase())[0];
    const matched = matchedKey ? parsed[matchedKey] : undefined;
    if (typeof matched === 'string' && matched.trim()) {
      values[label] = stripStyleTags(matched.trim());
    }
  });

  return values;
}

function stripStyleTags(value: string): string {
  return (value || '').replace(/<\/?(?:u|b)>/gi, '').replace(/\s+/g, ' ').trim();
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
