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
  config: IAiExtractionConfig
): Promise<{ [label: string]: string }> {
  const text = (ocrText || '').trim();
  if (!text || fieldLabels.length === 0) {
    return {};
  }

  const clipped = text.length > 14000 ? text.substring(0, 14000) : text;
  const payload = buildRequest(clipped, fieldLabels, config);
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
  config: IAiExtractionConfig
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

  const requestBody: {
    messages: { role: string; content: string }[];
    temperature: number;
    max_tokens: number;
    model?: string;
  } = {
    messages: [
      {
        role: 'system',
        content: 'You extract metadata from OCR text of AECOM project correspondence. Reply with JSON only. Use empty strings when a value is not clearly present. Copy original wording. Do not invent values.'
      },
      {
        role: 'user',
        content: buildUserPrompt(ocrText, fieldLabels)
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

function buildUserPrompt(ocrText: string, fieldLabels: string[]): string {
  const fieldHelp = [
    'Name: document name or identifier. Registration Number must use this same value.',
    'Registration Number: always copy Name exactly. Do not invent a different value.',
    'Leading BL: leading business line. Must be one of: Architecture, Building Engineering, Environment, Geotechnical, Digital, Land Supply and Municipal, MEP, Project and Construction Management, Program, Cost and Consultancy, Transportation, Unclassified, Urbanism and Planning, Water. Abbreviations such as ARC, BEG, ENV, GEO, ISD, LSM, MEP, PCM, PCC, TRA, UNC, UAP, WAT are also accepted.',
    'Project Number: up to 8 digits only. Copy digits only and omit letters, spaces, and extra characters.',
    'Sub-Project Number: sub-project or task number',
    'Organization: company or organization',
    'Sender: who sent the document',
    'Receiver: who the document is addressed to',
    'Subject: subject or title of the document',
    'File No: file number',
    'Ref No: reference number',
    'Issue Date: issue or document date',
    'Attachment: attachments mentioned',
    'Scan: scan number or scan mark',
    'Remark: remarks or notes',
    'Location: office, site, or city',
    'cc to AECOM: CC line if AECOM is copied'
  ].join('\n');

  return [
    'Extract these fields from the OCR text of a document. The layout changes between files and labels may be missing or different.',
    'Return a JSON object whose keys are exactly:',
    fieldLabels.join(', '),
    '',
    'Field meanings:',
    fieldHelp,
    '',
    'OCR text:',
    ocrText
  ].join('\n');
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

interface IChatCompletionResponse {
  choices?: {
    message?: {
      content?: string;
    };
  }[];
}
