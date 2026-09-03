import { IOcrPageResult, IOcrWord } from './IPdfOcr';
import { joinOcrWords } from './ocrSelection';

export interface ISignatureRegion {
  x0: number;
  y0: number;
  x1: number;
  y1: number;
}

export interface ISignatureAnalysis {
  region: ISignatureRegion | undefined;
  senderName: string;
  textBelow: string;
}

export async function analyzeSignature(page?: IOcrPageResult): Promise<ISignatureAnalysis> {
  const fromClosing = extractNameBetweenClosingAndTitle(page);
  const empty: ISignatureAnalysis = {
    region: undefined,
    senderName: fromClosing,
    textBelow: ''
  };
  if (!page || !page.imageUrl || page.width <= 0 || page.height <= 0) {
    return empty;
  }

  try {
    const image = await loadImage(page.imageUrl);
    const regions = findInkSignatures(image, page.words || []);
    let chosen: ISignatureAnalysis = empty;
    for (let index = 0; index < regions.length; index++) {
      const region = regions[index];
      const textBelow = joinOcrWords(wordsBelowSignature(page, region));
      const senderName = firstPersonName(textBelow);
      if (senderName) {
        chosen = { region, textBelow, senderName };
      } else if (!chosen.region) {
        chosen = { region, textBelow, senderName: '' };
      }
    }
    if (fromClosing) {
      return {
        ...chosen,
        senderName: fromClosing
      };
    }
    return chosen;
  } catch {
    return empty;
  }
}

export function asPersonName(value: string): string {
  return personNameFromLine(value);
}

export function extractReceiverAboveDearSir(page?: IOcrPageResult): string {
  if (!page) {
    return '';
  }
  try {
    const fromAttn = extractAttnValue(page);
    if (fromAttn) {
      return fromAttn;
    }
    const fromWords = addresseeBlockFromWords(page);
    if (fromWords) {
      return fromWords;
    }
    return addresseeBlockFromText(page.text || '');
  } catch {
    return '';
  }
}

export async function extractSubjectBelowDearSir(page?: IOcrPageResult): Promise<string> {
  if (!page) {
    return '';
  }
  return subjectFromUnderline(page);
}

export function dearSirBandRegion(page?: IOcrPageResult): ISignatureRegion | undefined {
  if (!page || page.width <= 0 || page.height <= 0) {
    return undefined;
  }
  const hit = findSalutationHit(page.words || []);
  if (!hit) {
    return {
      x0: 0,
      y0: Math.round(page.height * 0.12),
      x1: Math.round(page.width * 0.96),
      y1: Math.round(page.height * 0.68)
    };
  }
  return {
    x0: 0,
    y0: Math.max(0, hit.y0 - Math.max(140, page.height * 0.28)),
    x1: Math.round(page.width * 0.96),
    y1: Math.min(page.height, Math.max(hit.y1 + Math.max(160, page.height * 0.42), page.height * 0.68))
  };
}

function extractAttnValue(page: IOcrPageResult): string {
  const fromWords = attnFromWords(page);
  if (fromWords) {
    return fromWords;
  }
  return attnFromText(page.text || '');
}

function attnFromWords(page: IOcrPageResult): string {
  const words = (page.words || []).slice().sort((left, right) => {
    if (Math.abs(left.y0 - right.y0) > 8) {
      return left.y0 - right.y0;
    }
    return left.x0 - right.x0;
  });
  for (let index = 0; index < words.length; index++) {
    if (!isAttnLabelWord(words, index)) {
      continue;
    }
    const value = valueAfterAttn(words, index);
    if (value) {
      return value;
    }
  }
  return '';
}

function isAttnLabelWord(words: IOcrWord[], index: number): boolean {
  const key = wordKey(words[index].text || '');
  if (isAttnKey(key)) {
    return true;
  }
  if (key !== 'of') {
    return false;
  }
  const prev = words[index - 1];
  const prevKey = prev ? wordKey(prev.text || '') : '';
  return prevKey === 'attention';
}

function isAttnKey(key: string): boolean {
  return key === 'attn' ||
    key === 'atin' ||
    key === 'attm' ||
    key === 'atln' ||
    key === 'attention';
}

function valueAfterAttn(words: IOcrWord[], index: number): string {
  let start = index + 1;
  while (words[start]) {
    const raw = (words[start].text || '').trim();
    const key = wordKey(raw);
    if (/^[:.\-]+$/.test(raw) || key === 'of' || key === 'to') {
      start++;
      continue;
    }
    break;
  }

  const label = words[index];
  const lineMid = (label.y0 + label.y1) / 2;
  const lineHeight = Math.max(label.y1 - label.y0, 1);
  const collected: IOcrWord[] = [];
  for (let cursor = start; cursor < words.length; cursor++) {
    const word = words[cursor];
    const wordMid = (word.y0 + word.y1) / 2;
    if (Math.abs(wordMid - lineMid) > lineHeight * 0.85) {
      break;
    }
    if (word.x0 < label.x1 - 6) {
      continue;
    }
    if (isAttnValueStop(word.text || '')) {
      break;
    }
    collected.push(word);
  }
  const sameLine = joinOcrWords(collected).replace(/^[:.\s-]+/, '').trim();
  if (sameLine) {
    return sameLine;
  }

  const below: IOcrWord[] = [];
  for (let cursor = start; cursor < words.length; cursor++) {
    const word = words[cursor];
    if (word.y0 < label.y1 - lineHeight * 0.2) {
      continue;
    }
    if (word.y0 > label.y1 + lineHeight * 1.8) {
      break;
    }
    if (isAttnValueStop(word.text || '')) {
      break;
    }
    below.push(word);
  }
  return joinOcrWords(below).replace(/^[:.\s-]+/, '').trim();
}

function isAttnValueStop(text: string): boolean {
  const key = wordKey(text);
  return key === 'date' ||
    key === 'tel' ||
    key === 'fax' ||
    key === 'email' ||
    key === 'dear' ||
    key === 'our' ||
    key === 'your' ||
    key === 'page' ||
    isAttnKey(key);
}

function attnFromText(text: string): string {
  const lines = (text || '').split(/\r?\n/).map((line) => line.trim()).filter((line) => line.length > 0);
  const pattern = /(?:for\s+the\s+)?att(?:n|ention|in|m)\s*(?:of)?\s*[:.\-]?\s*(.+)$/i;
  for (let index = 0; index < lines.length; index++) {
    const match = lines[index].match(pattern);
    if (!match) {
      continue;
    }
    const value = (match[1] || '')
      .replace(/\b(date|tel|fax|email|dear|our\s+ref|your\s+ref)\b.*$/i, '')
      .replace(/^[:.\s-]+/, '')
      .trim();
    if (value) {
      return value;
    }
    const next = lines[index + 1];
    if (next && !isSalutationLine(next) && !isAddressBlockStop(next) && !/^att(?:n|ention)/i.test(next)) {
      return next.replace(/\s+/g, ' ').trim();
    }
  }
  return '';
}

function wordKey(text: string): string {
  return (text || '').toLowerCase().replace(/[^a-z0-9]+/g, '');
}

function addresseeBlockFromText(text: string): string {
  const raw = (text || '').split(/\r?\n/).map((line) => line.trim());
  let dearIndex = -1;
  for (let index = 0; index < raw.length; index++) {
    if (isSalutationLine(raw[index])) {
      dearIndex = index;
      break;
    }
  }
  if (dearIndex < 0) {
    const salutation = findSalutationIndex(text || '');
    if (salutation.index < 0) {
      return '';
    }
    const before = (text || '').substring(0, salutation.index).split(/\r?\n/).map((line) => line.trim());
    return firstLineOfConsecutiveBlock(before);
  }
  return firstLineOfConsecutiveBlock(raw.slice(0, dearIndex));
}

function addresseeBlockFromWords(page: IOcrPageResult): string {
  const hit = findSalutationHit(page.words || []);
  if (!hit) {
    return '';
  }
  const above = (page.words || []).filter((word) => {
    const midX = (word.x0 + word.x1) / 2;
    const midY = (word.y0 + word.y1) / 2;
    return midY < hit.y0 - 1 &&
      midY >= hit.y0 - Math.max(280, page.height * 0.45) &&
      midX <= page.width * 0.55;
  });
  const lines = groupWordsIntoLines(above).filter((line) => line.y1 < hit.y0);
  const block: string[] = [];
  let nextTop = hit.y0;
  for (let index = lines.length - 1; index >= 0; index--) {
    const line = lines[index];
    if (isSalutationLine(line.text)) {
      continue;
    }
    if (isAddressBlockStop(line.text)) {
      break;
    }
    const lineHeight = Math.max(12, line.y1 - line.y0);
    if (block.length > 0 && nextTop - line.y1 > lineHeight * 2.2) {
      break;
    }
    block.unshift(line.text);
    nextTop = line.y0;
    if (block.length >= 8) {
      break;
    }
  }
  return block.length > 0 ? block[0] : '';
}

function firstLineOfConsecutiveBlock(lines: string[]): string {
  const block: string[] = [];
  let sawGap = false;
  for (let index = lines.length - 1; index >= 0; index--) {
    const line = lines[index];
    if (!line) {
      if (block.length > 0) {
        break;
      }
      if (sawGap) {
        break;
      }
      sawGap = true;
      continue;
    }
    if (isSalutationLine(line)) {
      continue;
    }
    if (isAddressBlockStop(line)) {
      break;
    }
    block.unshift(line.replace(/\s+/g, ' ').trim());
    if (block.length >= 8) {
      break;
    }
  }
  return block.length > 0 ? block[0] : '';
}

function groupWordsIntoLines(words: IOcrWord[]): { text: string; y0: number; y1: number }[] {
  const sorted = words.slice().sort((left, right) => {
    if (Math.abs(left.y0 - right.y0) > 8) {
      return left.y0 - right.y0;
    }
    return left.x0 - right.x0;
  });
  const groups: IOcrWord[][] = [];
  sorted.forEach((word) => {
    const last = groups[groups.length - 1];
    if (!last) {
      groups.push([word]);
      return;
    }
    const lastMid = (last[0].y0 + last[0].y1) / 2;
    const wordMid = (word.y0 + word.y1) / 2;
    const lineHeight = Math.max(last[0].y1 - last[0].y0, 10);
    if (Math.abs(wordMid - lastMid) <= lineHeight * 0.6) {
      last.push(word);
      return;
    }
    groups.push([word]);
  });
  return groups.map((group) => {
    let y0 = group[0].y0;
    let y1 = group[0].y1;
    group.forEach((word) => {
      y0 = Math.min(y0, word.y0);
      y1 = Math.max(y1, word.y1);
    });
    return {
      text: joinOcrWords(group).replace(/\s+/g, ' ').trim(),
      y0,
      y1
    };
  }).filter((line) => line.text.length > 0);
}

async function subjectFromUnderline(page: IOcrPageResult): Promise<string> {
  if (!page.imageUrl) {
    return '';
  }

  try {
    const image = await loadImage(page.imageUrl);
    const canvas = document.createElement('canvas');
    canvas.width = image.width;
    canvas.height = image.height;
    const context = canvas.getContext('2d', { willReadFrequently: true });
    if (!context) {
      return '';
    }
    context.drawImage(image, 0, 0);

    const band = subjectScanBand(page, image.width, image.height);
    const x0 = band.x0;
    const x1 = band.x1;
    const y0 = band.y0;
    const y1 = band.y1;
    const width = x1 - x0;
    const height = y1 - y0;
    if (width < 24 || height < 8) {
      canvas.width = 0;
      canvas.height = 0;
      return '';
    }

    const pixels = context.getImageData(x0, y0, width, height).data;
    canvas.width = 0;
    canvas.height = 0;

    const rows: { y: number; left: number; right: number }[] = [];
    for (let row = 0; row < height; row++) {
      const span = longestDarkSpan(pixels, width, row, 140);
      if (!span) {
        continue;
      }
      const pageLeft = x0 + span.left;
      const pageRight = x0 + span.right;
      const spanWidth = pageRight - pageLeft;
      if (spanWidth < Math.max(36, image.width * 0.08) || spanWidth > image.width * 0.82) {
        continue;
      }
      if (pageLeft < image.width * 0.08 && spanWidth > image.width * 0.55) {
        continue;
      }
      rows.push({ y: y0 + row, left: pageLeft, right: pageRight });
    }

    const clustered: { y0: number; y1: number; left: number; right: number }[] = [];
    rows.forEach((row) => {
      const last = clustered[clustered.length - 1];
      if (!last || row.y - last.y1 > 2) {
        clustered.push({ y0: row.y, y1: row.y, left: row.left, right: row.right });
        return;
      }
      last.y1 = row.y;
      last.left = Math.min(last.left, row.left);
      last.right = Math.max(last.right, row.right);
    });

    const headings: string[] = [];
    let lastAcceptedY = 0;
    for (let index = 0; index < clustered.length; index++) {
      const cluster = clustered[index];
      if (cluster.y1 - cluster.y0 > 5) {
        continue;
      }
      const underline = { y: cluster.y1, left: cluster.left, right: cluster.right };
      const line = wordsOnUnderline(page, underline);
      if (!looksLikeSubjectHeading(line)) {
        if (headings.length > 0) {
          break;
        }
        continue;
      }
      if (headings.length > 0 && underline.y - lastAcceptedY > 40) {
        break;
      }
      if (headings.indexOf(line) < 0) {
        headings.push(line);
      }
      lastAcceptedY = underline.y;
      if (headings.length >= 3) {
        break;
      }
    }
    return headings.join(' ').trim();
  } catch {
    return '';
  }
}

function subjectScanBand(page: IOcrPageResult, width: number, height: number): ISignatureRegion {
  const hit = findSalutationHit(page.words || []);
  const closing = findClosingHit(page.words || []);
  const y0 = hit
    ? Math.min(height - 2, Math.floor(hit.y1 + 8))
    : Math.floor(height * 0.22);
  const y1 = Math.min(
    height,
    closing ? closing.y1 - 16 : height,
    y0 + Math.max(110, Math.round(height * 0.28))
  );
  return {
    x0: Math.floor(width * 0.06),
    y0,
    x1: Math.floor(width * 0.94),
    y1: Math.max(y0 + 8, y1)
  };
}

function longestDarkSpan(
  pixels: Uint8ClampedArray,
  width: number,
  row: number,
  maxLum: number
): { left: number; right: number } | undefined {
  let bestLeft = 0;
  let bestRight = -1;
  let runLeft = -1;
  let darkInRun = 0;
  const minRun = Math.max(24, Math.round(width * 0.08));

  const finishRun = (end: number) => {
    if (runLeft < 0) {
      return;
    }
    const runWidth = end - runLeft;
    const coverage = darkInRun / Math.max(1, runWidth);
    if (runWidth >= minRun && coverage > 0.34 && runWidth > bestRight - bestLeft) {
      bestLeft = runLeft;
      bestRight = end - 1;
    }
    runLeft = -1;
    darkInRun = 0;
  };

  for (let col = 0; col < width; col++) {
    const index = (row * width + col) * 4;
    const lum = 0.299 * pixels[index] + 0.587 * pixels[index + 1] + 0.114 * pixels[index + 2];
    if (lum < maxLum) {
      if (runLeft < 0) {
        runLeft = col;
        darkInRun = 0;
      }
      darkInRun++;
    } else if (runLeft >= 0 && col - runLeft - darkInRun > 6) {
      finishRun(col);
    }
  }
  finishRun(width);

  if (bestRight < bestLeft) {
    return undefined;
  }
  return { left: bestLeft, right: bestRight };
}

function wordsOnUnderline(
  page: IOcrPageResult,
  underline: { y: number; left: number; right: number }
): string {
  const pad = Math.max(6, (underline.right - underline.left) * 0.06);
  const words = (page.words || []).filter((word) => {
    const midX = (word.x0 + word.x1) / 2;
    return word.y1 <= underline.y + 4 &&
      word.y1 >= underline.y - 34 &&
      word.y0 < underline.y &&
      midX >= underline.left - pad &&
      midX <= underline.right + pad;
  });
  return stripSubjectLabel(joinOcrWords(words));
}

function looksLikeSubjectHeading(line: string): boolean {
  const trimmed = (line || '').trim();
  if (trimmed.length < 4 || trimmed.length > 140) {
    return false;
  }
  if (isSalutationLine(trimmed) || isClosingLine(trimmed) || isBodyStart(trimmed) || isAddressBlockStop(trimmed)) {
    return false;
  }
  return true;
}

function stripSubjectLabel(line: string): string {
  return (line || '')
    .replace(/^(re|subject|ref)\s*[:.\-]\s*/i, '')
    .replace(/\s+/g, ' ')
    .trim();
}

function isBodyStart(line: string): boolean {
  const trimmed = (line || '').trim();
  return trimmed.length > 140 ||
    /^(i|we|please|thank|further|with reference)\b/i.test(trimmed);
}

function findSalutationIndex(text: string): { index: number } {
  const match = (text || '').match(/\bdear\s+(s[il1]rs?|madams?|mesdames)\b/i) ||
    (text || '').match(/\bdear\s+sir\s*[/\\]?\s*madam\b/i);
  return { index: match && match.index !== undefined ? match.index : -1 };
}

function isSalutationLine(line: string): boolean {
  const key = normalizeKey(line);
  return /\bdear\s+(s[il1]rs?|madams?|mesdames)\b/.test(key) ||
    /\bdear\s+sir\s*madam/.test(key);
}

function isAddressBlockStop(line: string): boolean {
  const key = normalizeKey(line);
  return key.indexOf('our ref') === 0 ||
    key.indexOf('your ref') === 0 ||
    key.indexOf('by fax') === 0 ||
    key.indexOf('by email') === 0 ||
    key.indexOf('by post') === 0 ||
    /^date\b/.test(key) ||
    /^tel\b/.test(key) ||
    /^fax\b/.test(key) ||
    /^email\b/.test(key);
}

function findSalutationHit(words: IOcrWord[]): { y0: number; y1: number } | undefined {
  const sorted = words.slice().sort((left, right) => {
    if (Math.abs(left.y0 - right.y0) > 10) {
      return left.y0 - right.y0;
    }
    return left.x0 - right.x0;
  });
  for (let index = 0; index < sorted.length; index++) {
    if (!isDearWord(sorted[index].text || '')) {
      continue;
    }
    const lineMid = (sorted[index].y0 + sorted[index].y1) / 2;
    const nearby = sorted.slice(index, index + 6).filter((word) =>
      Math.abs((word.y0 + word.y1) / 2 - lineMid) < 16
    );
    const phrase = nearby.map((word) => word.text || '').join(' ');
    if (!isSalutationLine(phrase) && !nearby.some((word) => isSirWord(word.text || ''))) {
      continue;
    }
    let y0 = sorted[index].y0;
    let y1 = sorted[index].y1;
    nearby.forEach((word) => {
      y0 = Math.min(y0, word.y0);
      y1 = Math.max(y1, word.y1);
    });
    return { y0, y1 };
  }
  return undefined;
}

function isDearWord(text: string): boolean {
  const key = normalizeKey(text);
  return key === 'dear' || key === 'deor' || key === 'dcar' || key === 'dear,';
}

function isSirWord(text: string): boolean {
  const key = normalizeKey(text);
  return key === 'sir' ||
    key === 'sirs' ||
    key === 'sit' ||
    key === 'slr' ||
    key === 'madam' ||
    key === 'madams' ||
    key === 'mesdames';
}

function findInkSignatures(image: HTMLImageElement, words: IOcrWord[]): ISignatureRegion[] {
  const canvas = document.createElement('canvas');
  canvas.width = image.width;
  canvas.height = image.height;
  const context = canvas.getContext('2d', { willReadFrequently: true });
  if (!context) {
    return [];
  }
  context.drawImage(image, 0, 0);

  const x0 = Math.floor(image.width * 0.40);
  const y0 = Math.floor(image.height * 0.40);
  const x1 = image.width;
  const y1 = Math.floor(image.height * 0.97);
  const width = x1 - x0;
  const height = y1 - y0;
  if (width < 16 || height < 16) {
    return [];
  }

  const pixels = context.getImageData(x0, y0, width, height).data;
  canvas.width = 0;
  canvas.height = 0;

  const ink: number[] = [];
  const printed: boolean[] = [];
  for (let row = 0; row < height; row++) {
    let dark = 0;
    for (let col = 0; col < width; col++) {
      const index = (row * width + col) * 4;
      const lum = 0.299 * pixels[index] + 0.587 * pixels[index + 1] + 0.114 * pixels[index + 2];
      if (lum < 135) {
        dark++;
      }
    }
    ink.push(dark / width);
    printed.push(rowHasPrintedText(y0 + row, words, x0, x1));
  }

  const smooth: number[] = ink.map((_, row) => {
    const from = Math.max(0, row - 1);
    const to = Math.min(ink.length - 1, row + 1);
    let sum = 0;
    for (let index = from; index <= to; index++) {
      sum += ink[index];
    }
    return sum / (to - from + 1);
  });

  const minHeight = Math.max(10, Math.round(image.height * 0.016));
  const maxHeight = Math.max(minHeight + 1, Math.round(image.height * 0.14));
  const footerStart = Math.floor(image.height * 0.975);
  const regions: ISignatureRegion[] = [];

  for (let start = 0; start < smooth.length; start++) {
    if (printed[start] || smooth[start] < 0.028) {
      continue;
    }
    let end = start;
    while (
      end < smooth.length &&
      end - start < maxHeight &&
      !printed[end] &&
      smooth[end] >= 0.018
    ) {
      end++;
    }
    const runHeight = end - start;
    if (runHeight >= minHeight) {
      const top = y0 + start;
      const bottom = y0 + end;
      if (bottom < footerStart) {
        regions.push({
          x0,
          y0: top,
          x1,
          y1: bottom
        });
      }
    }
    start = Math.max(start, end - 1);
  }

  return regions;
}

function rowHasPrintedText(y: number, words: IOcrWord[], x0: number, x1: number): boolean {
  for (let index = 0; index < words.length; index++) {
    const word = words[index];
    const text = (word.text || '').trim();
    if (text.length < 2 || !/[A-Za-z\u3400-\u9FFF]{2,}/.test(text)) {
      continue;
    }
    const overlapsY = word.y0 <= y && word.y1 >= y;
    const overlapsX = word.x0 < x1 && word.x1 > x0;
    if (overlapsY && overlapsX) {
      return true;
    }
  }
  return false;
}

function wordsBelowSignature(page: IOcrPageResult, region: ISignatureRegion): IOcrWord[] {
  const y0 = region.y1 - 2;
  const y1 = Math.min(page.height, region.y1 + Math.max(96, page.height * 0.1));
  const x0 = Math.max(0, region.x0 - page.width * 0.06);
  return (page.words || []).filter((word) => {
    const midX = (word.x0 + word.x1) / 2;
    const midY = (word.y0 + word.y1) / 2;
    return midX >= x0 && midY >= y0 && midY <= y1;
  });
}

function extractNameBetweenClosingAndTitle(page?: IOcrPageResult): string {
  if (!page) {
    return '';
  }
  const fromWords = nameBetweenClosingAndTitleFromWords(page);
  if (fromWords) {
    return fromWords;
  }
  return nameBetweenClosingAndTitleFromText(page.text || '');
}

function nameBetweenClosingAndTitleFromWords(page: IOcrPageResult): string {
  const closing = findClosingHit(page.words || []);
  if (!closing) {
    return '';
  }
  const below = (page.words || []).filter((word) => {
    const midX = (word.x0 + word.x1) / 2;
    const midY = (word.y0 + word.y1) / 2;
    return midY > closing.y1 + 1 &&
      midY <= closing.y1 + Math.max(220, page.height * 0.34) &&
      midX >= closing.x0 - page.width * 0.22;
  });
  const lines = groupWordsIntoLines(below).map((line) => line.text);
  return pickNameAboveTitle(lines);
}

function nameBetweenClosingAndTitleFromText(text: string): string {
  const lines = (text || '')
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter((line) => line.length > 0);
  for (let index = 0; index < lines.length; index++) {
    if (!isClosingLine(lines[index])) {
      continue;
    }
    return pickNameAboveTitle(lines.slice(index + 1, index + 10));
  }
  return '';
}

function pickNameAboveTitle(lines: string[]): string {
  const zone: string[] = [];
  for (let index = 0; index < lines.length; index++) {
    const raw = (lines[index] || '').replace(/\s+/g, ' ').trim();
    if (!raw || isClosingLine(raw)) {
      continue;
    }
    if (isJobTitleLine(raw)) {
      break;
    }
    const cleaned = stripTrailingTitle(raw);
    if (!cleaned || isIgnorableBelowClosing(cleaned)) {
      continue;
    }
    zone.push(cleaned);
  }

  for (let index = 0; index < zone.length; index++) {
    const inner = firstParenthesesContent(zone[index]);
    if (inner) {
      return inner;
    }
  }
  for (let index = zone.length - 1; index >= 0; index--) {
    const name = personNameFromLine(zone[index]);
    if (name) {
      return name;
    }
  }
  return zone.length > 0 ? zone[zone.length - 1] : '';
}

function isJobTitleLine(line: string): boolean {
  const key = normalizeKey(line);
  if (!key) {
    return false;
  }
  if (/[\u3400-\u9FFF]*(工程師|總監|經理|主任|專員|顧問|秘書|署長|處長)/.test(line) &&
    !looksLikePersonName(line)) {
    return true;
  }
  const titles = [
    'director', 'manager', 'engineer', 'associate', 'consultant', 'officer',
    'secretary', 'architect', 'planner', 'surveyor', 'partner', 'chief',
    'assistant', 'principal', 'coordinator', 'specialist', 'supervisor',
    'technician', 'inspector', 'executive', 'president'
  ];
  const tokens = key.split(' ');
  for (let index = 0; index < titles.length; index++) {
    if (tokens.indexOf(titles[index]) >= 0) {
      return true;
    }
  }
  return false;
}

function stripTrailingTitle(line: string): string {
  const cut = (line || '').split(/,\s+(?=(?:ir|engr|eng|dr|mr|mrs|ms|prof)?\.?\s*(?:chief|director|manager|engineer|associate|consultant|officer|secretary|architect)\b)/i);
  return (cut[0] || line || '').replace(/\s+/g, ' ').trim();
}

function isIgnorableBelowClosing(line: string): boolean {
  const key = normalizeKey(line);
  return !key ||
    isIgnorableParen(key) ||
    key === 'signed' ||
    /^for and on behalf/.test(key) ||
    /^[-_.=]+$/.test(line);
}

function isClosingLine(line: string): boolean {
  const key = normalizeKey(line);
  return /(^|\s)yours\s+sincere/.test(' ' + key) ||
    /(^|\s)yours\s+faithful/.test(' ' + key) ||
    /(^|\s)yours\s+truly/.test(' ' + key);
}

function firstParenthesesContent(text: string): string {
  const matches = (text || '').match(/[\(\uFF08]\s*([^)\uFF09]{1,80}?)\s*[\)\uFF09]/g) || [];
  for (let index = 0; index < matches.length; index++) {
    const innerMatch = matches[index].match(/[\(\uFF08]\s*([^)\uFF09]{1,80}?)\s*[\)\uFF09]/);
    const inner = innerMatch ? innerMatch[1].replace(/\s+/g, ' ').trim() : '';
    if (inner && !isIgnorableParen(inner)) {
      return inner;
    }
  }
  return '';
}

function isIgnorableParen(value: string): boolean {
  const key = normalizeKey(value);
  return key === 'signed' ||
    key === 'signature' ||
    key === 'sgd' ||
    key === 'chop' ||
    key === 'seal';
}

function findClosingHit(words: IOcrWord[]): { x0: number; y1: number } | undefined {
  const sorted = words.slice().sort((left, right) => {
    if (Math.abs(left.y0 - right.y0) > 10) {
      return left.y0 - right.y0;
    }
    return left.x0 - right.x0;
  });
  for (let index = 0; index < sorted.length; index++) {
    const window = sorted.slice(index, index + 4);
    const lineMid = (sorted[index].y0 + sorted[index].y1) / 2;
    const sameLine = window.filter((word) => Math.abs((word.y0 + word.y1) / 2 - lineMid) < 12);
    const phrase = sameLine.map((word) => word.text || '').join(' ');
    if (!isClosingLine(phrase)) {
      continue;
    }
    let x0 = sameLine[0].x0;
    let y1 = sameLine[0].y1;
    sameLine.forEach((word) => {
      x0 = Math.min(x0, word.x0);
      y1 = Math.max(y1, word.y1);
    });
    return { x0, y1 };
  }
  return undefined;
}

function firstPersonName(text: string): string {
  const lines = (text || '')
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter((line) => line.length > 0);
  for (let index = 0; index < lines.length; index++) {
    const name = personNameFromLine(lines[index]);
    if (name) {
      return name;
    }
  }
  return '';
}

function personNameFromLine(line: string): string {
  const cleaned = (line || '')
    .replace(/^[\s(]+signed[\s)]+$/i, '')
    .replace(/^[-_.=]+$/, '')
    .trim();
  if (!cleaned || isNoiseLine(cleaned) || !looksLikePersonName(cleaned)) {
    return '';
  }
  return cleaned.replace(/\s+/g, ' ').trim();
}

function isNoiseLine(line: string): boolean {
  const key = normalizeKey(line);
  if (!key || /@|\d{3,}/.test(line) || /^https?:/i.test(line)) {
    return true;
  }
  const noise = [
    'tel', 'fax', 'email', 'www', 'http', 'page', 'enclosure', 'attachment',
    'limited', 'ltd', 'company', 'department', 'division', 'office',
    'director', 'manager', 'engineer', 'associate', 'consultant', 'officer',
    'secretary', 'architect', 'planner', 'surveyor', 'partner', 'chief',
    'yours', 'faithfully', 'sincerely', 'behalf'
  ];
  const tokens = key.split(' ');
  for (let index = 0; index < noise.length; index++) {
    if (tokens.indexOf(noise[index]) >= 0 && !hasHonorific(line)) {
      return true;
    }
  }
  return false;
}

function hasHonorific(line: string): boolean {
  return /^(ir|engr|eng|dr|mr|mrs|ms|prof)\b/i.test(line.trim());
}

function looksLikePersonName(line: string): boolean {
  const trimmed = line.trim();
  const cjk = trimmed.match(/[\u3400-\u9FFF]/g);
  if (cjk && cjk.length >= 2 && cjk.length <= 4) {
    return trimmed.replace(/[\u3400-\u9FFF\s.·]/g, '').length === 0;
  }

  const withoutTitle = trimmed.replace(/^(ir|engr|eng|dr|mr|mrs|ms|prof)\.?\s+/i, '');
  const tokens = withoutTitle.replace(/,/g, ' ').split(/\s+/).filter((token) => token.length > 0);
  if (tokens.length < 2 || tokens.length > 5) {
    return false;
  }

  return tokens.every((token) => {
    if (/^[A-Za-z]\.?$/.test(token) || /^[A-Za-z](?:\.[A-Za-z])+\.?$/.test(token)) {
      return true;
    }
    if (/^[A-Z][a-z]+(?:-[A-Z][a-z]+)?$/.test(token) || /^[A-Z]{2,12}$/.test(token)) {
      return true;
    }
    return false;
  });
}

function normalizeKey(text: string): string {
  return text.toLowerCase().replace(/[^a-z0-9]+/g, ' ').replace(/\s+/g, ' ').trim();
}

function loadImage(url: string): Promise<HTMLImageElement> {
  return new Promise((resolve, reject) => {
    const image = new Image();
    image.onload = () => resolve(image);
    image.onerror = () => reject(new Error('Unable to load page image.'));
    image.src = url;
  });
}
