import { IOcrPageResult, IOcrWord } from './IPdfOcr';
import { joinOcrWords } from './ocrSelection';

export async function trimOutgoingRefNoAtWatermark(
  pages: IOcrPageResult[] | undefined,
  refNo: string
): Promise<string> {
  const value = (refNo || '').trim();
  const list = pages || [];
  if (!value || list.length === 0) {
    return value;
  }

  for (let index = 0; index < list.length; index++) {
    const page = list[index];
    if (!page || !page.imageUrl || page.width < 8 || page.height < 8) {
      continue;
    }
    try {
      const trimmed = await trimRefOnPage(page, value);
      if (trimmed !== undefined) {
        return trimmed;
      }
    } catch {
      continue;
    }
  }
  return value;
}

async function trimRefOnPage(page: IOcrPageResult, value: string): Promise<string | undefined> {
  const words = locateRefValueWords(page, value);
  if (words.length === 0) {
    return undefined;
  }

  const image = await loadImage(page.imageUrl);
  const scaleX = page.width > 0 ? image.width / page.width : 1;
  const scaleY = page.height > 0 ? image.height / page.height : 1;
  const mapped = words.map((word) => scaleWord(word, scaleX, scaleY));
  const pixels = pagePixels(image);
  if (!pixels) {
    return value;
  }

  const cutX = watermarkCutX(pixels.data, pixels.width, pixels.height, mapped);
  if (cutX == null) {
    return value;
  }

  const kept = mapped.filter((word) => word.x0 < cutX - 2);
  if (kept.length === 0 || kept.length === mapped.length) {
    return value;
  }
  const trimmed = joinOcrWords(kept).replace(/^[:.\s-]+/, '').trim();
  return trimmed || value;
}

function locateRefValueWords(page: IOcrPageResult, value: string): IOcrWord[] {
  const sorted = sortWords(page.words || []);
  const matched = matchValueWords(sorted, value);
  if (matched.length > 0) {
    return matched;
  }
  const afterLabel = wordsAfterRefLabel(sorted);
  if (afterLabel.length === 0) {
    return [];
  }
  const joined = normalizeValue(joinOcrWords(afterLabel));
  const target = normalizeValue(value);
  if (!joined || !target) {
    return [];
  }
  if (joined.indexOf(target) >= 0 || target.indexOf(joined) >= 0) {
    return afterLabel;
  }
  return [];
}

function matchValueWords(words: IOcrWord[], value: string): IOcrWord[] {
  const target = normalizeValue(value);
  if (!target) {
    return [];
  }
  const items: { word: IOcrWord; norm: string }[] = [];
  for (let index = 0; index < words.length; index++) {
    const norm = normalizeValue(words[index].text || '');
    if (norm) {
      items.push({ word: words[index], norm });
    }
  }
  let contains: IOcrWord[] = [];
  for (let start = 0; start < items.length; start++) {
    let joined = '';
    const collected: IOcrWord[] = [];
    for (let end = start; end < items.length && end < start + 32; end++) {
      joined = joined ? (joined + ' ' + items[end].norm) : items[end].norm;
      collected.push(items[end].word);
      const compact = joined.replace(/\s+/g, '');
      const compactTarget = target.replace(/\s+/g, '');
      if (joined === target || compact === compactTarget) {
        return collected;
      }
      if ((joined.indexOf(target) >= 0 || compact.indexOf(compactTarget) >= 0) &&
        joined.length <= target.length + 24) {
        contains = collected.slice();
      }
      if (joined.length > target.length + 24 && compact.indexOf(compactTarget) < 0) {
        break;
      }
    }
  }
  return contains;
}

function wordsAfterRefLabel(words: IOcrWord[]): IOcrWord[] {
  for (let index = 0; index < words.length; index++) {
    const refIndex = ourOrPlainRefIndex(words, index);
    if (refIndex < 0) {
      continue;
    }
    const collected = collectSameLineValue(words, refIndex);
    if (collected.length > 0) {
      return collected;
    }
  }
  return [];
}

function ourOrPlainRefIndex(words: IOcrWord[], index: number): number {
  const key = normalizeToken(words[index].text || '');
  if (/^ourr+e+fs?(no|number)?$/.test(key) || key === 'ourreference') {
    return index;
  }
  if (!isRefWord(key)) {
    return -1;
  }
  const prev = words[index - 1];
  const prevKey = prev ? normalizeToken(prev.text || '') : '';
  if (prevKey === 'our' || prevKey === '0ur' || prevKey === 'ou') {
    return index;
  }
  if (prevKey === 'your' || prevKey === 'you' || prevKey === 'yr' || prevKey === 'youf' || prevKey === 'yor') {
    return -1;
  }
  if (/^y(ou)?rr+e+fs?(no|number)?$/.test(key)) {
    return -1;
  }
  if ((words[index].text || '').indexOf(':') >= 0) {
    return index;
  }
  return -1;
}

function collectSameLineValue(words: IOcrWord[], index: number): IOcrWord[] {
  let start = index + 1;
  if (words[start] && /^[:.-]+$/.test((words[start].text || '').trim())) {
    start++;
  }
  const nextKey = words[start] ? normalizeToken(words[start].text || '') : '';
  if (nextKey === 'no' || nextKey === 'number') {
    start++;
    if (words[start] && /^[:.-]+$/.test((words[start].text || '').trim())) {
      start++;
    }
  }
  const label = words[index];
  const lineMid = (label.y0 + label.y1) / 2;
  const lineHeight = Math.max(label.y1 - label.y0, 1);
  const collected: IOcrWord[] = [];
  for (let cursor = start; cursor < words.length; cursor++) {
    const word = words[cursor];
    const wordMid = (word.y0 + word.y1) / 2;
    if (Math.abs(wordMid - lineMid) > lineHeight * 0.8) {
      break;
    }
    if (word.x0 < label.x1 - 6) {
      continue;
    }
    if (isRefValueStop(word.text || '')) {
      break;
    }
    collected.push(word);
  }
  return collected;
}

function watermarkCutX(
  pixels: Uint8ClampedArray,
  width: number,
  height: number,
  words: IOcrWord[]
): number | undefined {
  const sorted = sortWords(words);
  if (sorted.length === 0) {
    return undefined;
  }
  const printed = inkProfile(pixels, width, height, sorted.slice(0, Math.min(2, sorted.length)));
  const lineH = medianHeight(sorted);

  for (let index = 1; index < sorted.length; index++) {
    const previous = sorted[index - 1];
    const word = sorted[index];
    const gap = word.x0 - previous.x1;
    if (gap > Math.max(28, lineH * 1.4) &&
      regionLooksLikeWatermark(pixels, width, height, previous.x1 + 2, word.x1, wordBand(sorted, lineH), printed, lineH)) {
      return word.x0;
    }
    if (wordLooksLikeWatermark(pixels, width, height, word, printed, lineH)) {
      return word.x0;
    }
  }

  const last = sorted[sorted.length - 1];
  const right = Math.min(width, last.x1 + Math.max(160, lineH * 8));
  if (regionLooksLikeWatermark(
    pixels,
    width,
    height,
    last.x1 + Math.max(8, lineH * 0.4),
    right,
    wordBand(sorted, lineH),
    printed,
    lineH
  )) {
    return last.x1 + Math.max(4, lineH * 0.2);
  }
  return undefined;
}

function wordLooksLikeWatermark(
  pixels: Uint8ClampedArray,
  width: number,
  height: number,
  word: IOcrWord,
  printed: IInkProfile,
  lineH: number
): boolean {
  const wordH = Math.max(1, word.y1 - word.y0);
  if (wordH > lineH * 1.7) {
    return true;
  }
  const band = {
    y0: word.y0 - Math.max(4, wordH * 0.35),
    y1: word.y1 + Math.max(4, wordH * 0.35)
  };
  if (regionLooksLikeWatermark(pixels, width, height, word.x0, word.x1, band, printed, lineH)) {
    return true;
  }
  if (isRefLikeToken(word.text || '')) {
    return false;
  }
  const stats = sampleRegion(pixels, width, height, word.x0, band.y0, word.x1, band.y1);
  return stats.faint > 0.08 && stats.faintMedian > printed.medianLum + 16;
}

function regionLooksLikeWatermark(
  pixels: Uint8ClampedArray,
  width: number,
  height: number,
  x0: number,
  x1: number,
  band: { y0: number; y1: number },
  printed: IInkProfile,
  lineH: number
): boolean {
  const left = Math.max(0, Math.floor(Math.min(x0, x1)));
  const right = Math.min(width, Math.ceil(Math.max(x0, x1)));
  if (right - left < 14) {
    return false;
  }
  const stats = sampleRegion(pixels, width, height, left, band.y0, right, band.y1);
  if (stats.total < 80) {
    return false;
  }
  const inkSpan = stats.inkBottom - stats.inkTop;
  if (stats.colored > 0.06 && stats.colored * (right - left) > 18) {
    return true;
  }
  if (stats.faint > 0.10 && stats.faint > stats.printed * 1.15 && stats.faintMedian > printed.medianLum + 20) {
    return true;
  }
  if (inkSpan > lineH * 2.05 && (stats.faint + stats.colored) > 0.05) {
    return true;
  }
  return false;
}

interface IInkProfile {
  medianLum: number;
}

interface IRegionStats {
  total: number;
  printed: number;
  faint: number;
  colored: number;
  faintMedian: number;
  inkTop: number;
  inkBottom: number;
}

function inkProfile(
  pixels: Uint8ClampedArray,
  width: number,
  height: number,
  words: IOcrWord[]
): IInkProfile {
  const lums: number[] = [];
  for (let index = 0; index < words.length; index++) {
    const word = words[index];
    const left = Math.max(0, Math.floor(word.x0));
    const right = Math.min(width, Math.ceil(word.x1));
    const top = Math.max(0, Math.floor(word.y0));
    const bottom = Math.min(height, Math.ceil(word.y1));
    for (let row = top; row < bottom; row += 2) {
      for (let col = left; col < right; col += 2) {
        const pixel = (row * width + col) * 4;
        const lum = luminance(pixels[pixel], pixels[pixel + 1], pixels[pixel + 2]);
        if (lum < 130) {
          lums.push(lum);
        }
      }
    }
  }
  return {
    medianLum: median(lums, 55)
  };
}

function sampleRegion(
  pixels: Uint8ClampedArray,
  width: number,
  height: number,
  x0: number,
  y0: number,
  x1: number,
  y1: number
): IRegionStats {
  const left = Math.max(0, Math.floor(x0));
  const right = Math.min(width, Math.ceil(x1));
  const top = Math.max(0, Math.floor(y0));
  const bottom = Math.min(height, Math.ceil(y1));
  let total = 0;
  let printed = 0;
  let faint = 0;
  let colored = 0;
  const faintLums: number[] = [];
  let inkTop = bottom;
  let inkBottom = top;
  for (let row = top; row < bottom; row += 2) {
    for (let col = left; col < right; col += 2) {
      const pixel = (row * width + col) * 4;
      const r = pixels[pixel];
      const g = pixels[pixel + 1];
      const b = pixels[pixel + 2];
      const lum = luminance(r, g, b);
      total++;
      const chroma = Math.max(r, g, b) - Math.min(r, g, b);
      if (lum < 85) {
        printed++;
        inkTop = Math.min(inkTop, row);
        inkBottom = Math.max(inkBottom, row);
      } else if (lum <= 198) {
        if (chroma > 42) {
          colored++;
          inkTop = Math.min(inkTop, row);
          inkBottom = Math.max(inkBottom, row);
        } else if (lum >= 88) {
          faint++;
          faintLums.push(lum);
          inkTop = Math.min(inkTop, row);
          inkBottom = Math.max(inkBottom, row);
        }
      }
    }
  }
  return {
    total,
    printed: total > 0 ? printed / total : 0,
    faint: total > 0 ? faint / total : 0,
    colored: total > 0 ? colored / total : 0,
    faintMedian: median(faintLums, 160),
    inkTop: inkTop === bottom ? top : inkTop,
    inkBottom: inkBottom === top ? top : inkBottom
  };
}

function wordBand(words: IOcrWord[], lineH: number): { y0: number; y1: number } {
  let y0 = words[0].y0;
  let y1 = words[0].y1;
  for (let index = 1; index < words.length; index++) {
    y0 = Math.min(y0, words[index].y0);
    y1 = Math.max(y1, words[index].y1);
  }
  const pad = Math.max(6, lineH * 1.15);
  return { y0: y0 - pad, y1: y1 + pad };
}

function isRefLikeToken(text: string): boolean {
  const token = (text || '').trim();
  if (!token) {
    return false;
  }
  if (/^[:./\\-]+$/.test(token)) {
    return true;
  }
  if (/[0-9]/.test(token) && token.length <= 24) {
    return true;
  }
  return /^[A-Za-z]{1,5}$/.test(token);
}

function isRefWord(key: string): boolean {
  return key === 'reference' || key === 'refno' || /^r+e+fs?(no|number)?$/.test(key);
}

function isRefValueStop(text: string): boolean {
  const key = normalizeToken(text);
  return key === 'date' ||
    key === 'tel' ||
    key === 'fax' ||
    key === 'email' ||
    key === 'our' ||
    key === 'your' ||
    key === 'yr' ||
    key === 'page' ||
    key === 'dear' ||
    /^(our|your|yr|my)?r+efs?(no|number)?$/.test(key) ||
    key === 'reference' ||
    key === 'referenceno';
}

function sortWords(words: IOcrWord[]): IOcrWord[] {
  return (words || []).slice().sort((left, right) => {
    const lineHeight = Math.max(left.y1 - left.y0, right.y1 - right.y0, 1);
    if (Math.abs(left.y0 - right.y0) > lineHeight * 0.5) {
      return left.y0 - right.y0;
    }
    return left.x0 - right.x0;
  });
}

function scaleWord(word: IOcrWord, scaleX: number, scaleY: number): IOcrWord {
  return {
    text: word.text,
    x0: word.x0 * scaleX,
    y0: word.y0 * scaleY,
    x1: word.x1 * scaleX,
    y1: word.y1 * scaleY
  };
}

function normalizeValue(value: string): string {
  return (value || '')
    .toLowerCase()
    .replace(/[^a-z0-9\u4e00-\u9fff]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function normalizeToken(text: string): string {
  return (text || '').toLowerCase().replace(/[^a-z0-9]+/g, '');
}

function luminance(r: number, g: number, b: number): number {
  return 0.299 * r + 0.587 * g + 0.114 * b;
}

function medianHeight(words: IOcrWord[]): number {
  const heights = words.map((word) => Math.max(1, word.y1 - word.y0)).sort((left, right) => left - right);
  return median(heights, 12);
}

function median(values: number[], fallback: number): number {
  if (!values || values.length === 0) {
    return fallback;
  }
  const sorted = values.slice().sort((left, right) => left - right);
  const mid = Math.floor(sorted.length / 2);
  if (sorted.length % 2 === 1) {
    return sorted[mid];
  }
  return (sorted[mid - 1] + sorted[mid]) / 2;
}

function pagePixels(image: HTMLImageElement): { data: Uint8ClampedArray; width: number; height: number } | undefined {
  const canvas = document.createElement('canvas');
  canvas.width = image.width;
  canvas.height = image.height;
  const context = canvas.getContext('2d', { willReadFrequently: true });
  if (!context) {
    canvas.width = 0;
    canvas.height = 0;
    return undefined;
  }
  context.drawImage(image, 0, 0);
  const imageData = context.getImageData(0, 0, image.width, image.height);
  canvas.width = 0;
  canvas.height = 0;
  return {
    data: imageData.data,
    width: image.width,
    height: image.height
  };
}

function loadImage(url: string): Promise<HTMLImageElement> {
  return new Promise((resolve, reject) => {
    const image = new Image();
    image.onload = () => resolve(image);
    image.onerror = () => reject(new Error('Unable to load page image.'));
    image.src = url;
  });
}
