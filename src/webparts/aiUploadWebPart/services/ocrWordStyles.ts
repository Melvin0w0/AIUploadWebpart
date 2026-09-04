import { IOcrWord } from './IPdfOcr';

/**
 * Convert-to-OCR underline sensitivity.
 * Edit this object only, then rebuild / refresh the workbench.
 *
 * More <u> tags  -> raise maxLum, lower minLineCoverage / wordCoverage / spanCoverage
 * Fewer <u> tags -> lower maxLum, raise minLineCoverage / wordCoverage / spanCoverage
 *
 * maxLum is 0-255. Higher = paler grey still counts as ink (more sensitive).
 */
export const UNDERLINE = {
  /** Ink darkness thresholds to try, darkest first. Typical 110-180. */
  maxLum: [90, 145],
  /** A line must have this much underline vs line width, or it is ignored. 0.15 = sensitive, 0.35 = strict. */
  minLineCoverage: 0.22,
  /** Extra absolute px floor for the line underline length. */
  minLinePx: 36,
  /** How far below the letters to look, as a fraction of line height. */
  lookBelow: 0.30,
  /** How far below one word to look for a per-word underline. */
  wordLookBelow: 0.28,
  /** Dark run must cover this fraction of the word width. 0.40 = sensitive, 0.70 = strict. */
  wordCoverage: 0.55,
  /** Word must overlap a detected underline span by this fraction of its width. */
  spanOverlap: 0.40,
  /** Dark run coverage inside a candidate span. 0.14 = sensitive, 0.30 = strict. */
  spanCoverage: 0.22,
  /** Merge nearby underline fragments. Larger = join a heading that has gaps between words. */
  mergeGapMin: 24,
  mergeGapMax: 48,
  mergeGapWordFactor: 1.5,
  /** Isolated underlined words shorter than this are dropped as noise. */
  dropIsolatedShorterThan: 4
};

export function annotateOcrWordStyles(words: IOcrWord[], image: ImageData): IOcrWord[] {
  if (!words || words.length === 0 || !image || image.width < 8 || image.height < 8) {
    return words;
  }
  const pixels = image.data;
  const width = image.width;
  const height = image.height;
  markUnderlinedWords(words, pixels, width, height);
  markBoldWords(words, pixels, width, height);
  return words;
}

function markUnderlinedWords(
  words: IOcrWord[],
  pixels: Uint8ClampedArray,
  width: number,
  height: number
): void {
  const lines = groupWordRows(words);
  const thresholds = UNDERLINE.maxLum;
  for (let lineIndex = 0; lineIndex < lines.length; lineIndex++) {
    const line = lines[lineIndex];
    const box = lineBox(line);
    const mergeGap = Math.max(
      UNDERLINE.mergeGapMin,
      Math.min(UNDERLINE.mergeGapMax, Math.round(medianWordGap(line) * UNDERLINE.mergeGapWordFactor))
    );
    let merged: { left: number; right: number }[] = [];
    for (let index = 0; index < thresholds.length; index++) {
      const found = underlineSpansUnderLine(box, pixels, width, height, thresholds[index], mergeGap);
      if (found.length > 0) {
        merged = found;
        break;
      }
    }

    for (let wordIndex = 0; wordIndex < line.length; wordIndex++) {
      const word = line[wordIndex];
      if (wordOverlapsSpans(word, merged) || wordHasUnderline(word, pixels, width, height, thresholds)) {
        word.underline = true;
      }
    }
    fillUnderlineGaps(line);
    dropIsolatedUnderlines(line);
  }
}

function fillUnderlineGaps(line: IOcrWord[]): void {
  for (let index = 1; index < line.length - 1; index++) {
    if (!line[index].underline && line[index - 1].underline && line[index + 1].underline) {
      line[index].underline = true;
    }
  }
}

function dropIsolatedUnderlines(line: IOcrWord[]): void {
  for (let index = 0; index < line.length; index++) {
    if (!line[index].underline) {
      continue;
    }
    const prev = index > 0 && line[index - 1].underline;
    const next = index < line.length - 1 && line[index + 1].underline;
    const letters = (line[index].text || '').replace(/[^A-Za-z0-9\u3400-\u9FFF]/g, '');
    if (!prev && !next && letters.length < UNDERLINE.dropIsolatedShorterThan) {
      line[index].underline = false;
    }
  }
}

function wordOverlapsSpans(word: IOcrWord, spans: { left: number; right: number }[]): boolean {
  const wordWidth = Math.max(4, word.x1 - word.x0);
  for (let index = 0; index < spans.length; index++) {
    const span = spans[index];
    const left = Math.max(span.left, word.x0);
    const right = Math.min(span.right, word.x1);
    if (right - left >= wordWidth * UNDERLINE.spanOverlap) {
      return true;
    }
  }
  return false;
}

function wordHasUnderline(
  word: IOcrWord,
  pixels: Uint8ClampedArray,
  width: number,
  height: number,
  thresholds: number[]
): boolean {
  const lineHeight = Math.max(8, word.y1 - word.y0);
  const row0 = Math.floor(word.y1);
  const row1 = Math.floor(word.y1 + Math.max(2, lineHeight * UNDERLINE.wordLookBelow));
  const col0 = Math.floor(word.x0);
  const col1 = Math.floor(word.x1);
  if (col1 - col0 < 8) {
    return false;
  }
  for (let t = 0; t < thresholds.length; t++) {
    if (rowBandHasUnderline(pixels, width, height, row0, row1, col0, col1, thresholds[t], UNDERLINE.wordCoverage)) {
      return true;
    }
  }
  return false;
}

function markBoldWords(
  words: IOcrWord[],
  pixels: Uint8ClampedArray,
  width: number,
  height: number
): void {
  const scored: { word: IOcrWord; density: number }[] = [];
  for (let index = 0; index < words.length; index++) {
    const word = words[index];
    if (!isStyleCandidate(word)) {
      continue;
    }
    scored.push({
      word,
      density: wordInkDensity(pixels, width, height, word, true)
    });
  }
  if (scored.length < 4) {
    return;
  }

  const densities = scored.map((item) => item.density).sort((left, right) => left - right);
  const pageMedian = densities[Math.floor(densities.length / 2)] || 0;
  const light = densities[Math.floor(densities.length * 0.15)] || 0;
  const heavy = densities[Math.floor(densities.length * 0.85)] || densities[densities.length - 1] || 0;
  if (heavy - light < 0.11) {
    return;
  }

  const lines = groupWordRows(words);
  for (let lineIndex = 0; lineIndex < lines.length; lineIndex++) {
    const line = lines[lineIndex];
    const lineScored: { word: IOcrWord; density: number }[] = [];
    for (let wordIndex = 0; wordIndex < line.length; wordIndex++) {
      const word = line[wordIndex];
      for (let scoredIndex = 0; scoredIndex < scored.length; scoredIndex++) {
        if (scored[scoredIndex].word === word) {
          lineScored.push(scored[scoredIndex]);
          break;
        }
      }
    }
    if (lineScored.length === 0) {
      continue;
    }
    const lineDensities = lineScored.map((item) => item.density).sort((left, right) => left - right);
    const lineMedian = lineDensities[Math.floor(lineDensities.length / 2)] || 0;
    const lineHeavy = lineDensities[lineDensities.length - 1] || 0;
    const wholeLineBold = !lineIsMostlyUnderlined(line) &&
      lineMedian >= pageMedian + 0.12 &&
      lineMedian >= pageMedian * 1.28;
    const cutoff = Math.max(pageMedian + 0.11, (lineMedian + lineHeavy) / 2, pageMedian * 1.28);
    for (let itemIndex = 0; itemIndex < lineScored.length; itemIndex++) {
      const item = lineScored[itemIndex];
      if (wholeLineBold || item.density >= cutoff) {
        item.word.bold = true;
      }
    }
  }

  let boldCount = 0;
  for (let index = 0; index < scored.length; index++) {
    if (scored[index].word.bold) {
      boldCount++;
    }
  }
  if (boldCount === 0 || boldCount / scored.length > 0.8) {
    for (let index = 0; index < scored.length; index++) {
      scored[index].word.bold = false;
    }
  }
}

function lineIsMostlyUnderlined(line: IOcrWord[]): boolean {
  if (line.length === 0) {
    return false;
  }
  let marked = 0;
  for (let index = 0; index < line.length; index++) {
    if (line[index].underline) {
      marked++;
    }
  }
  return marked / line.length >= 0.5;
}

function isStyleCandidate(word: IOcrWord): boolean {
  const text = (word.text || '').trim();
  if (!text || /^[\s.,;:!?()[\]'"“”‘’\-_/]+$/.test(text)) {
    return false;
  }
  if (word.x1 - word.x0 < 6 || word.y1 - word.y0 < 8) {
    return false;
  }
  return /[A-Za-z0-9\u3400-\u9FFF]/.test(text);
}

function groupWordRows(words: IOcrWord[]): IOcrWord[][] {
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
  return groups;
}

function lineBox(line: IOcrWord[]): { x0: number; x1: number; y0: number; y1: number } {
  let x0 = line[0].x0;
  let x1 = line[0].x1;
  let y0 = line[0].y0;
  let y1 = line[0].y1;
  for (let index = 1; index < line.length; index++) {
    const word = line[index];
    x0 = Math.min(x0, word.x0);
    x1 = Math.max(x1, word.x1);
    y0 = Math.min(y0, word.y0);
    y1 = Math.max(y1, word.y1);
  }
  return { x0, x1, y0, y1 };
}

function medianWordGap(line: IOcrWord[]): number {
  if (line.length < 2) {
    return 18;
  }
  const sorted = line.slice().sort((left, right) => left.x0 - right.x0);
  const gaps: number[] = [];
  for (let index = 1; index < sorted.length; index++) {
    gaps.push(Math.max(0, sorted[index].x0 - sorted[index - 1].x1));
  }
  gaps.sort((left, right) => left - right);
  return gaps[Math.floor(gaps.length / 2)] || 18;
}

function underlineSpansUnderLine(
  line: { x0: number; x1: number; y0: number; y1: number },
  pixels: Uint8ClampedArray,
  width: number,
  height: number,
  maxLum: number,
  mergeGap: number
): { left: number; right: number }[] {
  const lineHeight = Math.max(8, line.y1 - line.y0);
  const row0 = Math.floor(line.y1);
  const row1 = Math.floor(line.y1 + Math.max(2, lineHeight * UNDERLINE.lookBelow));
  const col0 = Math.floor(line.x0);
  const col1 = Math.floor(line.x1);
  const lineWidth = col1 - col0;
  if (lineWidth < 8) {
    return [];
  }

  const minFragment = Math.max(18, lineWidth * 0.08);
  let bestRow = -1;
  let bestScore = 0;
  for (let row = row0; row <= row1; row++) {
    if (row < 0 || row >= height) {
      continue;
    }
    const score = underlineScoreOnRow(pixels, width, row, col0, col1, maxLum, minFragment);
    if (score > bestScore) {
      bestScore = score;
      bestRow = row;
    }
  }
  const minLine = Math.max(UNDERLINE.minLinePx, lineWidth * UNDERLINE.minLineCoverage);
  if (bestRow < 0 || bestScore < minLine) {
    return [];
  }

  const collected: { left: number; right: number }[] = [];
  for (let row = bestRow - 1; row <= bestRow + 1; row++) {
    if (row < 0 || row >= height) {
      continue;
    }
    const spans = darkSpansOnRow(pixels, width, row, maxLum, 8);
    for (let spanIndex = 0; spanIndex < spans.length; spanIndex++) {
      const span = spans[spanIndex];
      const left = Math.max(col0 - 6, span.left);
      const right = Math.min(col1 + 6, span.right);
      if (right - left >= minFragment && (span.right - span.left) < width * 0.88) {
        collected.push({ left: span.left, right: span.right });
      }
    }
  }
  return mergeCloseSpans(collected, mergeGap);
}

function underlineScoreOnRow(
  pixels: Uint8ClampedArray,
  width: number,
  row: number,
  col0: number,
  col1: number,
  maxLum: number,
  minFragment: number
): number {
  const spans = darkSpansOnRow(pixels, width, row, maxLum, 8);
  let score = 0;
  for (let index = 0; index < spans.length; index++) {
    const left = Math.max(col0, spans[index].left);
    const right = Math.min(col1, spans[index].right);
    const overlap = right - left;
    if (overlap >= minFragment) {
      score += overlap;
    }
  }
  return score;
}

function rowBandHasUnderline(
  pixels: Uint8ClampedArray,
  width: number,
  height: number,
  row0: number,
  row1: number,
  col0: number,
  col1: number,
  maxLum: number,
  minCoverage: number
): boolean {
  const wordWidth = Math.max(1, col1 - col0);
  let hitRows = 0;
  for (let row = row0; row <= row1; row++) {
    if (row < 0 || row >= height) {
      continue;
    }
    let dark = 0;
    let run = 0;
    let bestRun = 0;
    for (let col = col0; col < col1; col++) {
      const index = (row * width + col) * 4;
      const lum = 0.299 * pixels[index] + 0.587 * pixels[index + 1] + 0.114 * pixels[index + 2];
      if (lum < maxLum) {
        dark++;
        run++;
        if (run > bestRun) {
          bestRun = run;
        }
      } else {
        run = 0;
      }
    }
    if (dark / wordWidth >= minCoverage && bestRun / wordWidth >= minCoverage * 0.8) {
      hitRows++;
    }
  }
  return hitRows >= 1 && hitRows <= 4;
}

function darkSpansOnRow(
  pixels: Uint8ClampedArray,
  width: number,
  row: number,
  maxLum: number,
  mergeGap: number
): { left: number; right: number }[] {
  const spans: { left: number; right: number }[] = [];
  const minRun = Math.max(8, Math.round(width * 0.01));
  let runLeft = -1;
  let darkInRun = 0;

  const finishRun = (end: number): void => {
    if (runLeft < 0) {
      return;
    }
    const runWidth = end - runLeft;
    const coverage = darkInRun / Math.max(1, runWidth);
    if (runWidth >= minRun && coverage > UNDERLINE.spanCoverage) {
      spans.push({ left: runLeft, right: end - 1 });
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
  return mergeCloseSpans(spans, mergeGap);
}

function mergeCloseSpans(
  spans: { left: number; right: number }[],
  maxGap: number
): { left: number; right: number }[] {
  if (spans.length === 0) {
    return spans;
  }
  const sorted = spans.slice().sort((left, right) => left.left - right.left);
  const merged: { left: number; right: number }[] = [sorted[0]];
  for (let index = 1; index < sorted.length; index++) {
    const last = merged[merged.length - 1];
    const next = sorted[index];
    if (next.left - last.right <= maxGap) {
      last.left = Math.min(last.left, next.left);
      last.right = Math.max(last.right, next.right);
    } else {
      merged.push({ left: next.left, right: next.right });
    }
  }
  return merged;
}

function wordInkDensity(
  pixels: Uint8ClampedArray,
  width: number,
  height: number,
  word: IOcrWord,
  excludeUnderlineBand: boolean
): number {
  const left = Math.max(0, Math.floor(word.x0));
  const top = Math.max(0, Math.floor(word.y0));
  const right = Math.min(width, Math.ceil(word.x1));
  const boxHeight = Math.max(1, word.y1 - word.y0);
  const bottomLimit = excludeUnderlineBand
    ? word.y1 - boxHeight * 0.28
    : word.y1;
  const bottom = Math.min(height, Math.ceil(bottomLimit));
  if (right - left < 2 || bottom - top < 2) {
    return 0;
  }
  let dark = 0;
  let total = 0;
  for (let row = top; row < bottom; row++) {
    for (let col = left; col < right; col++) {
      const index = (row * width + col) * 4;
      const lum = 0.299 * pixels[index] + 0.587 * pixels[index + 1] + 0.114 * pixels[index + 2];
      total++;
      if (lum < 145) {
        dark++;
      }
    }
  }
  return total > 0 ? dark / total : 0;
}
