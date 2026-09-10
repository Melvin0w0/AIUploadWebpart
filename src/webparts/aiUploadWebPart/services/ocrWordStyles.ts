import { IOcrStyleSpan, IOcrWord } from './IPdfOcr';

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
  /** Ink darkness 0-255. A pixel darker than this counts as underline ink. Black ink is usually 0-50, so 110 vs 190 often looks the same on real underlines. Typical 110-160. */
  maxLum: [245, 248],
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
  dropIsolatedShorterThan: 4,
  /**
   * Grid / table rules vs letter underlines.
   * A dark span that overshoots the letters this many px on both sides is a divider, not <u>.
   */
  separatorOvershootPx: 20,
  /** A dark span this many times wider than the text cluster is a column/table rule. */
  separatorWidthRatio: 1.5,
  /** A dark span covering this fraction of the page width is a page rule, not an underline. */
  pageRuleRatio: 0.78,
  /** Letter underlines sit this close to the baseline. Rules further down are table / separators. */
  tightLookBelow: 0.16
};

/**
 * Subject line joining after Dear Sir/Madam.
 * Tune maxLineGap only, then Convert again.
 *
 * Gap is in pixels: next line top minus previous line bottom.
 * Wrapped heading (join with a space) -> 24-48
 * One line only -> 0
 */
export const SUBJECT = {
  maxLineGap: 20
};

/**
 * Convert-to-OCR bold sensitivity. Tune extraOverMedian only, then Convert again.
 * More <b> -> 0.04   Fewer <b> -> 0.12   Default 0.06
 */
export const BOLD = {
  /** Pixel darker than this counts as ink inside a letter. Typical 130-160. */
  inkLum: 145,
  /**
   * Ignore this fraction of the word bottom so an underline is not counted as extra ink.
   * Use 0.20-0.35 only. This is NOT the bold sensitivity knob.
   */
  excludeBottom: 0.28,
  /**
   * How much heavier than typical page text a word must be to get <b>.
   * Lower = more <b> (0.04). Higher = fewer (0.12). Start at 0.06.
   */
  extraOverMedian: 0.02
};

function pushStyleSpan(
  debug: IOcrStyleSpan[] | undefined,
  kind: IOcrStyleSpan['kind'],
  left: number,
  right: number,
  y: number
): void {
  if (!debug) {
    return;
  }
  debug.push({
    kind,
    x0: left,
    x1: right,
    y
  });
}

export function annotateOcrWordStyles(
  words: IOcrWord[],
  image: ImageData,
  debugSpans?: IOcrStyleSpan[]
): IOcrWord[] {
  if (!words || words.length === 0 || !image || image.width < 8 || image.height < 8) {
    return words;
  }
  const pixels = image.data;
  const width = image.width;
  const height = image.height;
  markUnderlinedWords(words, pixels, width, height, debugSpans);
  markBoldWords(words, pixels, width, height);
  return words;
}

function markUnderlinedWords(
  words: IOcrWord[],
  pixels: Uint8ClampedArray,
  width: number,
  height: number,
  debugSpans?: IOcrStyleSpan[]
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
    if (lineLooksLikeTableRow(box, pixels, width, height, thresholds[0], debugSpans)) {
      continue;
    }
    let merged: { left: number; right: number }[] = [];
    let usedLum = thresholds[0];
    for (let index = 0; index < thresholds.length; index++) {
      const collected: IOcrStyleSpan[] = [];
      const found = underlineSpansUnderLine(
        box,
        pixels,
        width,
        height,
        thresholds[index],
        mergeGap,
        collected
      );
      if (found.length > 0) {
        merged = found;
        usedLum = thresholds[index];
        if (debugSpans) {
          for (let spanIndex = 0; spanIndex < collected.length; spanIndex++) {
            debugSpans.push(collected[spanIndex]);
          }
        }
        break;
      }
    }
    const wordThresholds = merged.length > 0 ? [usedLum] : thresholds;

    for (let wordIndex = 0; wordIndex < line.length; wordIndex++) {
      const word = line[wordIndex];
      if (
        wordOverlapsSpans(word, merged) ||
        wordHasUnderline(word, box, pixels, width, height, wordThresholds)
      ) {
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
  line: { x0: number; x1: number },
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
    if (
      rowBandHasUnderline(pixels, width, height, row0, row1, col0, col1, thresholds[t], UNDERLINE.wordCoverage) &&
      !wordSitsOnGridRule(word, line, pixels, width, height, row0, row1, thresholds[t])
    ) {
      return true;
    }
  }
  return false;
}

function wordSitsOnGridRule(
  word: IOcrWord,
  line: { x0: number; x1: number },
  pixels: Uint8ClampedArray,
  width: number,
  height: number,
  row0: number,
  row1: number,
  maxLum: number
): boolean {
  for (let row = row0; row <= row1; row++) {
    if (row < 0 || row >= height) {
      continue;
    }
    const spans = darkSpansOnRow(pixels, width, row, maxLum, 8);
    for (let index = 0; index < spans.length; index++) {
      const span = spans[index];
      const overlap = Math.min(word.x1, span.right) - Math.max(word.x0, span.left);
      if (overlap < (word.x1 - word.x0) * 0.35) {
        continue;
      }
      if (isGridRuleSpan(span, line.x0, line.x1, width)) {
        return true;
      }
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
  const cutoff = pageMedian + BOLD.extraOverMedian;
  for (let index = 0; index < scored.length; index++) {
    if (scored[index].density >= cutoff) {
      scored[index].word.bold = true;
    }
  }
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

/**
 * True when a dark horizontal span is a table / page divider rather than ink under letters.
 * Text underlines stay close to the glyph width; grid rules keep going past the words.
 */
export function isGridRuleSpan(
  span: { left: number; right: number },
  textLeft: number,
  textRight: number,
  pageWidth: number
): boolean {
  const textWidth = Math.max(1, textRight - textLeft);
  const spanWidth = span.right - span.left + 1;
  if (spanWidth >= pageWidth * UNDERLINE.pageRuleRatio) {
    return true;
  }
  const overshootLeft = textLeft - span.left;
  const overshootRight = span.right - textRight;
  if (
    spanWidth > textWidth * UNDERLINE.separatorWidthRatio &&
    overshootLeft > UNDERLINE.separatorOvershootPx &&
    overshootRight > UNDERLINE.separatorOvershootPx
  ) {
    return true;
  }
  if (
    spanWidth > textWidth * 2 &&
    (overshootLeft > 48 || overshootRight > 48)
  ) {
    return true;
  }
  return false;
}

function lineLooksLikeTableRow(
  line: { x0: number; x1: number; y0: number; y1: number },
  pixels: Uint8ClampedArray,
  width: number,
  height: number,
  maxLum: number,
  debugSpans?: IOcrStyleSpan[]
): boolean {
  const lineHeight = Math.max(8, line.y1 - line.y0);
  const below0 = Math.floor(line.y1);
  const below1 = Math.floor(line.y1 + Math.max(3, lineHeight * 0.55));
  const above0 = Math.floor(line.y0 - Math.max(3, lineHeight * 0.55));
  const above1 = Math.floor(line.y0 - 1);
  const below = bestOverlappingSpanInBand(pixels, width, height, below0, below1, line.x0, line.x1, maxLum);
  const above = bestOverlappingSpanInBand(pixels, width, height, above0, above1, line.x0, line.x1, maxLum);
  if (below && isGridRuleSpan(below, line.x0, line.x1, width)) {
    pushStyleSpan(debugSpans, 'separator', below.left, below.right, below.row);
    return true;
  }
  if (below && spanHasVerticalJoins(pixels, width, height, below, maxLum)) {
    pushStyleSpan(debugSpans, 'separator', below.left, below.right, below.row);
    return true;
  }
  if (above && below) {
    const leftDiff = Math.abs(above.left - below.left);
    const rightDiff = Math.abs(above.right - below.right);
    if (leftDiff <= 18 && rightDiff <= 18 && (isGridRuleSpan(above, line.x0, line.x1, width) || spanHasVerticalJoins(pixels, width, height, above, maxLum))) {
      pushStyleSpan(debugSpans, 'separator', below.left, below.right, below.row);
      pushStyleSpan(debugSpans, 'separator', above.left, above.right, above.row);
      return true;
    }
  }
  return false;
}

function bestOverlappingSpanInBand(
  pixels: Uint8ClampedArray,
  width: number,
  height: number,
  row0: number,
  row1: number,
  textLeft: number,
  textRight: number,
  maxLum: number
): { left: number; right: number; row: number } | undefined {
  const textWidth = Math.max(1, textRight - textLeft);
  let best: { left: number; right: number; row: number; overlap: number } | undefined;
  for (let row = row0; row <= row1; row++) {
    if (row < 0 || row >= height) {
      continue;
    }
    const spans = darkSpansOnRow(pixels, width, row, maxLum, 8);
    for (let index = 0; index < spans.length; index++) {
      const span = spans[index];
      const overlap = Math.min(textRight, span.right) - Math.max(textLeft, span.left);
      if (overlap < textWidth * UNDERLINE.minLineCoverage) {
        continue;
      }
      if (!best || overlap > best.overlap) {
        best = { left: span.left, right: span.right, row, overlap };
      }
    }
  }
  return best;
}

function spanHasVerticalJoins(
  pixels: Uint8ClampedArray,
  width: number,
  height: number,
  span: { left: number; right: number; row: number },
  maxLum: number
): boolean {
  const probe = 10;
  const columnHasJoin = (col: number): boolean => {
    let hits = 0;
    for (let delta = 1; delta <= probe; delta++) {
      if (span.row - delta >= 0 && isDarkPixel(pixels, width, span.row - delta, col, maxLum)) {
        hits++;
      }
      if (span.row + delta < height && isDarkPixel(pixels, width, span.row + delta, col, maxLum)) {
        hits++;
      }
    }
    return hits >= 5;
  };
  return columnHasJoin(span.left) && columnHasJoin(span.right);
}

function isDarkPixel(
  pixels: Uint8ClampedArray,
  width: number,
  row: number,
  col: number,
  maxLum: number
): boolean {
  if (col < 0 || col >= width) {
    return false;
  }
  const index = (row * width + col) * 4;
  const lum = 0.299 * pixels[index] + 0.587 * pixels[index + 1] + 0.114 * pixels[index + 2];
  return lum < maxLum;
}

function underlineSpansUnderLine(
  line: { x0: number; x1: number; y0: number; y1: number },
  pixels: Uint8ClampedArray,
  width: number,
  height: number,
  maxLum: number,
  mergeGap: number,
  debugSpans?: IOcrStyleSpan[]
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
      if (isGridRuleSpan(span, col0, col1, width)) {
        pushStyleSpan(debugSpans, 'separator', span.left, span.right, row);
        continue;
      }
      if (right - left >= minFragment) {
        collected.push({ left: span.left, right: span.right });
      }
    }
  }
  const merged = mergeCloseSpans(collected, mergeGap);
  for (let index = 0; index < merged.length; index++) {
    pushStyleSpan(debugSpans, 'underline', merged[index].left, merged[index].right, bestRow);
  }
  return merged;
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
    ? word.y1 - boxHeight * BOLD.excludeBottom
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
      if (lum < BOLD.inkLum) {
        dark++;
      }
    }
  }
  return total > 0 ? dark / total : 0;
}
