import { IOcrWord } from './IPdfOcr';

export interface ISelectionRect {
  x0: number;
  y0: number;
  x1: number;
  y1: number;
}

export function joinOcrWords(words: IOcrWord[]): string {
  if (!words || words.length === 0) {
    return '';
  }

  const sorted = words.slice().sort((left, right) => {
    const lineHeight = Math.max(left.y1 - left.y0, right.y1 - right.y0, 1);
    if (Math.abs(left.y0 - right.y0) > lineHeight * 0.5) {
      return left.y0 - right.y0;
    }
    return left.x0 - right.x0;
  });

  let text = wordText(sorted[0]);
  for (let index = 1; index < sorted.length; index++) {
    const previous = sorted[index - 1];
    const current = sorted[index];
    const lineHeight = Math.max(previous.y1 - previous.y0, 1);
    const isNewLine = current.y0 - previous.y0 > lineHeight * 0.6;
    text += isNewLine ? '\n' : joinGap(wordText(previous), wordText(current));
    text += wordText(current);
  }

  return text.trim();
}

function wordText(word: IOcrWord): string {
  return word && word.text ? word.text : '';
}

export function wordsInRect(words: IOcrWord[], rect: ISelectionRect): number[] {
  const x0 = Math.min(rect.x0, rect.x1);
  const y0 = Math.min(rect.y0, rect.y1);
  const x1 = Math.max(rect.x0, rect.x1);
  const y1 = Math.max(rect.y0, rect.y1);

  const indexes: number[] = [];
  words.forEach((word, index) => {
    const overlaps = word.x0 < x1 && word.x1 > x0 && word.y0 < y1 && word.y1 > y0;
    if (overlaps) {
      indexes.push(index);
    }
  });
  return indexes;
}

export function wordsInTextRange(
  words: IOcrWord[],
  startX: number,
  startY: number,
  endX: number,
  endY: number
): number[] {
  const lines = groupIndexedLines(words);
  if (lines.length === 0) {
    return [];
  }

  const start = locateLinePoint(lines, startX, startY);
  const end = locateLinePoint(lines, endX, endY);
  if (!start || !end) {
    return [];
  }

  let from = start;
  let to = end;
  if (start.line > end.line || (start.line === end.line && start.x > end.x)) {
    from = end;
    to = start;
  }

  const indexes: number[] = [];
  for (let lineIndex = from.line; lineIndex <= to.line; lineIndex++) {
    const line = lines[lineIndex];
    line.items.forEach((item) => {
      const word = item.word;
      if (from.line === to.line) {
        const left = Math.min(from.x, to.x);
        const right = Math.max(from.x, to.x);
        if (word.x1 >= left && word.x0 <= right) {
          indexes.push(item.index);
        }
        return;
      }
      if (lineIndex === from.line) {
        if (word.x1 >= from.x) {
          indexes.push(item.index);
        }
        return;
      }
      if (lineIndex === to.line) {
        if (word.x0 <= to.x) {
          indexes.push(item.index);
        }
        return;
      }
      indexes.push(item.index);
    });
  }
  return indexes;
}

interface IIndexedLine {
  items: { word: IOcrWord; index: number }[];
  y0: number;
  y1: number;
}

function groupIndexedLines(words: IOcrWord[]): IIndexedLine[] {
  const indexed = (words || [])
    .map((word, index) => ({ word, index }))
    .filter((item) => !!(item.word && (item.word.text || '').trim()))
    .sort((left, right) => {
      const lineHeight = Math.max(left.word.y1 - left.word.y0, right.word.y1 - right.word.y0, 1);
      if (Math.abs(left.word.y0 - right.word.y0) > lineHeight * 0.5) {
        return left.word.y0 - right.word.y0;
      }
      return left.word.x0 - right.word.x0;
    });

  const lines: IIndexedLine[] = [];
  indexed.forEach((item) => {
    const last = lines.length > 0 ? lines[lines.length - 1] : undefined;
    const height = Math.max(item.word.y1 - item.word.y0, 1);
    if (!last || item.word.y0 - last.y0 > height * 0.55) {
      lines.push({
        items: [item],
        y0: item.word.y0,
        y1: item.word.y1
      });
      return;
    }
    last.items.push(item);
    last.y0 = Math.min(last.y0, item.word.y0);
    last.y1 = Math.max(last.y1, item.word.y1);
  });
  return lines;
}

function locateLinePoint(
  lines: IIndexedLine[],
  x: number,
  y: number
): { line: number; x: number } | undefined {
  let best = -1;
  let bestDist = Number.POSITIVE_INFINITY;
  lines.forEach((line, index) => {
    if (y >= line.y0 && y <= line.y1) {
      best = index;
      bestDist = 0;
      return;
    }
    const dist = y < line.y0 ? line.y0 - y : y - line.y1;
    if (dist < bestDist) {
      bestDist = dist;
      best = index;
    }
  });
  if (best < 0) {
    return undefined;
  }
  const line = lines[best];
  const maxDist = Math.max(28, (line.y1 - line.y0) * 1.35);
  if (bestDist > maxDist) {
    return undefined;
  }
  return { line: best, x };
}

export function wordIndexAtPoint(words: IOcrWord[], x: number, y: number): number {
  let best = -1;
  let bestDist = Number.POSITIVE_INFINITY;
  for (let index = 0; index < words.length; index++) {
    const word = words[index];
    const padX = Math.max(2, (word.x1 - word.x0) * 0.08);
    const padY = Math.max(3, (word.y1 - word.y0) * 0.35);
    if (x >= word.x0 - padX && x <= word.x1 + padX && y >= word.y0 - padY && y <= word.y1 + padY) {
      const cx = (word.x0 + word.x1) / 2;
      const cy = (word.y0 + word.y1) / 2;
      const dist = Math.abs(x - cx) + Math.abs(y - cy);
      if (dist < bestDist) {
        bestDist = dist;
        best = index;
      }
    }
  }
  return best;
}

function joinGap(left: string, right: string): string {
  const last = (left || '').charAt((left || '').length - 1);
  const first = (right || '').charAt(0);
  const cjk = /[\u3400-\u9FFF]/;
  if (cjk.test(last) && cjk.test(first)) {
    return '';
  }
  return ' ';
}
