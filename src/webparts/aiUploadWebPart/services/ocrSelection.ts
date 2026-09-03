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

export function wordIndexAtPoint(words: IOcrWord[], x: number, y: number): number {
  for (let index = 0; index < words.length; index++) {
    const word = words[index];
    if (x >= word.x0 && x <= word.x1 && y >= word.y0 && y <= word.y1) {
      return index;
    }
  }
  return -1;
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
