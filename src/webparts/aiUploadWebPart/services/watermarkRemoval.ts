import type { PDFDocumentProxy } from 'pdfjs-dist';
import type { OptionalContentConfig } from 'pdfjs-dist/types/src/display/optional_content_config';

const BLOCK_SIZE = 32;
const TEXT_CONTRAST = 48;
const PAPER_GRAY = 200;
const CHROMA_MIN = 28;
const LIGHT_LUMA = 118;

function luma(r: number, g: number, b: number): number {
  return 0.299 * r + 0.587 * g + 0.114 * b;
}

function percentileFromHist(hist: Uint32Array, count: number, pct: number): number {
  const target = Math.max(1, Math.floor(count * pct));
  let seen = 0;
  for (let i = 0; i < 256; i++) {
    seen += hist[i];
    if (seen >= target) {
      return i;
    }
  }
  return 255;
}

function blockBackgrounds(data: Uint8ClampedArray, width: number, height: number): {
  cols: number;
  rows: number;
  values: Uint8Array;
} {
  const cols = Math.max(1, Math.ceil(width / BLOCK_SIZE));
  const rows = Math.max(1, Math.ceil(height / BLOCK_SIZE));
  const values = new Uint8Array(cols * rows);
  const hist = new Uint32Array(256);

  for (let by = 0; by < rows; by++) {
    const y0 = by * BLOCK_SIZE;
    const y1 = Math.min(height, y0 + BLOCK_SIZE);
    for (let bx = 0; bx < cols; bx++) {
      const x0 = bx * BLOCK_SIZE;
      const x1 = Math.min(width, x0 + BLOCK_SIZE);
      hist.fill(0);
      let count = 0;
      for (let y = y0; y < y1; y++) {
        let i = (y * width + x0) * 4;
        for (let x = x0; x < x1; x++) {
          hist[data[i] | 0]++;
          count++;
          i += 4;
        }
      }
      values[by * cols + bx] = percentileFromHist(hist, count, 0.9);
    }
  }

  return { cols, rows, values };
}

function sampleBackground(
  map: { cols: number; rows: number; values: Uint8Array },
  x: number,
  y: number,
  width: number,
  height: number
): number {
  const fx = Math.max(0, Math.min(map.cols - 1, (x + 0.5) * map.cols / width - 0.5));
  const fy = Math.max(0, Math.min(map.rows - 1, (y + 0.5) * map.rows / height - 0.5));
  const x0 = Math.floor(fx);
  const y0 = Math.floor(fy);
  const x1 = Math.min(map.cols - 1, x0 + 1);
  const y1 = Math.min(map.rows - 1, y0 + 1);
  const tx = fx - x0;
  const ty = fy - y0;
  const a = map.values[y0 * map.cols + x0];
  const b = map.values[y0 * map.cols + x1];
  const c = map.values[y1 * map.cols + x0];
  const d = map.values[y1 * map.cols + x1];
  return (a * (1 - tx) + b * tx) * (1 - ty) + (c * (1 - tx) + d * tx) * ty;
}

function suppressLightWatermark(image: ImageData): void {
  const { data, width, height } = image;
  const chroma = new Uint8Array(width * height);
  for (let i = 0, p = 0; i < data.length; i += 4, p++) {
    chroma[p] = Math.max(data[i], data[i + 1], data[i + 2]) - Math.min(data[i], data[i + 1], data[i + 2]);
    const y = luma(data[i], data[i + 1], data[i + 2]);
    data[i] = y;
    data[i + 1] = y;
    data[i + 2] = y;
  }

  const bgMap = blockBackgrounds(data, width, height);

  for (let y = 0; y < height; y++) {
    let i = y * width * 4;
    for (let x = 0; x < width; x++) {
      const gray = data[i];
      const bg = sampleBackground(bgMap, x, y, width, height);
      const contrast = bg - gray;
      const lightWash = gray >= LIGHT_LUMA && chroma[y * width + x] >= CHROMA_MIN;
      if (gray >= PAPER_GRAY || contrast < TEXT_CONTRAST || lightWash) {
        data[i] = 255;
        data[i + 1] = 255;
        data[i + 2] = 255;
      }
      i += 4;
    }
  }
}

export function createOcrCanvasWithoutWatermark(source: HTMLCanvasElement): HTMLCanvasElement {
  const out = document.createElement('canvas');
  out.width = source.width;
  out.height = source.height;
  const context = out.getContext('2d', { willReadFrequently: true });
  if (!context) {
    return source;
  }

  context.drawImage(source, 0, 0);
  try {
    const image = context.getImageData(0, 0, out.width, out.height);
    suppressLightWatermark(image);
    context.putImageData(image, 0, 0);
  } catch {
    return source;
  }

  return out;
}

const WATERMARK_LAYER = /watermark|draft|confidential|sample|do\s*not\s*copy/i;

export async function getOcrOptionalContentConfig(
  pdf: PDFDocumentProxy
): Promise<OptionalContentConfig | undefined> {
  try {
    const config = await pdf.getOptionalContentConfig();
    const groups = config.getGroups() as Record<string, { name?: string }> | null;
    if (!groups) {
      return undefined;
    }

    let hidden = false;
    for (const id of Object.keys(groups)) {
      const name = String(groups[id]?.name || id);
      if (WATERMARK_LAYER.test(name)) {
        config.setVisibility(id, false);
        hidden = true;
      }
    }
    return hidden ? config : undefined;
  } catch {
    return undefined;
  }
}

