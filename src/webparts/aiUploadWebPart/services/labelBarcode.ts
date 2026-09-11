import { prepareZXingModule, writeBarcode } from 'zxing-wasm/writer';
import zxingWriterWasm from '../assets/zxing_writer.wasm.jpg';

let zxingReady: Promise<void> | undefined;

export async function drawBarcodeCanvas(
  format: 'Code128' | 'DataMatrix',
  content: string,
  width?: number,
  height?: number
): Promise<HTMLCanvasElement> {
  const canvasWidth = width ?? 220;
  const canvasHeight = height ?? (format === 'DataMatrix' ? canvasWidth : 100);
  const canvas = document.createElement('canvas');
  canvas.width = canvasWidth;
  canvas.height = canvasHeight;
  const ctx = canvas.getContext('2d');
  if (!ctx) {
    throw new Error('Unable to create DataMatrix canvas context');
  }
  ctx.fillStyle = 'white';
  ctx.fillRect(0, 0, canvasWidth, canvasHeight);

  try {
    await ensureZXing();
    const result = format === 'DataMatrix'
      ? await writeBarcode(content, {
        format: 'DataMatrix',
        scale: 4,
        forceSquareDataMatrix: true
      })
      : await writeBarcode(content, {
        format: 'Code128',
        scale: 0
      });
    if (!result || !result.svg || canvasWidth <= 0 || canvasHeight <= 0) {
      throw new Error(result && result.error ? result.error : 'DataMatrix SVG generation failed.');
    }
    const svgBlob = new Blob([result.svg], { type: 'image/svg+xml' });
    const svgUrl = URL.createObjectURL(svgBlob);
    await new Promise<void>((resolve, reject) => {
      const img = new Image();
      img.onload = () => {
        ctx.drawImage(img, 0, 0, canvasWidth, canvasHeight);
        URL.revokeObjectURL(svgUrl);
        resolve();
      };
      img.onerror = () => {
        URL.revokeObjectURL(svgUrl);
        reject(new Error('Loading DataMatrix SVG failed'));
      };
      img.src = svgUrl;
    });
    return canvas;
  } catch {
    return errorCanvas(canvasWidth, canvasHeight);
  }
}

function errorCanvas(width: number, height: number): HTMLCanvasElement {
  const canvas = document.createElement('canvas');
  canvas.width = width;
  canvas.height = height;
  const ctx = canvas.getContext('2d');
  if (!ctx) {
    return canvas;
  }
  ctx.fillStyle = 'white';
  ctx.fillRect(0, 0, width, height);
  ctx.fillStyle = 'red';
  ctx.font = 'bold 40px Arial';
  ctx.textAlign = 'center';
  ctx.fillText('Error', width / 2, height / 2 + 10);
  return canvas;
}

async function ensureZXing(): Promise<void> {
  if (!zxingReady) {
    zxingReady = loadZXing();
  }
  await zxingReady;
}

async function loadZXing(): Promise<void> {
  const response = await fetch(zxingWriterWasm);
  if (!response.ok) {
    throw new Error('Could not load the barcode engine.');
  }
  const wasmBinary = await response.arrayBuffer();
  await prepareZXingModule({
    overrides: {
      wasmBinary
    },
    fireImmediately: true
  });
}
