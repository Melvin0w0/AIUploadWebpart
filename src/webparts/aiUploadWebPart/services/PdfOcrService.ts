import * as pdfjsLib from 'pdfjs-dist';
import type { PDFDocumentProxy } from 'pdfjs-dist';
import Tesseract from 'tesseract.js';
import pdfWorkerAsset from '../assets/pdf.worker.min.jpg';
import { IOcrPageResult, IOcrProgress, IOcrResult } from './IPdfOcr';

const MAX_RENDER_WIDTH = 1600;
const TESSERACT_CDN = 'https://cdn.jsdelivr.net/npm';

let pdfWorkerReady: Promise<void> | undefined;

function pageHasClosing(text: string): boolean {
  const key = ' ' + (text || '').toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim() + ' ';
  return key.indexOf(' yours sincere') >= 0 ||
    key.indexOf(' yours faithful') >= 0 ||
    key.indexOf(' yours truly') >= 0;
}

function ensurePdfJsWorker(): Promise<void> {
  if (!pdfWorkerReady) {
    pdfWorkerReady = loadPdfJsWorker();
  }
  return pdfWorkerReady;
}

async function loadPdfJsWorker(): Promise<void> {
  const response = await fetch(pdfWorkerAsset);
  if (!response.ok) {
    throw new Error('Unable to load the bundled PDF worker.');
  }

  const source = await response.text();
  const blob = new Blob([source], { type: 'application/javascript' });
  pdfjsLib.GlobalWorkerOptions.workerSrc = URL.createObjectURL(blob);
}

export class PdfOcrService {
  public static async extractText(
    source: File | Uint8Array,
    language: string,
    onProgress: (progress: IOcrProgress) => void,
    onPage?: (page: IOcrPageResult) => void
  ): Promise<IOcrResult> {
    await ensurePdfJsWorker();

    onProgress({
      page: 0,
      totalPages: 0,
      percent: 0,
      status: 'Loading OCR engine'
    });

    const worker = await Tesseract.createWorker(language, 1, {
      workerPath: `${TESSERACT_CDN}/tesseract.js@5.1.1/dist/worker.min.js`,
      corePath: `${TESSERACT_CDN}/tesseract.js-core@5.1.0/tesseract-core-simd.wasm.js`,
      langPath: 'https://tessdata.projectnaptha.com/4.0.0'
    });

    try {
      const data = source instanceof Uint8Array
        ? source.slice()
        : new Uint8Array(await source.arrayBuffer());
      const pdf = await pdfjsLib.getDocument({
        data
      }).promise;
      if (pdf.numPages < 1) {
        throw new Error('The PDF has no pages to recognize.');
      }

      const pages: IOcrPageResult[] = [];
      const totalPages = pdf.numPages;

      for (let pageNum = 1; pageNum <= totalPages; pageNum++) {
        onProgress({
          page: pageNum,
          totalPages,
          percent: Math.round(((pageNum - 1) / totalPages) * 100),
          status: `Rendering page ${pageNum} of ${totalPages}`
        });

        const canvas = await PdfOcrService._renderPage(pdf, pageNum);

        onProgress({
          page: pageNum,
          totalPages,
          percent: Math.round(((pageNum - 0.5) / totalPages) * 100),
          status: `OCR page ${pageNum} of ${totalPages}`
        });

        const result = await worker.recognize(canvas);
        const pageText = (result.data.text || '').trim();
        const imageUrl = await PdfOcrService._canvasToObjectUrl(canvas);
        const ocrWords = result.data.words || [];
        const pageResult: IOcrPageResult = {
          pageNumber: pageNum,
          text: pageText,
          imageUrl,
          width: canvas.width,
          height: canvas.height,
          words: ocrWords
            .filter((word) => !!word.text && word.text.trim().length > 0)
            .map((word) => ({
              text: word.text,
              x0: word.bbox.x0,
              y0: word.bbox.y0,
              x1: word.bbox.x1,
              y1: word.bbox.y1
            }))
        };
        pages.push(pageResult);
        if (onPage) {
          onPage(pageResult);
        }

        canvas.width = 0;
        canvas.height = 0;

        if (pageHasClosing(pageText)) {
          break;
        }
      }

      onProgress({
        page: pages.length,
        totalPages: pages.length,
        percent: 100,
        status: 'Completed'
      });

      return {
        text: pages.map((page) => `----- Page ${page.pageNumber} -----\n${page.text}`).join('\n\n'),
        pages
      };
    } finally {
      await worker.terminate();
    }
  }

  private static async _renderPage(
    pdf: PDFDocumentProxy,
    pageNum: number
  ): Promise<HTMLCanvasElement> {
    const page = await pdf.getPage(pageNum);
    const unscaled = page.getViewport({ scale: 1 });
    const scale = unscaled.width > MAX_RENDER_WIDTH
      ? MAX_RENDER_WIDTH / unscaled.width
      : 2;
    const viewport = page.getViewport({ scale });

    const canvas = document.createElement('canvas');
    const context = canvas.getContext('2d', { willReadFrequently: true });
    if (!context) {
      throw new Error('Unable to create canvas context for OCR.');
    }

    canvas.width = Math.floor(viewport.width);
    canvas.height = Math.floor(viewport.height);

    const renderTask = page.render({
      canvasContext: context,
      viewport
    });
    await renderTask.promise;

    return canvas;
  }

  private static _canvasToObjectUrl(canvas: HTMLCanvasElement): Promise<string> {
    return new Promise((resolve, reject) => {
      canvas.toBlob((blob) => {
        if (!blob) {
          reject(new Error('Unable to create a PDF page preview.'));
          return;
        }
        resolve(URL.createObjectURL(blob));
      }, 'image/jpeg', 0.85);
    });
  }
}
