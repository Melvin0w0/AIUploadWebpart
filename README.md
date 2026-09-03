# AI Upload

SharePoint Online web part that lets a user upload a scanned PDF (no text layer), run OCR in the browser, and show the recognized text in a textarea.

## Features

- Upload a PDF from the local computer
- Render each page and run OCR with [Tesseract.js](https://tesseract.projectnaptha.com/)
- Languages: English, Traditional Chinese, Simplified Chinese, or mixed
- Display results in an editable textarea, with copy and clear actions

OCR runs in the browser. The PDF is not sent to a custom backend. The first run downloads Tesseract language data from a public CDN.

## Prerequisites

- Node.js 22.14 or later (not Node 23+)
- A SharePoint Online tenant for testing

## Getting started

```bash
npm install
npm start
```

`npm start` opens the local workbench. Add the **AI Upload** web part to a modern page or to `https://{your-tenant}.sharepoint.com/_layouts/15/workbench.aspx`.

Set `initialPage` in `config/serve.json` to your tenant workbench URL if needed.

## Package for SharePoint

```bash
npm run build
```

The `.sppkg` file is written to `sharepoint/solution/ai-upload.sppkg`. Upload it to the tenant or site app catalog, then add **AI Upload** to a page.

## How it works

1. The user selects a PDF and an OCR language, then clicks **Convert to OCR**.
2. [PDF.js](https://mozilla.github.io/pdf.js/) renders each page to a canvas.
3. Tesseract.js reads the canvas image and returns text.
4. The combined text is shown in the textarea.

## Notes

- First-time OCR is slower because language models are downloaded (especially Chinese).
- Large or high-resolution PDFs take more time and memory. Prefer a few pages when testing.
- OCR accuracy depends on scan quality. Tilted, low-contrast, or handwritten pages are less reliable.
