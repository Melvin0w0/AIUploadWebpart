import { PDFDocument } from 'pdf-lib';

export async function appendLabelPageToPdf(pdfBytes: Uint8Array, labelPng: ArrayBuffer): Promise<Uint8Array> {
  const pdfDoc = await PDFDocument.load(pdfBytes, { ignoreEncryption: true });
  const page = pdfDoc.addPage([595, 842]);
  const image = await pdfDoc.embedPng(labelPng);
  const pdfImgWidth = 650;
  const pdfImgHeight = pdfImgWidth * (image.height / image.width);
  page.drawImage(image, {
    x: 50,
    y: page.getHeight() - pdfImgHeight - 20,
    width: pdfImgWidth,
    height: pdfImgHeight
  });
  return pdfDoc.save();
}
