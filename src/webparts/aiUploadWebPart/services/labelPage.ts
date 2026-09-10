import { SPHttpClient } from '@microsoft/sp-http';
import { LabelType, labelImageFileName } from '../constants/labelType';
import { UploadType } from '../constants/uploadType';
import { ILabelStaff } from './notificationSetup';
import { drawBarcodeCanvas } from './labelBarcode';

export type { ILabelStaff };

export interface ILabelPageInput {
  labelType: LabelType;
  projectNumber: string;
  leadingBl: string;
  registrationNumber: string;
  organization: string;
  sender: string;
  receiver: string;
  subject: string;
  subProjectNumber: string;
  issueDateIso: string;
  refNo: string;
  hasAttachment: boolean;
  ccToAecom: boolean;
  uploadType: UploadType;
  staff: ILabelStaff;
  hasScan: boolean;
  siteUrls: string[];
}

interface ILabelPoint {
  x: number;
  y: number;
}

interface ILabelLayout {
  projectNumber: ILabelPoint;
  leadingBL: ILabelPoint;
  floor: ILabelPoint;
  PD: ILabelPoint;
  PM: ILabelPoint;
  DC: ILabelPoint;
  B1: ILabelPoint;
  attachmentYes?: ILabelPoint;
  attachmentNo?: ILabelPoint;
  scanYes?: ILabelPoint;
  scanNo?: ILabelPoint;
}

const LABEL_LAYOUTS: { [key: string]: ILabelLayout } = {
  confidential: {
    projectNumber: { x: 115, y: 175 },
    leadingBL: { x: 60, y: 203 },
    floor: { x: 80, y: 230 },
    PD: { x: 60, y: 315 },
    PM: { x: 65, y: 345 },
    DC: { x: 75, y: 370 },
    B1: { x: 115, y: 400 },
    attachmentYes: { x: 133, y: 260 },
    attachmentNo: { x: 173, y: 260 },
    scanYes: { x: 280, y: 260 },
    scanNo: { x: 320, y: 260 }
  },
  invoice: {
    projectNumber: { x: 110, y: 170 },
    leadingBL: { x: 60, y: 203 },
    floor: { x: 75, y: 230 },
    PD: { x: 58, y: 318 },
    PM: { x: 60, y: 347 },
    DC: { x: 70, y: 373 },
    B1: { x: 113, y: 400 },
    attachmentYes: { x: 130, y: 260 },
    attachmentNo: { x: 170, y: 260 },
    scanYes: { x: 275, y: 260 },
    scanNo: { x: 318, y: 260 }
  },
  site: {
    projectNumber: { x: 110, y: 123 },
    leadingBL: { x: 60, y: 150 },
    floor: { x: 80, y: 180 },
    PD: { x: 60, y: 235 },
    PM: { x: 65, y: 265 },
    DC: { x: 75, y: 290 },
    B1: { x: 115, y: 320 }
  },
  default: {
    projectNumber: { x: 115, y: 175 },
    leadingBL: { x: 65, y: 203 },
    floor: { x: 80, y: 230 },
    PD: { x: 60, y: 315 },
    PM: { x: 65, y: 345 },
    DC: { x: 75, y: 370 },
    B1: { x: 115, y: 400 },
    attachmentYes: { x: 130, y: 260 },
    attachmentNo: { x: 170, y: 260 },
    scanYes: { x: 275, y: 260 },
    scanNo: { x: 318, y: 260 }
  }
};

const SECOND_BARCODE: { [key: string]: { x: number; baseYOffset: number } } = {
  confidential: { x: 380, baseYOffset: 50 },
  invoice: { x: 380, baseYOffset: 100 },
  site: { x: 380, baseYOffset: 50 },
  default: { x: 380, baseYOffset: 30 }
};

export async function generateLabelPagePng(http: SPHttpClient, input: ILabelPageInput): Promise<Blob> {
  const img = await loadLabelTemplate(http, input.siteUrls, input.labelType);
  const labelCanvas = document.createElement('canvas');
  labelCanvas.width = img.width;
  labelCanvas.height = img.height;
  const labelCtx = labelCanvas.getContext('2d');
  if (!labelCtx) {
    throw new Error('Could not create the label canvas.');
  }
  labelCtx.drawImage(img, 0, 0);

  const pos = LABEL_LAYOUTS[input.labelType] || LABEL_LAYOUTS.default;
  labelCtx.font = 'bold 24px Arial';
  labelCtx.fillStyle = 'black';
  labelCtx.textAlign = 'left';
  labelCtx.fillText(input.projectNumber, pos.projectNumber.x, pos.projectNumber.y);
  labelCtx.fillText(input.leadingBl, pos.leadingBL.x, pos.leadingBL.y);
  labelCtx.fillText(input.staff.floor || '', pos.floor.x, pos.floor.y);
  labelCtx.fillText(input.staff.pdUser || '', pos.PD.x, pos.PD.y);
  labelCtx.fillText(input.staff.pmUser || '', pos.PM.x, pos.PM.y);
  labelCtx.fillText(input.staff.dcUser || '', pos.DC.x, pos.DC.y);
  labelCtx.fillText(input.staff.b1User || '', pos.B1.x, pos.B1.y);

  if (input.labelType !== 'site' && pos.attachmentYes && pos.attachmentNo && pos.scanYes && pos.scanNo) {
    const attachment = input.hasAttachment ? pos.attachmentYes : pos.attachmentNo;
    const scan = input.hasScan ? pos.scanYes : pos.scanNo;
    labelCtx.fillText('✓', attachment.x, attachment.y);
    labelCtx.fillText('✓', scan.x, scan.y);
  }

  const leftSpace = 360;
  const rightSpace = 600;
  const extraBottomHeight = 218;
  const finalCanvas = document.createElement('canvas');
  finalCanvas.width = labelCanvas.width + leftSpace + rightSpace;
  finalCanvas.height = labelCanvas.height + extraBottomHeight;
  const finalCtx = finalCanvas.getContext('2d');
  if (!finalCtx) {
    throw new Error('Could not create the label page canvas.');
  }
  finalCtx.fillStyle = 'white';
  finalCtx.fillRect(0, 0, finalCanvas.width, finalCanvas.height);

  const labelOffsetX = 350;
  finalCtx.drawImage(labelCanvas, labelOffsetX, 0);

  const now = new Date();
  const yy = now.getFullYear().toString().slice(-2);
  const mm = pad2(now.getMonth() + 1);
  const dd = pad2(now.getDate());
  const hh = pad2(now.getHours());
  const min = pad2(now.getMinutes());
  const ss = pad2(now.getSeconds());
  const timeCode = `${yy}${mm}${dd} ${hh}${min}${ss}9999`;
  const simpleTime = `${yy}${mm}${dd} ${hh}${min}${ss}`;

  const barcodeCanvas1 = await drawBarcodeCanvas('Code128', timeCode, 220, 100);
  finalCtx.drawImage(barcodeCanvas1, 8, 110, 320, 90);
  finalCtx.fillStyle = 'black';
  finalCtx.textAlign = 'left';
  finalCtx.font = 'bold 13px Arial';
  finalCtx.fillText('AECOM RECEIVED ON', 13, 85);
  finalCtx.font = '13px Arial';
  finalCtx.fillText(timeCode, 13, 102);

  const barcodeCanvas2 = await drawBarcodeCanvas('Code128', simpleTime, 220, 100);
  const second = SECOND_BARCODE[input.labelType] || SECOND_BARCODE.default;
  const secondY = labelCanvas.height - second.baseYOffset;
  finalCtx.drawImage(barcodeCanvas2, second.x, secondY, 320, 80);
  finalCtx.font = 'bold 13px Arial';
  finalCtx.fillText(simpleTime, second.x + 18, secondY - 5);

  const blCode = leadingBlCode(input.leadingBl);
  const registration = input.registrationNumber || '';
  const qrContent = [
    registration,
    registration.substring(0, Math.max(0, registration.length - 5)).replace(/^I/, ''),
    blCode,
    input.organization || '',
    input.sender || '',
    input.receiver || '',
    input.subject || '',
    input.projectNumber || '',
    input.subProjectNumber || 'None',
    input.issueDateIso || '',
    input.refNo || '',
    input.hasAttachment ? 'Y' : 'N',
    input.ccToAecom ? 'Y' : 'N',
    uploadTypeCode(input.uploadType),
    ''
  ].join('*');

  const qrX = labelOffsetX + labelCanvas.width - 220 + 280;
  const qrY = 180;
  const dmCanvas = await drawBarcodeCanvas('DataMatrix', qrContent, 220, 220);
  finalCtx.drawImage(dmCanvas, qrX, qrY);

  finalCtx.textAlign = 'left';
  finalCtx.fillStyle = 'black';
  finalCtx.font = 'bold 22px Arial';
  finalCtx.fillText(registration || 'N/A', qrX + 13, qrY - 62);
  finalCtx.font = 'bold 20px Arial';
  finalCtx.fillText('Remote Upload', qrX + 13, qrY - 12);
  if (blCode && input.leadingBl) {
    finalCtx.font = 'bold 18px Arial';
    finalCtx.fillText(`${blCode}  ${input.leadingBl}`, qrX + 13, qrY - 38);
  }

  return canvasToPng(finalCanvas);
}

async function loadLabelTemplate(http: SPHttpClient, siteUrls: string[], labelType: LabelType): Promise<HTMLImageElement> {
  const fileName = labelImageFileName(labelType);
  const blob = await fetchLabelImage(http, siteUrls, fileName);
  const imgUrl = URL.createObjectURL(blob);
  try {
    return await loadImage(imgUrl);
  } finally {
    URL.revokeObjectURL(imgUrl);
  }
}

async function fetchLabelImage(http: SPHttpClient, siteUrls: string[], fileName: string): Promise<Blob> {
  const urls: string[] = [];
  siteUrls.forEach((siteUrl) => {
    labelImageUrls(siteUrl, fileName).forEach((url) => {
      if (url && urls.indexOf(url) < 0) {
        urls.push(url);
      }
    });
  });
  let lastError = `Could not load label image ${fileName}.`;
  for (let i = 0; i < urls.length; i++) {
    try {
      const response = await http.get(urls[i], SPHttpClient.configurations.v1, {
        headers: {
          Accept: 'application/octet-stream'
        }
      });
      if (response.ok) {
        const blob = await response.blob();
        const blobType = (blob.type || '').toLowerCase();
        if (blob && blob.size > 200 && blobType.indexOf('json') < 0 && blobType.indexOf('xml') < 0 && blobType.indexOf('html') < 0) {
          return blob;
        }
      }
      lastError = `Could not load label image ${fileName} (${response.status}).`;
    } catch (err) {
      lastError = err instanceof Error ? err.message : lastError;
    }
  }
  throw new Error(lastError);
}

function labelImageUrls(siteUrl: string, fileName: string): string[] {
  const sitePath = urlPath(siteUrl);
  const encodedName = encodeURIComponent(fileName);
  const relative = `${sitePath}/SiteAssets/P1 Labal/${fileName}`;
  return [
    `${trimSlash(siteUrl)}/_api/web/GetFileByServerRelativeUrl('${escapeOData(relative)}')/$value`,
    `${trimSlash(siteUrl)}/_api/web/GetFileByServerRelativePath(decodedUrl='${escapeOData(relative)}')/$value`,
    `${trimSlash(siteUrl)}/SiteAssets/P1%20Labal/${encodedName}`
  ];
}

function loadImage(url: string): Promise<HTMLImageElement> {
  return new Promise((resolve, reject) => {
    const image = new Image();
    image.onload = () => resolve(image);
    image.onerror = () => reject(new Error('Label template image failed to load.'));
    image.src = url;
  });
}

function canvasToPng(canvas: HTMLCanvasElement): Promise<Blob> {
  return new Promise((resolve, reject) => {
    canvas.toBlob((blob) => {
      if (blob) {
        resolve(blob);
      } else {
        reject(new Error('Could not convert the label page to an image.'));
      }
    }, 'image/png', 1);
  });
}

function leadingBlCode(leadingBl: string): string {
  switch ((leadingBl || '').toLowerCase()) {
    case 'architecture':
      return 'A';
    case 'building engineering':
      return 'B';
    case 'construction services':
    case 'project and construction management':
      return 'S';
    case 'pdd':
    case 'design, planning and economics':
    case 'urbanism and planning':
      return 'D';
    case 'environment':
      return 'E';
    case 'geotechnical':
      return 'G';
    case 'isd':
    case 'digital':
      return 'I';
    case 'land supply and municipal':
      return 'L';
    case 'mep':
      return 'M';
    case 'pcc':
    case 'program, cost and consultancy':
      return 'C';
    case 'transportation':
      return 'T';
    case 'unclassified':
      return 'U';
    case 'water':
    case 'water and urban development':
      return 'W';
    default:
      return leadingBl || 'U';
  }
}

function uploadTypeCode(uploadType: UploadType): string {
  if (uploadType === 'confidentialInvoice') {
    return 'Invoice';
  }
  if (uploadType === 'confidentialMisc') {
    return 'MISC';
  }
  return 'N';
}

function pad2(value: number): string {
  return value < 10 ? `0${value}` : String(value);
}

function urlPath(siteUrl: string): string {
  try {
    return new URL(siteUrl).pathname.replace(/\/+$/, '') || '';
  } catch {
    return '';
  }
}

function trimSlash(value: string): string {
  return (value || '').replace(/\/+$/, '');
}

function escapeOData(value: string): string {
  return (value || '').replace(/'/g, "''");
}
