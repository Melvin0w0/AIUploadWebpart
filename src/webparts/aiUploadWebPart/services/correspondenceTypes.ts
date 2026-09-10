import { CorrespondenceKind } from '../constants/incomingName';
import { IOcrPageResult } from './IPdfOcr';
import { ISignatureAnalysis } from './signatureSender';

export interface IPickedValue {
  value: string;
  source: string;
}

export interface IDetectedFields {
  firstPage?: IOcrPageResult;
  closingPage?: IOcrPageResult;
  signature: ISignatureAnalysis;
  receiverName: string;
  subjectText: string;
  refNo: string;
  yourRef: string;
  projectNumber: string;
  organization: string;
  issueDate: string;
  memoSender: string;
  letterType: string;
  sources: { [key: string]: string };
}

export interface IAiExtractionHints {
  page?: IOcrPageResult;
  closingPage?: IOcrPageResult;
  signature?: ISignatureAnalysis;
  receiverName?: string;
  subjectText?: string;
  refNo?: string;
  yourRef?: string;
  organization?: string;
  kind?: CorrespondenceKind;
  letterType?: string;
}

export function emptyDetectedFields(firstPage?: IOcrPageResult): IDetectedFields {
  return {
    firstPage,
    signature: {
      region: undefined,
      senderName: '',
      textBelow: ''
    },
    receiverName: '',
    subjectText: '',
    refNo: '',
    yourRef: '',
    projectNumber: '',
    organization: '',
    issueDate: '',
    memoSender: '',
    letterType: '',
    sources: {}
  };
}

export function picked(value: string, source: string): IPickedValue {
  const text = value || '';
  if (!text.trim()) {
    return { value: '', source: '' };
  }
  return { value: text, source };
}

export function tryPicked(fn: () => IPickedValue): IPickedValue {
  try {
    const result = fn();
    if (!result || !result.value) {
      return picked('', '');
    }
    return result;
  } catch {
    return picked('', '');
  }
}

export function tryText(fn: () => string): string {
  try {
    return fn() || '';
  } catch {
    return '';
  }
}

export async function tryTextAsync(fn: () => Promise<string>): Promise<string> {
  try {
    return (await fn()) || '';
  } catch {
    return '';
  }
}
