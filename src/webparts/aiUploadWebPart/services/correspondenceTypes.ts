import { CorrespondenceKind } from '../constants/incomingName';
import { IOcrPageResult } from './IPdfOcr';
import { ISignatureAnalysis } from './signatureSender';

export interface IDetectedFields {
  firstPage?: IOcrPageResult;
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
}

export interface IAiExtractionHints {
  page?: IOcrPageResult;
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
    letterType: ''
  };
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
