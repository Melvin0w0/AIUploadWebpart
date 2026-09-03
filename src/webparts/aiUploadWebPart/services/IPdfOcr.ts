export interface IOcrProgress {
  page: number;
  totalPages: number;
  percent: number;
  status: string;
}

export interface IOcrWord {
  text: string;
  x0: number;
  y0: number;
  x1: number;
  y1: number;
}

export interface IOcrPageResult {
  pageNumber: number;
  text: string;
  imageUrl: string;
  width: number;
  height: number;
  words: IOcrWord[];
}

export interface IOcrResult {
  text: string;
  pages: IOcrPageResult[];
}

export type OcrLanguage = 'eng' | 'chi_tra' | 'chi_sim' | 'eng+chi_tra' | 'eng+chi_sim';
