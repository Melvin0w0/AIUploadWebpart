import { IOcrPageResult, IOcrWord } from './IPdfOcr';
import { joinOcrWords, stripOcrStyleTags } from './ocrSelection';
import { SUBJECT } from './ocrWordStyles';

export interface ISignatureRegion {
  x0: number;
  y0: number;
  x1: number;
  y1: number;
}

export interface ISignatureAnalysis {
  region: ISignatureRegion | undefined;
  regionPageNumber?: number;
  senderName: string;
  textBelow: string;
}

export async function analyzeSignature(page?: IOcrPageResult): Promise<ISignatureAnalysis> {
  const empty: ISignatureAnalysis = {
    region: undefined,
    senderName: '',
    textBelow: ''
  };
  if (!page || !page.imageUrl || page.width <= 0 || page.height <= 0) {
    return empty;
  }

  try {
    const image = await loadImage(page.imageUrl);
    const closing = findLastClosingHit(page.words || []);
    const regions = preferRightmostSignatures(
      findInkSignatures(image, page.words || []),
      image.width
    );
    const region = pickSignatureBesideClosing(regions, closing, page.width, page.height);
    if (!region) {
      return empty;
    }
    return {
      region,
      textBelow: joinOcrWords(wordsAroundSignature(page, region)),
      senderName: ''
    };
  } catch {
    return empty;
  }
}

export async function analyzeDocumentSignature(pages?: IOcrPageResult[]): Promise<ISignatureAnalysis> {
  const list = pages || [];
  const closingPage = list.length > 0 ? list[list.length - 1] : undefined;
  const empty: ISignatureAnalysis = {
    region: undefined,
    senderName: '',
    textBelow: ''
  };
  const signature = await analyzeSignature(closingPage).catch(() => empty);
  if (closingPage && signature.region) {
    return {
      ...signature,
      regionPageNumber: closingPage.pageNumber
    };
  }
  return signature;
}

export function asPersonName(value: string): string {
  const line = (value || '').replace(/\s+/g, ' ').trim();
  if (!isUsableSenderName(line)) {
    return '';
  }
  return personNameFromLine(line) || loosePersonName(line);
}

export function extractOutgoingSender(page?: IOcrPageResult, region?: ISignatureRegion): string {
  if (!page) {
    return '';
  }
  const fromFull = senderFromPage(page, undefined);
  if (fromFull) {
    return fromFull;
  }
  if (region) {
    return senderFromPage(page, region);
  }
  return '';
}

export function extractOutgoingSenderFromPages(pages?: IOcrPageResult[], region?: ISignatureRegion, regionPageNumber?: number): string {
  const list = pages || [];
  for (let index = list.length - 1; index >= 0; index--) {
    const page = list[index];
    const pageRegion = region && (!regionPageNumber || page.pageNumber === regionPageNumber) ? region : undefined;
    const name = extractOutgoingSender(page, pageRegion);
    if (name) {
      return name;
    }
  }
  return '';
}

export function extractReceiverAboveDearSir(page?: IOcrPageResult): string {
  if (!page) {
    return '';
  }
  try {
    const fromAttn = extractAttnValue(page);
    if (fromAttn) {
      const cleaned = cleanReceiverName(fromAttn);
      if (cleaned) {
        return cleaned;
      }
    }
    const fromFloor = receiverAboveFloorLine(page);
    if (fromFloor) {
      return fromFloor;
    }
    const fromFloorText = receiverAboveFloorLineFromText(page.text || '');
    if (fromFloorText) {
      return fromFloorText;
    }
    const fromByHand = receiverBelowByHand(page);
    if (fromByHand) {
      return fromByHand;
    }
    return '';
  } catch {
    return '';
  }
}

export function extractOrganizationAboveAddressee(page?: IOcrPageResult): string {
  if (!page) {
    return '';
  }
  try {
    const fromWords = departmentsBelowOurRefAboveDear(page);
    if (fromWords) {
      return fromWords;
    }
    const fromText = departmentsBelowOurRefAboveDearFromText(page.text || '');
    if (fromText) {
      return fromText;
    }
    const fromFloor = organizationAboveFloorLine(page);
    if (fromFloor) {
      return fromFloor;
    }
    return organizationAboveFloorLineFromText(page.text || '');
  } catch {
    return '';
  }
}

function isOurRefLine(line: string): boolean {
  const key = normalizeKey(line);
  return key.indexOf('our ref') === 0;
}

function departmentsBelowOurRefAboveDear(page: IOcrPageResult): string {
  const lines = groupWordsIntoLines(page.words || []);
  if (lines.length === 0) {
    return '';
  }
  const dear = findSalutationHit(page.words || []);
  let ourRef: { y1: number } | undefined;
  for (let index = 0; index < lines.length; index++) {
    const line = lines[index];
    if (!isOurRefLine(line.text)) {
      continue;
    }
    if (dear && line.y0 >= dear.y0) {
      continue;
    }
    ourRef = line;
    break;
  }
  if (!ourRef) {
    return '';
  }
  const hits: string[] = [];
  for (let index = 0; index < lines.length; index++) {
    const line = lines[index];
    if (line.y0 < ourRef.y1 - 2) {
      continue;
    }
    if (dear && line.y1 > dear.y0 + 2) {
      continue;
    }
    if (isOurRefLine(line.text) || isSalutationLine(line.text) || isDeliveryLine(line.text)) {
      continue;
    }
    const text = joinOcrWords(line.words).replace(/\s+/g, ' ').trim();
    if (!isDepartmentLine(text)) {
      continue;
    }
    hits.push(text);
  }
  return hits.join(' ').replace(/\s+/g, ' ').trim();
}

function departmentsBelowOurRefAboveDearFromText(text: string): string {
  const lines = (text || '').split(/\r?\n/).map((line) => line.replace(/\s+/g, ' ').trim()).filter((line) => line.length > 0);
  let ourIndex = -1;
  let dearIndex = lines.length;
  for (let index = 0; index < lines.length; index++) {
    if (ourIndex < 0 && isOurRefLine(lines[index])) {
      ourIndex = index;
    }
    if (isSalutationLine(lines[index])) {
      dearIndex = index;
      break;
    }
  }
  if (ourIndex < 0) {
    return '';
  }
  const hits: string[] = [];
  for (let index = ourIndex + 1; index < dearIndex; index++) {
    const line = lines[index];
    if (isOurRefLine(line) || isDeliveryLine(line)) {
      continue;
    }
    if (isDepartmentLine(line)) {
      hits.push(line);
    }
  }
  return hits.join(' ').replace(/\s+/g, ' ').trim();
}

function isFloorLine(line: string): boolean {
  return /\b\d{1,3}\s*[/\\\uFF0F]\s*f\b/i.test((line || '').replace(/\s+/g, ' ').trim());
}

function stripOrganizationSymbols(value: string): string {
  return (value || '')
    .replace(/<[^>]+>/g, ' ')
    .replace(/[^a-zA-Z0-9\u4e00-\u9fff]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function organizationAboveFloorLine(page: IOcrPageResult): string {
  const lines = groupWordsIntoLines(page.words || []);
  if (lines.length === 0) {
    return '';
  }
  const dear = findSalutationHit(page.words || []);
  for (let index = 0; index < lines.length; index++) {
    const line = lines[index];
    if (dear && line.y1 > dear.y0 + 2) {
      continue;
    }
    if (!isFloorLine(line.text)) {
      continue;
    }
    for (let previous = index - 1; previous >= 0; previous--) {
      const candidate = lines[previous];
      if (dear && candidate.y1 > dear.y0 + 2) {
        continue;
      }
      if (isOurRefLine(candidate.text) || isSalutationLine(candidate.text) || isDeliveryLine(candidate.text) || isFloorLine(candidate.text)) {
        continue;
      }
      const text = stripOrganizationSymbols(joinOcrWords(candidate.words));
      if (text) {
        return text;
      }
    }
  }
  return '';
}

function organizationAboveFloorLineFromText(text: string): string {
  const lines = (text || '').split(/\r?\n/).map((line) => line.replace(/\s+/g, ' ').trim()).filter((line) => line.length > 0);
  let dearIndex = lines.length;
  for (let index = 0; index < lines.length; index++) {
    if (isSalutationLine(lines[index])) {
      dearIndex = index;
      break;
    }
  }
  for (let index = 0; index < dearIndex; index++) {
    if (!isFloorLine(lines[index])) {
      continue;
    }
    for (let previous = index - 1; previous >= 0; previous--) {
      const candidate = lines[previous];
      if (isOurRefLine(candidate) || isDeliveryLine(candidate) || isFloorLine(candidate)) {
        continue;
      }
      const cleaned = stripOrganizationSymbols(candidate);
      if (cleaned) {
        return cleaned;
      }
    }
  }
  return '';
}

function isReceiverSkipLine(line: string): boolean {
  return isDepartmentLine(line) ||
    isDirectorLine(line) ||
    isDeliveryLine(line) ||
    isFloorLine(line) ||
    isOurRefLine(line) ||
    isSalutationLine(line) ||
    isAttnLine(line) ||
    isAddressBlockStop(line);
}

function startsWithHonorific(line: string): boolean {
  const key = normalizeKey(line);
  return /^(mr|mrs|ms|miss|mdm|dr|ir|prof|professor|engr?|messrs|sir|madam|mx)\b/.test(key) ||
    /^(?:先生|女士|小姐|太太)/.test((line || '').trim());
}

function pickReceiverLine(raw: string): string {
  const cleaned = cleanReceiverName(raw);
  if (!cleaned) {
    return '';
  }
  if (startsWithHonorific(raw) || looksLikePersonName(cleaned) || looksLikePersonName(raw)) {
    return cleaned;
  }
  return '';
}

function receiverAboveFloorLine(page: IOcrPageResult): string {
  const lines = groupWordsIntoLines(page.words || []);
  if (lines.length === 0) {
    return '';
  }
  const dear = findSalutationHit(page.words || []);
  for (let index = 0; index < lines.length; index++) {
    const line = lines[index];
    if (dear && line.y1 > dear.y0 + 2) {
      continue;
    }
    if (!isFloorLine(line.text)) {
      continue;
    }
    for (let previous = index - 1; previous >= 0; previous--) {
      const candidate = lines[previous];
      if (dear && candidate.y1 > dear.y0 + 2) {
        continue;
      }
      const raw = joinOcrWords(candidate.words).replace(/\s+/g, ' ').trim();
      if (!raw || isReceiverSkipLine(raw)) {
        continue;
      }
      const picked = pickReceiverLine(raw);
      if (picked) {
        return picked;
      }
    }
  }
  return '';
}

function receiverAboveFloorLineFromText(text: string): string {
  const lines = (text || '').split(/\r?\n/).map((line) => line.replace(/\s+/g, ' ').trim()).filter((line) => line.length > 0);
  let dearIndex = lines.length;
  for (let index = 0; index < lines.length; index++) {
    if (isSalutationLine(lines[index])) {
      dearIndex = index;
      break;
    }
  }
  for (let index = 0; index < dearIndex; index++) {
    if (!isFloorLine(lines[index])) {
      continue;
    }
    for (let previous = index - 1; previous >= 0; previous--) {
      const raw = lines[previous];
      if (!raw || isReceiverSkipLine(raw)) {
        continue;
      }
      const picked = pickReceiverLine(raw);
      if (picked) {
        return picked;
      }
    }
  }
  return '';
}

function receiverBelowByHand(page: IOcrPageResult): string {
  const lines = linesBelowDelivery(page, 'hand');
  for (let index = 0; index < lines.length; index++) {
    const name = cleanReceiverName(lines[index]);
    if (name && looksLikePersonName(name)) {
      return name;
    }
  }
  for (let index = 0; index < lines.length; index++) {
    const name = cleanReceiverName(lines[index]);
    if (name && !isDepartmentLine(name) && !isDirectorLine(name) && !isDeliveryLine(name)) {
      return name;
    }
  }
  return '';
}

function cleanReceiverName(value: string): string {
  let text = stripParenthetical(value);
  text = text
    .replace(/^(?:(?:mr|mrs|ms|miss|dr|ir|prof(?:essor)?|engr?|sir|madam|mdm|mx|messrs)\b\.?\s*)+/i, '')
    .replace(/^(?:先生|女士|小姐|太太)\s*/g, '')
    .replace(/\s*(?:先生|女士|小姐|太太)\s*$/g, '')
    .replace(/\s+/g, ' ')
    .trim();
  return text;
}

function stripParenthetical(value: string): string {
  let text = (value || '').trim();
  let previous = '';
  while (text !== previous) {
    previous = text;
    text = text
      .replace(/[(\uFF08][^)\uFF09]*[)\uFF09]/g, '')
      .replace(/[()\uFF08\uFF09]/g, '')
      .replace(/\s+/g, ' ')
      .trim();
  }
  return text;
}

function linesBelowDelivery(page: IOcrPageResult, kind: 'post' | 'hand'): string[] {
  const hit = findDeliveryHit(page, kind);
  if (!hit) {
    return linesBelowDeliveryFromText(page.text || '', kind);
  }
  const column = deliveryColumn(page, hit, kind);
  const dear = findSalutationHit(page.words || []);
  const stopY = dear ? dear.y0 - 2 : page.height;
  const columnWords = (page.words || []).filter((word) => {
    const midX = (word.x0 + word.x1) / 2;
    return midX >= column.left && midX <= column.right;
  });
  const lines = groupWordsIntoLines(columnWords);
  const collected: string[] = [];
  for (let index = 0; index < lines.length; index++) {
    const line = lines[index];
    if (line.y1 <= hit.y1 - 4) {
      continue;
    }
    if (line.y0 >= stopY) {
      break;
    }
    const text = (line.text || '').replace(/\s+/g, ' ').trim();
    if (!text || isDeliveryLine(text) || isAttnLine(text) || isSalutationLine(text)) {
      continue;
    }
    collected.push(text);
    if (collected.length >= 8) {
      break;
    }
  }
  return collected.length > 0 ? collected : linesBelowDeliveryFromText(page.text || '', kind);
}

function linesBelowDeliveryFromText(text: string, kind: 'post' | 'hand'): string[] {
  const lines = (text || '').split(/\r?\n/).map((line) => line.replace(/\s+/g, ' ').trim()).filter((line) => line.length > 0);
  const collected: string[] = [];
  let capturing = false;
  for (let index = 0; index < lines.length; index++) {
    const line = lines[index];
    if (isDeliveryKindLine(line, kind) || (kind === 'post' && isCombinedDeliveryLine(line)) || (kind === 'hand' && isCombinedDeliveryLine(line))) {
      capturing = true;
      continue;
    }
    if (!capturing) {
      continue;
    }
    if (isSalutationLine(line) || isAttnLine(line) || isDeliveryLine(line)) {
      break;
    }
    const key = normalizeKey(line);
    if (key.indexOf('our ref') === 0 || key.indexOf('your ref') === 0) {
      break;
    }
    collected.push(line);
    if (collected.length >= 8) {
      break;
    }
  }
  return collected;
}

function findDeliveryHit(
  page: IOcrPageResult,
  kind: 'post' | 'hand'
): { x0: number; y0: number; x1: number; y1: number } | undefined {
  const words = (page.words || []).slice().sort((left, right) => {
    if (Math.abs(left.y0 - right.y0) > 8) {
      return left.y0 - right.y0;
    }
    return left.x0 - right.x0;
  });
  for (let index = 0; index < words.length; index++) {
    const key = wordKey(words[index].text || '');
    const compact = kind === 'post' ? /^byp(o|e)?st$/ : /^byh(a|e)?nd$/;
    if (compact.test(key)) {
      return wordBox(words[index]);
    }
    if (key !== 'by' && key !== '8y') {
      continue;
    }
    const lineMid = (words[index].y0 + words[index].y1) / 2;
    const lineHeight = Math.max(10, words[index].y1 - words[index].y0);
    for (let cursor = index + 1; cursor < Math.min(words.length, index + 5); cursor++) {
      const next = words[cursor];
      if (Math.abs((next.y0 + next.y1) / 2 - lineMid) > lineHeight * 0.85) {
        break;
      }
      const nextKey = wordKey(next.text || '');
      if (kind === 'post' && isPostWord(nextKey)) {
        return mergeBoxes(wordBox(words[index]), wordBox(next));
      }
      if (kind === 'hand' && isHandWord(nextKey)) {
        return mergeBoxes(wordBox(words[index]), wordBox(next));
      }
    }
  }
  return undefined;
}

function deliveryColumn(
  page: IOcrPageResult,
  hit: { x0: number; y0: number; x1: number; y1: number },
  kind: 'post' | 'hand'
): { left: number; right: number } {
  const other = findDeliveryHit(page, kind === 'post' ? 'hand' : 'post');
  if (other && Math.abs(other.y0 - hit.y0) < Math.max(24, (hit.y1 - hit.y0) * 2)) {
    if (hit.x0 <= other.x0) {
      return { left: 0, right: (hit.x1 + other.x0) / 2 };
    }
    return { left: (other.x1 + hit.x0) / 2, right: page.width };
  }
  return {
    left: Math.max(0, hit.x0 - 28),
    right: Math.min(page.width, Math.max(hit.x1 + 90, hit.x0 + page.width * 0.45))
  };
}

function wordBox(word: IOcrWord): { x0: number; y0: number; x1: number; y1: number } {
  return { x0: word.x0, y0: word.y0, x1: word.x1, y1: word.y1 };
}

function mergeBoxes(
  left: { x0: number; y0: number; x1: number; y1: number },
  right: { x0: number; y0: number; x1: number; y1: number }
): { x0: number; y0: number; x1: number; y1: number } {
  return {
    x0: Math.min(left.x0, right.x0),
    y0: Math.min(left.y0, right.y0),
    x1: Math.max(left.x1, right.x1),
    y1: Math.max(left.y1, right.y1)
  };
}

function isPostWord(key: string): boolean {
  return key === 'post' || key === 'pest' || key === 'pst' || key === 'postage';
}

function isHandWord(key: string): boolean {
  return key === 'hand' || key === 'hnd' || key === 'hands';
}

function isDeliveryLine(line: string): boolean {
  return isDeliveryKindLine(line, 'post') || isDeliveryKindLine(line, 'hand') || isCombinedDeliveryLine(line);
}

function isCombinedDeliveryLine(line: string): boolean {
  const key = normalizeKey(line);
  return /\bby\s+hand\b/.test(key) && /\bby\s+post\b/.test(key);
}

function isDeliveryKindLine(line: string, kind: 'post' | 'hand'): boolean {
  const key = normalizeKey(line);
  if (kind === 'post') {
    return /^by\s+(post|pest|pst)\b/.test(key) || /\bby\s+post\b/.test(key);
  }
  return /^by\s+hand\b/.test(key) || /\bby\s+hand\b/.test(key);
}

function isDepartmentLine(line: string): boolean {
  const text = (line || '').replace(/\s+/g, ' ').trim();
  if (!text) {
    return false;
  }
  return /\bdepartments?\b/i.test(text) ||
    /\bdept\.?\b/i.test(text) ||
    /(?:署|處)\s*$/.test(text);
}

function isOrgishToken(token: string): boolean {
  const key = (token || '').toLowerCase().replace(/[^a-z]/g, '');
  if (!key) {
    return false;
  }
  if (key === 'engineer' || key === 'engineers' || key === 'officer' || key === 'officers') {
    return false;
  }
  return key === 'civil' ||
    key === 'electrical' ||
    key === 'mechanical' ||
    key === 'geotechnical' ||
    key === 'environmental' ||
    key === 'structural' ||
    key === 'highways' ||
    key === 'highway' ||
    key.indexOf('enginee') === 0 ||
    key.indexOf('engmee') === 0 ||
    key.indexOf('enqine') === 0 ||
    key.indexOf('depart') === 0 ||
    key === 'dept' ||
    key === 'office' ||
    key === 'offices' ||
    key.indexOf('bureau') === 0 ||
    key.indexOf('divis') === 0 ||
    key.indexOf('drainag') === 0 ||
    key.indexOf('authorit') === 0 ||
    key.indexOf('corporat') === 0 ||
    key.indexOf('govern') === 0 ||
    key === 'ministry' ||
    key === 'section' ||
    key === 'branch' ||
    key === 'committee' ||
    key === 'commission' ||
    key === 'buildings' ||
    key === 'works';
}

function hasOrgOrDeptToken(line: string): boolean {
  const text = (line || '').replace(/\s+/g, ' ').trim();
  if (!text) {
    return false;
  }
  if (/工程(署|處|部|拓展)?|路政|渠務|水務|建築署|環保/.test(text)) {
    return true;
  }
  const tokens = normalizeKey(text).split(' ').filter((token) => token.length > 0);
  for (let index = 0; index < tokens.length; index++) {
    if (isOrgishToken(tokens[index])) {
      return true;
    }
  }
  return false;
}

function isOrgUnitLine(line: string): boolean {
  const text = (line || '').replace(/\s+/g, ' ').trim();
  if (!text) {
    return false;
  }
  if (isDepartmentLine(text) || hasOrgOrDeptToken(text)) {
    return true;
  }
  const key = normalizeKey(text);
  if (/\b(division|office|bureau|branch|section|authority|commission|committee|ministry|government|corporation)\b/.test(key)) {
    return true;
  }
  if (/^(hong\s+kong|hksar)(\s|$)/.test(key) || (/\bhong\s+kong\b/.test(key) && key.split(' ').length <= 4)) {
    return true;
  }
  if (/(?:部|局|科|組|所)\s*$/.test(text)) {
    return true;
  }
  return false;
}

function isCompanyOrFirmLine(line: string): boolean {
  const text = (line || '').replace(/\s+/g, ' ').trim();
  if (!text) {
    return false;
  }
  const key = normalizeKey(text);
  if (/\b(limited|ltd|inc|incorporated|company|corp|corporation|group|holdings|plc|partners?|llp)\b/.test(key)) {
    return true;
  }
  if (/\basia\b/.test(key) && !hasHonorific(text)) {
    return true;
  }
  if (/公司|集團|企業|有限/.test(text)) {
    return true;
  }
  return false;
}

function isNonSenderLine(line: string): boolean {
  return isActingForLine(line) ||
    isOrgUnitLine(line) ||
    isCompanyOrFirmLine(line) ||
    isIgnorableBelowClosing(line);
}

function isUsableSenderName(value: string): boolean {
  const name = (value || '').replace(/\s+/g, ' ').trim();
  if (!name || isNonSenderLine(name) || isJobTitleLine(name) || isNoiseLine(name) || hasOrgOrDeptToken(name)) {
    return false;
  }
  return !!personNameFromLine(name) || !!loosePersonName(name);
}

function isDirectorLine(line: string): boolean {
  return /^(?:the\s+)?directors?\b/i.test((line || '').replace(/\s+/g, ' ').trim());
}

function organizationAboveAttn(page: IOcrPageResult): string {
  const fromWords = firstLineAboveAttnFromWords(page);
  if (fromWords) {
    return fromWords;
  }
  return firstLineAboveAttnFromText(page.text || '');
}

function isAttnLine(line: string): boolean {
  return /^(?:for\s+the\s+)?(?:att(?:n|ention)|atin|attm|atln)\b/i.test((line || '').replace(/\s+/g, ' ').trim());
}

function firstLineAboveAttnFromWords(page: IOcrPageResult): string {
  const lines = groupWordsIntoLines(page.words || []);
  if (lines.length === 0) {
    return '';
  }
  const dear = findSalutationHit(page.words || []);
  const attnIndex = findAttnLineIndex(lines, dear ? dear.y0 : -1);
  if (attnIndex < 0) {
    return '';
  }
  const above: string[] = [];
  let nextBottom = lines[attnIndex].y0;
  for (let index = attnIndex - 1; index >= 0 && above.length < 4; index--) {
    const line = lines[index];
    const text = completeAddresseeLine(page, line);
    if (!text || isAttnLine(text) || isSalutationLine(text)) {
      continue;
    }
    if (isAddressBlockStop(text)) {
      break;
    }
    const lineHeight = Math.max(12, line.y1 - line.y0);
    if (nextBottom - line.y1 > lineHeight * 2.8) {
      break;
    }
    above.unshift(text);
    nextBottom = line.y0;
  }
  return above.length > 0 ? above[0] : '';
}

function firstLineAboveAttnFromText(text: string): string {
  const lines = (text || '').split(/\r?\n/).map((line) => (line || '').replace(/\s+/g, ' ').trim());
  let attnIndex = -1;
  for (let index = 0; index < lines.length; index++) {
    if (isAttnLine(lines[index])) {
      attnIndex = index;
      break;
    }
  }
  if (attnIndex < 0) {
    return '';
  }
  const above: string[] = [];
  for (let index = attnIndex - 1; index >= 0 && above.length < 4; index--) {
    const line = lines[index];
    if (!line) {
      if (above.length > 0) {
        break;
      }
      continue;
    }
    if (isAttnLine(line) || isSalutationLine(line)) {
      continue;
    }
    if (isAddressBlockStop(line)) {
      break;
    }
    above.unshift(line);
  }
  return above.length > 0 ? above[0] : '';
}

function findAttnLineIndex(lines: { text: string; y0: number; y1: number }[], dearY: number): number {
  const hits: number[] = [];
  for (let index = 0; index < lines.length; index++) {
    if (isAttnLine(lines[index].text)) {
      hits.push(index);
    }
  }
  if (hits.length === 0) {
    return -1;
  }
  if (dearY >= 0) {
    const aboveDear = hits.filter((index) => lines[index].y1 <= dearY + 6);
    if (aboveDear.length > 0) {
      return aboveDear[aboveDear.length - 1];
    }
  }
  return hits[0];
}

function organizationFromLines(lines: string[]): string {
  for (let index = 0; index < lines.length; index++) {
    const line = (lines[index] || '').replace(/\s+/g, ' ').trim();
    if (!line || isSalutationLine(line) || isAddressBlockStop(line)) {
      continue;
    }
    if (/^(?:for\s+the\s+)?att(?:n|ention)\b/i.test(line)) {
      continue;
    }
    if (!/^[A-Za-z]/.test(line)) {
      continue;
    }
    return line;
  }
  return '';
}

export async function extractSubjectBelowDearSir(page?: IOcrPageResult): Promise<string> {
  if (!page) {
    return '';
  }
  const fromBoldAndUnderline = subjectFromBoldAndUnderlinedLines(page);
  if (fromBoldAndUnderline) {
    return fromBoldAndUnderline;
  }
  return subjectFromReBlock(page);
}

export function subjectAppearsInPage(page: IOcrPageResult, value: string): boolean {
  const needle = normalizeSubjectMatch(value);
  if (needle.length < 4) {
    return false;
  }
  const haystacks = [
    normalizeSubjectMatch(page.text || ''),
    normalizeSubjectMatch(joinOcrWords(page.words || []))
  ];
  for (let index = 0; index < haystacks.length; index++) {
    if (haystacks[index].indexOf(needle) >= 0) {
      return true;
    }
  }
  return false;
}

function normalizeSubjectMatch(text: string): string {
  return (text || '')
    .toLowerCase()
    .replace(/[\u2018\u2019\u201c\u201d']/g, '')
    .replace(/[^a-z0-9\u4e00-\u9fff]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

export function dearSirBandRegion(page?: IOcrPageResult): ISignatureRegion | undefined {
  if (!page || page.width <= 0 || page.height <= 0) {
    return undefined;
  }
  const hit = findSalutationHit(page.words || []);
  const closing = findClosingHit(page.words || []);
  if (!hit) {
    return {
      x0: 0,
      y0: Math.round(page.height * 0.12),
      x1: Math.round(page.width * 0.96),
      y1: closing ? Math.max(8, closing.y0) : Math.round(page.height * 0.68)
    };
  }
  const y1 = closing
    ? Math.max(hit.y1 + 8, closing.y0)
    : Math.min(page.height, Math.max(hit.y1 + Math.max(160, page.height * 0.42), page.height * 0.68));
  return {
    x0: 0,
    y0: hit.y1,
    x1: Math.round(page.width * 0.96),
    y1: Math.min(page.height, y1)
  };
}

function extractAttnValue(page: IOcrPageResult): string {
  const fromWords = attnFromWords(page);
  if (fromWords) {
    return fromWords;
  }
  return attnFromText(page.text || '');
}

function attnFromWords(page: IOcrPageResult): string {
  const words = (page.words || []).slice().sort((left, right) => {
    if (Math.abs(left.y0 - right.y0) > 8) {
      return left.y0 - right.y0;
    }
    return left.x0 - right.x0;
  });
  for (let index = 0; index < words.length; index++) {
    if (!isAttnLabelWord(words, index)) {
      continue;
    }
    const value = valueAfterAttn(words, index);
    if (value) {
      return value;
    }
  }
  return '';
}

function isAttnLabelWord(words: IOcrWord[], index: number): boolean {
  const key = wordKey(words[index].text || '');
  if (isAttnKey(key)) {
    return true;
  }
  if (key !== 'of') {
    return false;
  }
  const prev = words[index - 1];
  const prevKey = prev ? wordKey(prev.text || '') : '';
  return prevKey === 'attention';
}

function isAttnKey(key: string): boolean {
  return key === 'attn' ||
    key === 'atin' ||
    key === 'attm' ||
    key === 'atln' ||
    key === 'attention';
}

function valueAfterAttn(words: IOcrWord[], index: number): string {
  let start = index + 1;
  while (words[start]) {
    const raw = (words[start].text || '').trim();
    const key = wordKey(raw);
    if (/^[:.-]+$/.test(raw) || key === 'of' || key === 'to') {
      start++;
      continue;
    }
    break;
  }

  const label = words[index];
  const lineMid = (label.y0 + label.y1) / 2;
  const lineHeight = Math.max(label.y1 - label.y0, 1);
  const collected: IOcrWord[] = [];
  for (let cursor = start; cursor < words.length; cursor++) {
    const word = words[cursor];
    const wordMid = (word.y0 + word.y1) / 2;
    if (Math.abs(wordMid - lineMid) > lineHeight * 0.85) {
      break;
    }
    if (word.x0 < label.x1 - 6) {
      continue;
    }
    if (isAttnValueStop(word.text || '')) {
      break;
    }
    collected.push(word);
  }
  const sameLine = joinOcrWords(collected).replace(/^[:.\s-]+/, '').trim();
  if (sameLine) {
    return sameLine;
  }

  const below: IOcrWord[] = [];
  for (let cursor = start; cursor < words.length; cursor++) {
    const word = words[cursor];
    if (word.y0 < label.y1 - lineHeight * 0.2) {
      continue;
    }
    if (word.y0 > label.y1 + lineHeight * 1.8) {
      break;
    }
    if (isAttnValueStop(word.text || '')) {
      break;
    }
    below.push(word);
  }
  return joinOcrWords(below).replace(/^[:.\s-]+/, '').trim();
}

function isAttnValueStop(text: string): boolean {
  const key = wordKey(text);
  return key === 'date' ||
    key === 'tel' ||
    key === 'fax' ||
    key === 'email' ||
    key === 'dear' ||
    key === 'our' ||
    key === 'your' ||
    key === 'page' ||
    isAttnKey(key);
}

function attnFromText(text: string): string {
  const lines = (text || '').split(/\r?\n/).map((line) => line.trim()).filter((line) => line.length > 0);
  const pattern = /(?:for\s+the\s+)?att(?:n|ention|in|m)\s*(?:of)?\s*[:.-]?\s*(.+)$/i;
  for (let index = 0; index < lines.length; index++) {
    const match = lines[index].match(pattern);
    if (!match) {
      continue;
    }
    const value = (match[1] || '')
      .replace(/\b(date|tel|fax|email|dear|our\s+ref|your\s+ref)\b.*$/i, '')
      .replace(/^[:.\s-]+/, '')
      .trim();
    if (value) {
      return value;
    }
    const next = lines[index + 1];
    if (next && !isSalutationLine(next) && !isAddressBlockStop(next) && !/^att(?:n|ention)/i.test(next)) {
      return next.replace(/\s+/g, ' ').trim();
    }
  }
  return '';
}

function wordKey(text: string): string {
  return (text || '').toLowerCase().replace(/[^a-z0-9]+/g, '');
}

function addresseeBlockFromText(text: string): string {
  const raw = (text || '').split(/\r?\n/).map((line) => line.trim());
  let dearIndex = -1;
  for (let index = 0; index < raw.length; index++) {
    if (isSalutationLine(raw[index])) {
      dearIndex = index;
      break;
    }
  }
  if (dearIndex < 0) {
    const salutation = findSalutationIndex(text || '');
    if (salutation.index < 0) {
      return '';
    }
    const before = (text || '').substring(0, salutation.index).split(/\r?\n/).map((line) => line.trim());
    return firstLineOfConsecutiveBlock(before);
  }
  return firstLineOfConsecutiveBlock(raw.slice(0, dearIndex));
}

function addresseeBlockFromWords(page: IOcrPageResult): string {
  const block = consecutiveAddresseeLines(page);
  return block.length > 0 ? block[0].text : '';
}

function consecutiveAddresseeLines(page: IOcrPageResult): { text: string; y0: number; y1: number }[] {
  const hit = findSalutationHit(page.words || []);
  if (!hit) {
    return [];
  }
  const above = (page.words || []).filter((word) => {
    const midX = (word.x0 + word.x1) / 2;
    const midY = (word.y0 + word.y1) / 2;
    return midY < hit.y0 - 1 &&
      midY >= hit.y0 - Math.max(380, page.height * 0.58) &&
      midX <= page.width * 0.92;
  });
  const lines = groupWordsIntoLines(above).filter((line) => line.y1 < hit.y0);
  const block: { text: string; y0: number; y1: number }[] = [];
  let nextTop = hit.y0;
  for (let index = lines.length - 1; index >= 0; index--) {
    const line = lines[index];
    if (isSalutationLine(line.text)) {
      continue;
    }
    if (isAddressBlockStop(line.text)) {
      break;
    }
    const lineHeight = Math.max(12, line.y1 - line.y0);
    if (block.length > 0 && nextTop - line.y1 > lineHeight * 2.8) {
      break;
    }
    block.unshift(line);
    nextTop = line.y0;
    if (block.length >= 12) {
      break;
    }
  }
  return block;
}

function completeAddresseeLine(
  page: IOcrPageResult,
  line: { text: string; y0: number; y1: number }
): string {
  const midY = (line.y0 + line.y1) / 2;
  const lineHeight = Math.max(10, line.y1 - line.y0);
  const onLine = (page.words || []).filter((word) => {
    const wordMid = (word.y0 + word.y1) / 2;
    return Math.abs(wordMid - midY) <= lineHeight * 0.7;
  }).sort((left, right) => left.x0 - right.x0);

  if (onLine.length === 0) {
    return (line.text || '').replace(/\s+/g, ' ').trim();
  }

  const leftCluster: IOcrWord[] = [];
  const gapLimit = Math.max(40, page.width * 0.065);
  for (let index = 0; index < onLine.length; index++) {
    const word = onLine[index];
    if (isRightColumnLabel(onLine, index)) {
      break;
    }
    if (leftCluster.length > 0) {
      const previous = leftCluster[leftCluster.length - 1];
      if (word.x0 - previous.x1 > gapLimit) {
        break;
      }
    }
    leftCluster.push(word);
  }

  const completed = joinOcrWords(leftCluster).replace(/\s+/g, ' ').trim();
  const original = (line.text || '').replace(/\s+/g, ' ').trim();
  return completed.length >= original.length ? completed : original;
}

function isRightColumnLabel(words: IOcrWord[], index: number): boolean {
  const key = wordKey(words[index].text || '');
  if (key === 'date' || key === 'tel' || key === 'fax' || key === 'email' || key === 'page') {
    return true;
  }
  const next = words[index + 1];
  const nextKey = next ? wordKey(next.text || '') : '';
  if ((key === 'our' || key === 'your' || key === 'you' || key === 'yr') &&
    (/^r+e+fs?(no|number)?$/.test(nextKey) || nextKey === 'reference')) {
    return true;
  }
  if (/^(our|your)?r+e+fs?(no|number)?$/.test(key) || key === 'reference' || key === 'ourreference' || key === 'yourreference') {
    const prev = words[index - 1];
    const prevKey = prev ? wordKey(prev.text || '') : '';
    if (key.indexOf('our') === 0 || key.indexOf('your') === 0) {
      return true;
    }
    return prevKey === 'our' || prevKey === 'your' || prevKey === 'you' || prevKey === 'yr';
  }
  const prev = words[index - 1];
  const prevKey = prev ? wordKey(prev.text || '') : '';
  return (key === 'ref' || key === 'reference') &&
    (prevKey === 'our' || prevKey === 'your' || prevKey === 'you' || prevKey === 'yr');
}

function consecutiveBlockFromText(text: string): string[] {
  const raw = (text || '').split(/\r?\n/).map((line) => line.trim());
  let dearIndex = -1;
  for (let index = 0; index < raw.length; index++) {
    if (isSalutationLine(raw[index])) {
      dearIndex = index;
      break;
    }
  }
  if (dearIndex < 0) {
    const salutation = findSalutationIndex(text || '');
    if (salutation.index < 0) {
      return [];
    }
    const before = (text || '').substring(0, salutation.index).split(/\r?\n/).map((line) => line.trim());
    return consecutiveLinesFromList(before);
  }
  return consecutiveLinesFromList(raw.slice(0, dearIndex));
}

function firstLineOfConsecutiveBlock(lines: string[]): string {
  const block = consecutiveLinesFromList(lines);
  return block.length > 0 ? block[0] : '';
}

function consecutiveLinesFromList(lines: string[]): string[] {
  const block: string[] = [];
  let sawGap = false;
  for (let index = lines.length - 1; index >= 0; index--) {
    const line = lines[index];
    if (!line) {
      if (block.length > 0) {
        break;
      }
      if (sawGap) {
        break;
      }
      sawGap = true;
      continue;
    }
    if (isSalutationLine(line)) {
      continue;
    }
    if (isAddressBlockStop(line)) {
      break;
    }
    block.unshift(line.replace(/\s+/g, ' ').trim());
    if (block.length >= 12) {
      break;
    }
  }
  return block;
}

function groupWordsIntoLines(words: IOcrWord[]): { text: string; x0: number; x1: number; y0: number; y1: number; words: IOcrWord[] }[] {
  const sorted = words.slice().sort((left, right) => {
    if (Math.abs(left.y0 - right.y0) > 8) {
      return left.y0 - right.y0;
    }
    return left.x0 - right.x0;
  });
  const groups: IOcrWord[][] = [];
  sorted.forEach((word) => {
    const last = groups[groups.length - 1];
    if (!last) {
      groups.push([word]);
      return;
    }
    const lastMid = (last[0].y0 + last[0].y1) / 2;
    const wordMid = (word.y0 + word.y1) / 2;
    const lineHeight = Math.max(last[0].y1 - last[0].y0, 10);
    if (Math.abs(wordMid - lastMid) <= lineHeight * 0.6) {
      last.push(word);
      return;
    }
    groups.push([word]);
  });
  return mapWordGroupsToLines(groups);
}

function medianWordHeight(words: IOcrWord[]): number {
  const heights = (words || [])
    .map((word) => Math.max(1, word.y1 - word.y0))
    .sort((left, right) => left - right);
  if (heights.length === 0) {
    return 16;
  }
  return heights[Math.floor(heights.length / 2)];
}

function groupWordsIntoTightLines(words: IOcrWord[]): { text: string; x0: number; x1: number; y0: number; y1: number; words: IOcrWord[] }[] {
  const medianH = medianWordHeight(words);
  const threshold = Math.max(6, Math.min(12, medianH * 0.4));
  const sorted = words.slice().sort((left, right) => {
    if (Math.abs(left.y0 - right.y0) > threshold) {
      return left.y0 - right.y0;
    }
    return left.x0 - right.x0;
  });
  const groups: IOcrWord[][] = [];
  sorted.forEach((word) => {
    const last = groups[groups.length - 1];
    if (!last) {
      groups.push([word]);
      return;
    }
    const lastY0 = last.reduce((sum, item) => sum + item.y0, 0) / last.length;
    if (Math.abs(word.y0 - lastY0) <= threshold) {
      last.push(word);
      return;
    }
    groups.push([word]);
  });
  return mapWordGroupsToLines(groups);
}

function mapWordGroupsToLines(groups: IOcrWord[][]): { text: string; x0: number; x1: number; y0: number; y1: number; words: IOcrWord[] }[] {
  return groups.map((group) => {
    let x0 = group[0].x0;
    let x1 = group[0].x1;
    let y0 = group[0].y0;
    let y1 = group[0].y1;
    group.forEach((word) => {
      x0 = Math.min(x0, word.x0);
      x1 = Math.max(x1, word.x1);
      y0 = Math.min(y0, word.y0);
      y1 = Math.max(y1, word.y1);
    });
    return {
      text: joinOcrWords(group).replace(/\s+/g, ' ').trim(),
      x0,
      x1,
      y0,
      y1,
      words: group
    };
  }).filter((line) => line.text.length > 0);
}

function lineHasBoldAndUnderline(line: { words: IOcrWord[] }): boolean {
  for (let index = 0; index < line.words.length; index++) {
    const word = line.words[index];
    if (word.bold && word.underline) {
      return true;
    }
  }
  return false;
}

function collectBodyLines(
  page: IOcrPageResult,
  accept: (line: { words: IOcrWord[]; text: string }) => boolean,
  joinNearby?: boolean
): string {
  const pageWords = page.words || [];
  const lines = groupWordsIntoLines(pageWords);
  const dear = findSalutationHit(pageWords);
  const closing = findClosingHit(pageWords);
  const maxGap = joinNearby ? Math.max(0, SUBJECT.maxLineGap) : 0;
  const kept: { words: IOcrWord[]; y1: number }[] = [];
  for (let index = 0; index < lines.length; index++) {
    const line = lines[index];
    if (dear && line.y1 <= dear.y1) {
      continue;
    }
    if (closing && line.y0 >= closing.y0) {
      break;
    }
    if (isSalutationLine(line.text) || isClosingLine(line.text) || isDeliveryLine(line.text) || isAddressBlockStop(line.text)) {
      continue;
    }
    if (kept.length > 0) {
      const gap = line.y0 - kept[kept.length - 1].y1;
      if (gap > maxGap) {
        break;
      }
    }
    if (!accept(line)) {
      continue;
    }
    kept.push({ words: line.words, y1: line.y1 });
    if (!joinNearby) {
      break;
    }
  }
  if (kept.length === 0) {
    return '';
  }
  return kept
    .map((item) => stripSubjectLabel(joinOcrWords(item.words)).replace(/\s+/g, ' ').trim())
    .filter((text) => text.length > 0)
    .join(' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function subjectFromBoldAndUnderlinedLines(page: IOcrPageResult): string {
  return collectBodyLines(page, (line) => lineHasBoldAndUnderline(line) && looksLikeBoldSubject(line.text), true);
}

function looksLikeUnderlinedSubject(line: string): boolean {
  const trimmed = (line || '').trim();
  const letters = trimmed.replace(/[\s_\-.=]/g, '');
  if (letters.length < 2 || trimmed.length > 220) {
    return false;
  }
  if (/^[_.=-]{3,}$/.test(trimmed)) {
    return false;
  }
  if (isSalutationLine(trimmed) || isClosingLine(trimmed) || isAddressBlockStop(trimmed)) {
    return false;
  }
  return true;
}

function looksLikeBoldSubject(line: string): boolean {
  if (!looksLikeUnderlinedSubject(line)) {
    return false;
  }
  if (isBodyStart(line)) {
    return false;
  }
  const trimmed = (line || '').trim();
  if (trimmed.length > 100 && /^(i|we|please|thank)\b/i.test(trimmed)) {
    return false;
  }
  return true;
}

function stripSubjectLabel(line: string): string {
  return (line || '')
    .replace(/^(re|subject|ref|主旨|事由|關於|关于)\s*[:.-\uFF1A]\s*/i, '')
    .replace(/\s+/g, ' ')
    .trim();
}

function subjectFromReBlock(page: IOcrPageResult): string {
  const fromWords = subjectFromReWords(page);
  if (fromWords) {
    return fromWords;
  }
  return subjectFromReText(page.text || '');
}

function subjectFromReWords(page: IOcrPageResult): string {
  const lines = groupWordsIntoLines(page.words || []);
  if (lines.length === 0) {
    return '';
  }
  const dear = findSalutationHit(page.words || []);
  const closing = findClosingHit(page.words || []);
  const hits: number[] = [];
  for (let index = 0; index < lines.length; index++) {
    if (matchReLabel(lines[index].text) === undefined) {
      continue;
    }
    const line = lines[index];
    if (dear && line.y1 <= dear.y1 - 2) {
      continue;
    }
    if (closing && line.y0 >= closing.y0 - 2) {
      continue;
    }
    hits.push(index);
  }
  if (hits.length === 0) {
    return '';
  }
  return collectReContinuation(lines, hits[0], dear ? dear.y0 : -1);
}

function subjectFromReText(text: string): string {
  const lines = (text || '').split(/\r?\n/).map((line) => line.trim());
  const compact: { text: string; sourceIndex: number }[] = [];
  for (let index = 0; index < lines.length; index++) {
    if (lines[index]) {
      compact.push({ text: lines[index], sourceIndex: index });
    }
  }
  let dearIndex = -1;
  let closingIndex = compact.length;
  for (let index = 0; index < compact.length; index++) {
    if (dearIndex < 0 && isSalutationLine(compact[index].text)) {
      dearIndex = index;
    }
    if (isClosingLine(compact[index].text)) {
      closingIndex = index;
      break;
    }
  }
  const hits: number[] = [];
  const start = dearIndex >= 0 ? dearIndex : 0;
  for (let index = start; index < closingIndex; index++) {
    if (matchReLabel(compact[index].text) !== undefined) {
      hits.push(index);
    }
  }
  if (hits.length === 0) {
    return '';
  }
  const mapped = compact.map((line) => ({
    text: line.text,
    y0: line.sourceIndex,
    y1: line.sourceIndex
  }));
  return collectReContinuation(mapped, hits[0], dearIndex);
}

function matchReLabel(line: string): { value: string } | undefined {
  const trimmed = (line || '').replace(/\s+/g, ' ').trim();
  if (!trimmed || /\b(?:our|your?|yr)\s+(?:re|ref)\b/i.test(trimmed)) {
    return undefined;
  }
  const match = trimmed.match(/^(?:re|subject|主旨|事由|關於|关于)\s*[:;\uFF1A]\s*(.*)$/i);
  if (match) {
    return { value: stripSubjectLabel(match[1] || '') };
  }
  if (/^(?:re|subject|主旨|事由|關於|关于)$/i.test(trimmed)) {
    return { value: '' };
  }
  return undefined;
}

function collectReContinuation(
  lines: { text: string; y0: number; y1: number }[],
  startIndex: number,
  dearY: number
): string {
  const start = lines[startIndex];
  if (!start) {
    return '';
  }
  const label = matchReLabel(start.text);
  const block: string[] = [];
  if (label && label.value) {
    block.push(cleanSubjectLine(label.value));
  }
  const reIsAboveDear = dearY >= 0 && start.y1 < dearY - 2;
  for (let index = startIndex + 1; index < lines.length && block.length < 12; index++) {
    const line = lines[index];
    const text = cleanSubjectLine(line.text);
    if (!text) {
      if (block.length > 0) {
        break;
      }
      continue;
    }
    if (reIsAboveDear && line.y0 >= dearY - 2) {
      break;
    }
    const previous = lines[index - 1];
    const lineHeight = Math.max(12, previous.y1 - previous.y0, 1);
    if (line.y0 - previous.y1 > lineHeight * 3.2) {
      break;
    }
    if (isSubjectContinueStop(text)) {
      break;
    }
    block.push(text);
  }
  return block.join(' ').replace(/\s+/g, ' ').trim();
}

function cleanSubjectLine(line: string): string {
  return stripSubjectLabel(line)
    .replace(/\b(date|tel|fax|email)\s*[:.].*$/i, '')
    .replace(/\s+/g, ' ')
    .trim();
}

function isSubjectContinueStop(line: string): boolean {
  return isSalutationLine(line) ||
    isClosingLine(line) ||
    isBodyStart(line) ||
    isAddressBlockStop(line) ||
    /^(?:attn|attention|cc)\b/i.test((line || '').trim());
}

function isBodyStart(line: string): boolean {
  const trimmed = (line || '').trim();
  return /^(i|we|please|thank|further|with reference)\b/i.test(trimmed);
}

function findSalutationIndex(text: string): { index: number } {
  const match = (text || '').match(/\bdear\s+(s[il1]rs?|madams?|mesdames)\b/i) ||
    (text || '').match(/\bdear\s+sir\s*[/\\]?\s*madam\b/i);
  return { index: match && match.index !== undefined ? match.index : -1 };
}

function isSalutationLine(line: string): boolean {
  if (/敬啟者|敬启者|鈞鑒|台鑒/.test(line || '')) {
    return true;
  }
  const key = normalizeKey(line);
  if (!key) {
    return false;
  }
  if (/\bdear\s+(s[il1]rs?|madams?|mesdames)\b/.test(key) || /\bdear\s+sir\s*madam/.test(key)) {
    return true;
  }
  if (!/^dear\b/.test(key)) {
    return false;
  }
  const tokens = key.split(' ');
  return tokens.length <= 12 && key.length <= 90;
}

function isAddressBlockStop(line: string): boolean {
  const key = normalizeKey(line);
  return key.indexOf('our ref') === 0 ||
    key.indexOf('your ref') === 0 ||
    key.indexOf('by fax') === 0 ||
    key.indexOf('by email') === 0 ||
    key.indexOf('by post') === 0 ||
    /^date\b/.test(key) ||
    /^tel\b/.test(key) ||
    /^fax\b/.test(key) ||
    /^email\b/.test(key);
}

function findSalutationHit(words: IOcrWord[]): { y0: number; y1: number } | undefined {
  const sorted = words.slice().sort((left, right) => {
    if (Math.abs(left.y0 - right.y0) > 10) {
      return left.y0 - right.y0;
    }
    return left.x0 - right.x0;
  });
  for (let index = 0; index < sorted.length; index++) {
    if (!isDearWord(sorted[index].text || '')) {
      continue;
    }
    const lineMid = (sorted[index].y0 + sorted[index].y1) / 2;
    const nearby = sorted.slice(index, index + 6).filter((word) =>
      Math.abs((word.y0 + word.y1) / 2 - lineMid) < 16
    );
    const phrase = nearby.map((word) => word.text || '').join(' ');
    if (!isSalutationLine(phrase) && !nearby.some((word) => isSirWord(word.text || ''))) {
      continue;
    }
    let y0 = sorted[index].y0;
    let y1 = sorted[index].y1;
    nearby.forEach((word) => {
      y0 = Math.min(y0, word.y0);
      y1 = Math.max(y1, word.y1);
    });
    return { y0, y1 };
  }
  const lines = groupWordsIntoLines(words);
  for (let index = 0; index < lines.length; index++) {
    if (isSalutationLine(lines[index].text)) {
      return { y0: lines[index].y0, y1: lines[index].y1 };
    }
  }
  return undefined;
}

function isDearWord(text: string): boolean {
  const key = normalizeKey(text);
  return key === 'dear' || key === 'deor' || key === 'dcar' || key === 'dear,';
}

function isSirWord(text: string): boolean {
  const key = normalizeKey(text);
  return key === 'sir' ||
    key === 'sirs' ||
    key === 'sit' ||
    key === 'slr' ||
    key === 'madam' ||
    key === 'madams' ||
    key === 'mesdames';
}

function findInkSignatures(image: HTMLImageElement, words: IOcrWord[]): ISignatureRegion[] {
  const canvas = document.createElement('canvas');
  canvas.width = image.width;
  canvas.height = image.height;
  const context = canvas.getContext('2d', { willReadFrequently: true });
  if (!context) {
    return [];
  }
  context.drawImage(image, 0, 0);

  const x0 = Math.floor(image.width * 0.08);
  const y0 = Math.floor(image.height * 0.40);
  const x1 = image.width;
  const y1 = Math.floor(image.height * 0.97);
  const width = x1 - x0;
  const height = y1 - y0;
  if (width < 16 || height < 16) {
    return [];
  }

  const pixels = context.getImageData(x0, y0, width, height).data;
  canvas.width = 0;
  canvas.height = 0;

  const maxPrintedH = Math.max(28, medianWordHeight(words) * 2.4);
  const ink: number[] = [];
  const printed: boolean[] = [];
  for (let row = 0; row < height; row++) {
    let dark = 0;
    for (let col = 0; col < width; col++) {
      const index = (row * width + col) * 4;
      const lum = 0.299 * pixels[index] + 0.587 * pixels[index + 1] + 0.114 * pixels[index + 2];
      if (lum < 135) {
        dark++;
      }
    }
    ink.push(dark / width);
    printed.push(rowHasPrintedText(y0 + row, words, x0, x1, maxPrintedH));
  }

  const smooth: number[] = ink.map((_, row) => {
    const from = Math.max(0, row - 1);
    const to = Math.min(ink.length - 1, row + 1);
    let sum = 0;
    for (let index = from; index <= to; index++) {
      sum += ink[index];
    }
    return sum / (to - from + 1);
  });

  let printedSum = 0;
  let printedCount = 0;
  for (let row = 0; row < height; row++) {
    if (printed[row] && ink[row] > 0.008 && ink[row] < 0.07) {
      printedSum += ink[row];
      printedCount++;
    }
  }
  const printedTypical = printedCount > 0 ? printedSum / printedCount : 0.022;
  const isInkRow = (row: number): boolean => {
    if (smooth[row] < 0.018) {
      return false;
    }
    if (!printed[row]) {
      return smooth[row] >= 0.028;
    }
    return smooth[row] >= Math.max(0.05, printedTypical + 0.035);
  };

  const minHeight = Math.max(10, Math.round(image.height * 0.016));
  const maxHeight = Math.max(minHeight + 1, Math.round(image.height * 0.18));
  const footerStart = Math.floor(image.height * 0.975);
  const regions: ISignatureRegion[] = [];

  for (let start = 0; start < smooth.length; start++) {
    if (!isInkRow(start)) {
      continue;
    }
    let end = start;
    while (end < smooth.length && end - start < maxHeight && isInkRow(end)) {
      end++;
    }
    const runHeight = end - start;
    if (runHeight >= minHeight) {
      const top = y0 + start;
      const bottom = y0 + end;
      if (bottom < footerStart) {
        const bounds = inkHorizontalBounds(pixels, width, height, start, end);
        const left = bounds ? bounds.left : 0;
        const right = bounds ? bounds.right : width - 1;
        if (right - left >= 8 && (right - left) < width * 0.78) {
          regions.push({
            x0: x0 + left,
            y0: top,
            x1: x0 + right + 1,
            y1: bottom
          });
        }
      }
    }
    start = Math.max(start, end - 1);
  }

  return mergeSignatureRegions(regions, image.width, image.height);
}

function preferRightmostSignatures(regions: ISignatureRegion[], pageWidth: number): ISignatureRegion[] {
  if (regions.length === 0) {
    return [];
  }
  const sorted = regions.slice().sort((left, right) => {
    const rightCenter = (right.x0 + right.x1) / 2;
    const leftCenter = (left.x0 + left.x1) / 2;
    if (Math.abs(rightCenter - leftCenter) > pageWidth * 0.08) {
      return rightCenter - leftCenter;
    }
    return left.y0 - right.y0;
  });
  return sorted;
}

function inkHorizontalBounds(
  pixels: Uint8ClampedArray,
  width: number,
  height: number,
  rowStart: number,
  rowEnd: number
): { left: number; right: number } | undefined {
  const runH = Math.max(1, Math.min(height, rowEnd) - rowStart);
  if (runH < 1 || width < 8) {
    return undefined;
  }
  const colScore: number[] = [];
  for (let col = 0; col < width; col++) {
    let dark = 0;
    for (let row = rowStart; row < rowEnd && row < height; row++) {
      const index = (row * width + col) * 4;
      const lum = 0.299 * pixels[index] + 0.587 * pixels[index + 1] + 0.114 * pixels[index + 2];
      if (lum < 135) {
        dark++;
      }
    }
    colScore.push(dark / runH);
  }
  const threshold = 0.045;
  let left = 0;
  while (left < width && colScore[left] < threshold) {
    left++;
  }
  let right = width - 1;
  while (right > left && colScore[right] < threshold) {
    right--;
  }
  if (right - left < 8) {
    return undefined;
  }
  const pad = Math.max(6, Math.round(width * 0.02));
  return {
    left: Math.max(0, left - pad),
    right: Math.min(width - 1, right + pad)
  };
}

function mergeSignatureRegions(regions: ISignatureRegion[], pageWidth: number, pageHeight: number): ISignatureRegion[] {
  if (regions.length === 0) {
    return [];
  }
  const sorted = regions.slice().sort((left, right) => left.y0 - right.y0);
  const maxGap = Math.max(18, pageHeight * 0.028);
  const maxXGap = Math.max(24, pageWidth * 0.08);
  const merged: ISignatureRegion[] = [{
    x0: sorted[0].x0,
    y0: sorted[0].y0,
    x1: sorted[0].x1,
    y1: sorted[0].y1
  }];
  for (let index = 1; index < sorted.length; index++) {
    const prev = merged[merged.length - 1];
    const current = sorted[index];
    const xGap = current.x0 > prev.x1
      ? current.x0 - prev.x1
      : (prev.x0 > current.x1 ? prev.x0 - current.x1 : 0);
    if (current.y0 - prev.y1 <= maxGap && xGap <= maxXGap) {
      prev.y1 = Math.max(prev.y1, current.y1);
      prev.x0 = Math.min(prev.x0, current.x0);
      prev.x1 = Math.max(prev.x1, current.x1);
    } else {
      merged.push({
        x0: current.x0,
        y0: current.y0,
        x1: current.x1,
        y1: current.y1
      });
    }
  }
  return merged;
}

function rowHasPrintedText(y: number, words: IOcrWord[], x0: number, x1: number, maxWordHeight?: number): boolean {
  const maxH = maxWordHeight != null ? maxWordHeight : Number.POSITIVE_INFINITY;
  for (let index = 0; index < words.length; index++) {
    const word = words[index];
    const text = (word.text || '').trim();
    if (text.length < 2 || !/[A-Za-z\u3400-\u9FFF]{2,}/.test(text)) {
      continue;
    }
    if (word.y1 - word.y0 > maxH) {
      continue;
    }
    const overlapsY = word.y0 <= y && word.y1 >= y;
    const overlapsX = word.x0 < x1 && word.x1 > x0;
    if (overlapsY && overlapsX) {
      return true;
    }
  }
  return false;
}

function wordsAroundSignature(page: IOcrPageResult, region: ISignatureRegion): IOcrWord[] {
  const y0 = region.y0 - Math.max(8, page.height * 0.012);
  const y1 = Math.min(page.height, region.y1 + Math.max(96, page.height * 0.12));
  const x0 = Math.max(0, region.x0 - page.width * 0.06);
  return (page.words || []).filter((word) => {
    const midX = (word.x0 + word.x1) / 2;
    return midX >= x0 && word.y1 > y0 && word.y0 < y1;
  });
}

function senderNameAroundRegion(page: IOcrPageResult, region: ISignatureRegion): string {
  const lines = groupWordsIntoTightLines(wordsAroundSignature(page, region));
  return pickNameAboveTitle(lines.map((line) => ({ text: line.text, y0: line.y0 })), region);
}

function pickSignatureBesideClosing(
  regions: ISignatureRegion[],
  closing: { y0: number; y1: number } | undefined,
  pageWidth: number,
  pageHeight: number
): ISignatureRegion | undefined {
  if (regions.length === 0) {
    return undefined;
  }
  const nearClosing = closing
    ? regions.filter((region) => region.y1 >= closing.y0 - pageHeight * 0.1)
    : regions;
  const pool = nearClosing.length > 0 ? nearClosing : regions;
  return preferRightmostSignatures(pool, pageWidth)[0];
}

function signatureColumnRange(page: IOcrPageResult, region?: ISignatureRegion): { minX: number; maxX: number } {
  if (!region) {
    return { minX: 0, maxX: page.width };
  }
  const pad = Math.max(48, page.width * 0.12);
  const center = (region.x0 + region.x1) / 2;
  if (center >= page.width * 0.5) {
    return {
      minX: Math.min(page.width * 0.42, Math.max(0, region.x0 - pad)),
      maxX: page.width
    };
  }
  return {
    minX: 0,
    maxX: Math.max(page.width * 0.58, Math.min(page.width, region.x1 + pad))
  };
}

function buildOutgoingSignatureBlock(
  page: IOcrPageResult,
  closing: { y0: number; y1: number },
  region: ISignatureRegion | undefined,
  ccTop: number
): { text: string; y0: number; y1: number }[] {
  const zoneY0 = closing.y0 - 8;
  const zoneY1 = Math.min(page.height * 0.97, ccTop - 2);
  const column = signatureColumnRange(page, region);
  const words = (page.words || []).filter((word) => {
    const midX = (word.x0 + word.x1) / 2;
    return word.y1 > zoneY0 && word.y0 < zoneY1 && midX >= column.minX && midX <= column.maxX;
  });
  return groupWordsIntoTightLines(words)
    .map((line) => ({
      text: stripOcrStyleTags(line.text).replace(/\s+/g, ' ').trim(),
      y0: line.y0,
      y1: line.y1
    }))
    .filter((line) => {
      if (!line.text || isClosingLine(line.text) || isOutgoingCcLine(line.text)) {
        return false;
      }
      return true;
    });
}

function isAecomAsiaCompanyLine(line: string): boolean {
  const key = normalizeKey(line);
  if (!key) {
    return false;
  }
  const hasAecom = /\baecom\b/.test(key) || /\baeco[mn]\b/.test(key);
  const hasAsia = /\basia\b/.test(key);
  const hasCompany = /\b(company|co|limited|ltd)\b/.test(key);
  return hasAecom && hasAsia && hasCompany;
}

function pickSenderBelowAecomAsia(
  block: { text: string; y0: number; y1: number }[],
  region?: ISignatureRegion
): string {
  let aecomIndex = -1;
  for (let index = 0; index < block.length; index++) {
    if (isAecomAsiaCompanyLine(block[index].text)) {
      aecomIndex = index;
    }
  }
  if (aecomIndex >= 0) {
    for (let index = aecomIndex + 1; index < block.length; index++) {
      const text = block[index].text;
      if (isJobTitleLine(text) || isOrgUnitLine(text) || isOutgoingCcLine(text) || isAecomAsiaCompanyLine(text)) {
        break;
      }
      if (outgoingNameLineScore(text) > 0) {
        return text;
      }
    }
  }
  return pickBestNameLineInSignatureBlock(block, region);
}

function pickBestNameLineInSignatureBlock(
  block: { text: string; y0: number; y1: number }[],
  region?: ISignatureRegion
): string {
  let titleIndex = -1;
  for (let index = 0; index < block.length; index++) {
    if (isJobTitleLine(block[index].text) || isOrgUnitLine(block[index].text)) {
      titleIndex = index;
      break;
    }
  }
  const candidates = (titleIndex >= 0 ? block.slice(0, titleIndex) : block)
    .map((line, index) => ({
      text: line.text,
      y0: line.y0,
      score: outgoingNameLineScore(line.text),
      index
    }))
    .filter((item) => item.score > 0);
  if (candidates.length === 0) {
    return '';
  }
  candidates.sort((left, right) => {
    if (right.score !== left.score) {
      return right.score - left.score;
    }
    if (region) {
      const leftDist = Math.abs(((left.y0) - region.y1));
      const rightDist = Math.abs(((right.y0) - region.y1));
      if (leftDist !== rightDist) {
        return leftDist - rightDist;
      }
    }
    return right.y0 - left.y0;
  });
  return candidates[0].text;
}

function outgoingNameLineScore(line: string): number {
  const text = stripOcrStyleTags(line || '').replace(/\s+/g, ' ').trim();
  if (!text ||
    isActingForLine(text) ||
    isIgnorableBelowClosing(text) ||
    isCompanyOrFirmLine(text) ||
    isOrgUnitLine(text) ||
    isJobTitleLine(text) ||
    hasOrgOrDeptToken(text) ||
    isOutgoingCcLine(text) ||
    isClosingLine(text)) {
    return -1;
  }
  const person = personNameFromLine(text) || loosePersonName(text);
  if (!person && !looksLikePersonName(text)) {
    return -1;
  }
  let score = 1;
  if (looksLikePersonName(text)) {
    score += 6;
  }
  if (person) {
    score += 3;
  }
  if (hasHonorific(text)) {
    score += 2;
  }
  const tokens = text.replace(/^(ir|engr|eng|dr|mr|mrs|ms|prof)\.?\s+/i, '').split(/\s+/).filter((token) => token.length > 0);
  if (tokens.length >= 2 && tokens.length <= 4) {
    score += 2;
  }
  if (tokens.length === 3) {
    score += 1;
  }
  return score;
}

function findOutgoingCcTop(page: IOcrPageResult, minY: number): number {
  const lines = groupWordsIntoTightLines(page.words || []);
  for (let index = 0; index < lines.length; index++) {
    if (lines[index].y0 < minY) {
      continue;
    }
    if (isOutgoingCcLine(lines[index].text)) {
      return lines[index].y0;
    }
  }
  return page.height;
}

function isOutgoingCcLine(line: string): boolean {
  const text = stripOcrStyleTags(line || '').replace(/\s+/g, ' ').trim();
  if (!text) {
    return false;
  }
  const lower = text.toLowerCase();
  const key = normalizeKey(text);
  return /^(c\.?\s*c\.?|cc|copy\s+to|copied\s+to)\b/.test(lower) ||
    /^(c c|cc)\b/.test(key) ||
    /^(副本|抄送|副本送|副本抄送)\b/.test(text);
}

function senderFromPage(page: IOcrPageResult, region?: ISignatureRegion): string {
  const closing = findLastClosingHit(page.words || []);
  const closingBand = closing || {
    y0: page.height * 0.62,
    y1: page.height * 0.62
  };
  const ccTop = findOutgoingCcTop(page, closingBand.y1);
  const block = buildOutgoingSenderLines(page, closingBand, region, ccTop);
  return completeLineAbove(block);
}

function buildOutgoingSenderLines(
  page: IOcrPageResult,
  closing: { y0: number; y1: number },
  region: ISignatureRegion | undefined,
  ccTop: number
): { text: string; y0: number }[] {
  const zoneY0 = closing.y1 - 4;
  const zoneY1 = Math.min(page.height * 0.97, ccTop - 2);
  const column = signatureColumnRange(page, region);
  const words = (page.words || []).filter((word) => {
    const midX = (word.x0 + word.x1) / 2;
    return word.y1 > zoneY0 && word.y0 < zoneY1 && midX >= column.minX && midX <= column.maxX;
  });
  return groupWordsIntoLines(words)
    .map((line) => ({
      text: stripOcrStyleTags(joinOcrWords(line.words || []) || line.text).replace(/\s+/g, ' ').trim(),
      y0: line.y0
    }))
    .filter((line) => !!line.text);
}

function completeLineAbove(lines: { text: string; y0: number }[]): string {
  const block: string[] = [];
  for (let index = 0; index < lines.length; index++) {
    const raw = (lines[index].text || '').replace(/\s+/g, ' ').trim();
    if (isOutgoingCcLine(raw)) {
      break;
    }
    if (raw) {
      block.push(raw);
    }
  }
  let anchor = -1;
  for (let index = 0; index < block.length; index++) {
    const raw = block[index];
    if (isClosingLine(raw) || isIgnorableBelowClosing(raw)) {
      continue;
    }
    if (isJobTitleLine(raw) || isOrgUnitLine(raw)) {
      anchor = index;
      break;
    }
  }
  if (anchor < 0) {
    return '';
  }
  for (let index = anchor - 1; index >= 0; index--) {
    const raw = block[index];
    if (!raw || isClosingLine(raw) || isIgnorableBelowClosing(raw) || isOutgoingCcLine(raw)) {
      continue;
    }
    return raw;
  }
  return '';
}

function pickNameAboveTitle(lines: { text: string; y0: number }[], region?: ISignatureRegion): string {
  const found: { name: string; y0: number }[] = [];
  for (let index = 0; index < lines.length; index++) {
    const raw = stripOcrStyleTags(lines[index].text || '').replace(/\s+/g, ' ').trim();
    if (isOutgoingCcLine(raw)) {
      break;
    }
    if (!raw || isClosingLine(raw) || isActingForLine(raw) || isIgnorableBelowClosing(raw)) {
      continue;
    }
    if (isOrgUnitLine(raw) || isCompanyOrFirmLine(raw)) {
      const mixed = bestPersonNameWindow(raw);
      if (isUsableSenderName(mixed)) {
        found.push({ name: mixed, y0: lines[index].y0 });
      }
      if (found.length > 0) {
        break;
      }
      continue;
    }
    const cleaned = stripTrailingTitle(raw);
    if (!cleaned || isActingForLine(cleaned) || isIgnorableBelowClosing(cleaned)) {
      continue;
    }
    if (isOrgUnitLine(cleaned) || isCompanyOrFirmLine(cleaned)) {
      const mixed = bestPersonNameWindow(cleaned);
      if (isUsableSenderName(mixed)) {
        found.push({ name: mixed, y0: lines[index].y0 });
      }
      if (found.length > 0) {
        break;
      }
      continue;
    }
    if (isJobTitleLine(raw) || isJobTitleLine(cleaned)) {
      const nameOnTitleLine = preferredNameFromLine(cleaned) || bestPersonNameWindow(cleaned);
      if (isUsableSenderName(nameOnTitleLine)) {
        found.push({ name: nameOnTitleLine, y0: lines[index].y0 });
      }
      if (found.length > 0) {
        break;
      }
      continue;
    }
    const names = candidateNamesFromLine(cleaned);
    for (let nameIndex = 0; nameIndex < names.length; nameIndex++) {
      if (isUsableSenderName(names[nameIndex])) {
        found.push({ name: names[nameIndex], y0: lines[index].y0 });
      }
    }
  }

  if (found.length === 0) {
    return '';
  }
  found.sort((left, right) => left.y0 - right.y0);
  if (region) {
    const near = found.filter((item) => item.y0 >= region.y0 - 36 && item.y0 <= region.y1 + 96);
    if (near.length > 0) {
      return near[near.length - 1].name;
    }
  }
  return found[found.length - 1].name;
}

function candidateNamesFromLine(line: string): string[] {
  const primary = preferredNameFromLine(line);
  if (primary) {
    return [primary];
  }
  const left = leftPersonName(line);
  if (left) {
    return [left];
  }
  const windowName = bestPersonNameWindow(line);
  return windowName ? [windowName] : [];
}

function bestPersonNameWindow(line: string): string {
  const tokens = (line || '').replace(/,/g, ' ').split(/\s+/).filter((token) => token.length > 0);
  let best = '';
  for (let start = 0; start < tokens.length; start++) {
    for (let len = 2; len <= 4 && start + len <= tokens.length; len++) {
      const slice = tokens.slice(start, start + len).join(' ');
      const name = personNameFromLine(slice) || loosePersonName(slice);
      if (!isUsableSenderName(name)) {
        continue;
      }
      if (name.split(/\s+/).length > best.split(/\s+/).filter((token) => token.length > 0).length) {
        best = name;
      }
    }
  }
  return best;
}

function preferredNameFromLine(line: string): string {
  return nameFromZoneLine(line) || loosePersonName(line);
}

function nameFromZoneLine(line: string): string {
  const inner = firstParenthesesContent(line);
  if (inner) {
    const fromInner = personNameFromLine(inner) || loosePersonName(inner);
    if (fromInner) {
      return fromInner;
    }
  }
  return personNameFromLine(line);
}

function loosePersonName(line: string): string {
  const text = (line || '')
    .replace(/^[(\uFF08]\s*/, '')
    .replace(/\s*[)\uFF09]$/, '')
    .replace(/\s+/g, ' ')
    .trim();
  if (!text || isJobTitleLine(text) || isNoiseLine(text) || isActingForLine(text) || isOrgUnitLine(text) || isCompanyOrFirmLine(text) || hasOrgOrDeptToken(text)) {
    return '';
  }
  if (/工程師|總監|經理|主任|專員|顧問|秘書|署長|處長/.test(text)) {
    return '';
  }
  const cjk = text.match(/[\u3400-\u9FFF]/g);
  if (cjk && cjk.length >= 2 && cjk.length <= 4 && text.replace(/[\u3400-\u9FFF\s.·]/g, '').length === 0) {
    return text;
  }
  const withoutTitle = text.replace(/^(ir|engr|eng|dr|mr|mrs|ms|prof)\.?\s+/i, '');
  const tokens = withoutTitle.replace(/,/g, ' ').split(/\s+/).filter((token) => token.length > 0);
  if (tokens.length < 2 || tokens.length > 5) {
    return '';
  }
  const ok = tokens.every((token) =>
    /^[A-Za-z]\.?$/.test(token) ||
    /^[A-Za-z](?:\.[A-Za-z])+\.?$/.test(token) ||
    /^[A-Z][a-z]+(?:-[A-Z][a-z]+)?$/.test(token) ||
    /^[A-Z]{2,12}$/.test(token)
  );
  return ok ? text : '';
}

function leftPersonName(line: string): string {
  const withoutTitle = (line || '').replace(/^(ir|engr|eng|dr|mr|mrs|ms|prof)\.?\s+/i, '');
  const tokens = withoutTitle.replace(/,/g, ' ').split(/\s+/).filter((token) => token.length > 0);
  if (tokens.length < 4) {
    return '';
  }
  const firstThree = preferredNameFromLine(tokens.slice(0, 3).join(' '));
  if (firstThree) {
    return firstThree;
  }
  return preferredNameFromLine(tokens.slice(0, 2).join(' '));
}

function isActingForLine(line: string): boolean {
  const key = normalizeKey(line);
  return /^for\b/.test(key) ||
    /\bon\s+behal[f]?f?\b/.test(key) ||
    /\bbehaif\b/.test(key) ||
    /\bbehalf\b/.test(key);
}

function isJobTitleLine(line: string): boolean {
  const key = normalizeKey(line);
  if (!key) {
    return false;
  }
  if (/工程師|總監|經理|主任|專員|顧問|秘書|署長|處長/.test(line)) {
    return true;
  }
  const titles = [
    'director', 'manager', 'engineer', 'associate', 'consultant', 'officer',
    'secretary', 'architect', 'planner', 'surveyor', 'partner', 'chief',
    'assistant', 'principal', 'coordinator', 'specialist', 'supervisor',
    'technician', 'inspector', 'executive', 'president', 'leader', 'head'
  ];
  const tokens = key.split(' ');
  for (let index = 0; index < titles.length; index++) {
    if (tokens.indexOf(titles[index]) >= 0) {
      return true;
    }
  }
  return false;
}

function stripTrailingTitle(line: string): string {
  let text = (line || '').replace(/\s+/g, ' ').trim();
  text = text.split(/,\s+(?=(?:ir|engr|eng|dr|mr|mrs|ms|prof)?\.?\s*(?:chief|director|manager|engineer|associate|consultant|officer|secretary|architect)\b)/i)[0] || text;
  text = text.replace(/\s+(?:(?:ir|engr|eng|dr|mr|mrs|ms|prof)\.?\s+)?(?:chief\s+)?(?:director|manager|engineer|associate|consultant|officer|secretary|architect|planner|surveyor|partner|coordinator|specialist|supervisor|technician|inspector|executive|president|leader)\b.*$/i, '');
  text = text.replace(/\s*(工程師|總監|經理|主任|專員|顧問|秘書|署長|處長)\s*$/, '');
  return text.replace(/\s+/g, ' ').trim();
}

function isIgnorableBelowClosing(line: string): boolean {
  const key = normalizeKey(line);
  return !key ||
    isIgnorableParen(key) ||
    key === 'signed' ||
    /^for and on behalf/.test(key) ||
    /^for and beha/.test(key) ||
    /^[-_.=]+$/.test(line);
}

function isClosingLine(line: string): boolean {
  if (/此致|順頌|顺颂|專此|专此|敬祝/.test(line || '')) {
    return true;
  }
  const key = normalizeKey(line);
  return /(^|\s)yours?\s+sincer/.test(' ' + key) ||
    /(^|\s)yours?\s+faith/.test(' ' + key) ||
    /(^|\s)yours?\s+falth/.test(' ' + key) ||
    /(^|\s)yours?\s+tru/.test(' ' + key);
}

function firstParenthesesContent(text: string): string {
  const matches = (text || '').match(/[(\uFF08]\s*([^)\uFF09]{1,80}?)\s*[)\uFF09]/g) || [];
  for (let index = 0; index < matches.length; index++) {
    const innerMatch = matches[index].match(/[(\uFF08]\s*([^)\uFF09]{1,80}?)\s*[)\uFF09]/);
    const inner = innerMatch ? innerMatch[1].replace(/\s+/g, ' ').trim() : '';
    if (inner && !isIgnorableParen(inner)) {
      return inner;
    }
  }
  return '';
}

function isIgnorableParen(value: string): boolean {
  const key = normalizeKey(value);
  return key === 'signed' ||
    key === 'signature' ||
    key === 'sgd' ||
    key === 'chop' ||
    key === 'seal';
}

function findLastClosingHit(words: IOcrWord[]): { x0: number; y0: number; y1: number } | undefined {
  const lines = groupWordsIntoLines(words);
  let hit: { x0: number; y0: number; y1: number } | undefined;
  for (let index = 0; index < lines.length; index++) {
    const combined = lines[index + 1] ? (lines[index].text + ' ' + lines[index + 1].text) : lines[index].text;
    if (isClosingLine(lines[index].text) || isClosingLine(combined)) {
      hit = {
        x0: lines[index].x0,
        y0: lines[index].y0,
        y1: lines[index].y1
      };
    }
  }
  return hit || findClosingHit(words);
}

function findClosingHit(words: IOcrWord[]): { x0: number; y0: number; y1: number } | undefined {
  const sorted = words.slice().sort((left, right) => {
    if (Math.abs(left.y0 - right.y0) > 10) {
      return left.y0 - right.y0;
    }
    return left.x0 - right.x0;
  });
  for (let index = 0; index < sorted.length; index++) {
    const window = sorted.slice(index, index + 4);
    const lineMid = (sorted[index].y0 + sorted[index].y1) / 2;
    const sameLine = window.filter((word) => Math.abs((word.y0 + word.y1) / 2 - lineMid) < 12);
    const phrase = sameLine.map((word) => word.text || '').join(' ');
    if (!isClosingLine(phrase) && !isClosingLine(sorted[index].text || '')) {
      continue;
    }
    let x0 = sameLine[0].x0;
    let y0 = sameLine[0].y0;
    let y1 = sameLine[0].y1;
    sameLine.forEach((word) => {
      x0 = Math.min(x0, word.x0);
      y0 = Math.min(y0, word.y0);
      y1 = Math.max(y1, word.y1);
    });
    return { x0, y0, y1 };
  }
  const lines = groupWordsIntoLines(words);
  for (let index = 0; index < lines.length; index++) {
    if (isClosingLine(lines[index].text)) {
      return { x0: lines[index].x0, y0: lines[index].y0, y1: lines[index].y1 };
    }
  }
  return undefined;
}

function firstPersonName(text: string): string {
  const lines = (text || '')
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter((line) => line.length > 0);
  for (let index = 0; index < lines.length; index++) {
    if (isNonSenderLine(lines[index]) || isJobTitleLine(lines[index])) {
      const mixed = bestPersonNameWindow(lines[index]);
      if (isUsableSenderName(mixed)) {
        return mixed;
      }
      continue;
    }
    const name = personNameFromLine(lines[index]) || loosePersonName(lines[index]) || bestPersonNameWindow(lines[index]);
    if (isUsableSenderName(name)) {
      return name;
    }
  }
  return bestPersonNameWindow(text || '');
}

function personNameFromLine(line: string): string {
  const cleaned = (line || '')
    .replace(/^[\s(]+signed[\s)]+$/i, '')
    .replace(/^[-_.=]+$/, '')
    .trim();
  if (!cleaned || isJobTitleLine(cleaned) || isNoiseLine(cleaned) || isOrgUnitLine(cleaned) || isCompanyOrFirmLine(cleaned) || !looksLikePersonName(cleaned)) {
    return '';
  }
  return cleaned.replace(/\s+/g, ' ').trim();
}

function isNoiseLine(line: string): boolean {
  const key = normalizeKey(line);
  if (!key || /@|\d{3,}/.test(line) || /^https?:/i.test(line)) {
    return true;
  }
  const noise = [
    'tel', 'fax', 'email', 'www', 'http', 'page', 'enclosure', 'attachment',
    'limited', 'ltd', 'company', 'department', 'division', 'office',
    'director', 'manager', 'engineer', 'associate', 'consultant', 'officer',
    'secretary', 'architect', 'planner', 'surveyor', 'partner', 'chief',
    'yours', 'faithfully', 'sincerely', 'behalf'
  ];
  const tokens = key.split(' ');
  for (let index = 0; index < noise.length; index++) {
    if (tokens.indexOf(noise[index]) >= 0 && !hasHonorific(line)) {
      return true;
    }
  }
  return false;
}

function hasHonorific(line: string): boolean {
  return /^(ir|engr|eng|dr|mr|mrs|ms|prof)\b/i.test(line.trim());
}

function looksLikePersonName(line: string): boolean {
  const trimmed = line.trim();
  if (/工程師|總監|經理|主任|專員|顧問|秘書|署長|處長/.test(trimmed) || isOrgUnitLine(trimmed) || isCompanyOrFirmLine(trimmed) || hasOrgOrDeptToken(trimmed)) {
    return false;
  }
  const cjk = trimmed.match(/[\u3400-\u9FFF]/g);
  if (cjk && cjk.length >= 2 && cjk.length <= 4) {
    return trimmed.replace(/[\u3400-\u9FFF\s.·]/g, '').length === 0;
  }

  const withoutTitle = trimmed.replace(/^(ir|engr|eng|dr|mr|mrs|ms|prof)\.?\s+/i, '');
  const tokens = withoutTitle.replace(/,/g, ' ').split(/\s+/).filter((token) => token.length > 0);
  if (tokens.length < 2 || tokens.length > 5) {
    return false;
  }

  return tokens.every((token) => {
    if (/^[A-Za-z]\.?$/.test(token) || /^[A-Za-z](?:\.[A-Za-z])+\.?$/.test(token)) {
      return true;
    }
    if (/^[A-Z][a-z]+(?:-[A-Z][a-z]+)?$/.test(token) || /^[A-Z]{2,12}$/.test(token)) {
      return true;
    }
    return false;
  });
}

function normalizeKey(text: string): string {
  return text.toLowerCase().replace(/[^a-z0-9]+/g, ' ').replace(/\s+/g, ' ').trim();
}

function loadImage(url: string): Promise<HTMLImageElement> {
  return new Promise((resolve, reject) => {
    const image = new Image();
    image.onload = () => resolve(image);
    image.onerror = () => reject(new Error('Unable to load page image.'));
    image.src = url;
  });
}
