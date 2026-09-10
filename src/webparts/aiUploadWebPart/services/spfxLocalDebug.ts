declare const DEBUG: boolean;

const DEVTOOLS_POLL_MS = 400;
const CONSOLE_CHECK_MS = 900;
const SHORTCUT_HOLD_MS = 800;
const DOCKED_WIDTH_GAP = 160;
const DOCKED_HEIGHT_GAP = 200;
const TABLE_OPEN_MS = 8;

export function isSpfxServeDebug(): boolean {
  if (typeof DEBUG === 'boolean' && DEBUG) {
    return true;
  }
  if (typeof document === 'undefined' || typeof window === 'undefined') {
    return false;
  }
  const search = window.location.search || '';
  const debugQuery = /(?:\?|&)debug=true(?:&|$)/i.test(search) && /debugManifestsFile=/i.test(search);
  const scripts = document.getElementsByTagName('script');
  for (let index = 0; index < scripts.length; index++) {
    if (/localhost:4321/i.test(scripts[index].src || '')) {
      return true;
    }
  }
  return debugQuery;
}

function isEmbeddedFrame(): boolean {
  if (typeof window === 'undefined') {
    return false;
  }
  try {
    return window.self !== window.top;
  } catch {
    return true;
  }
}

function isDevToolsShortcut(event: KeyboardEvent): boolean {
  if (event.key === 'F12' || event.code === 'F12' || event.keyCode === 123) {
    return true;
  }
  const key = (event.key || '').toLowerCase();
  const code = event.code || '';
  const openKey = key === 'i' || key === 'j' || key === 'c' || code === 'KeyI' || code === 'KeyJ' || code === 'KeyC';
  return openKey && event.shiftKey && (event.ctrlKey || event.metaKey);
}

function isDockedDevToolsOpen(): boolean {
  if (typeof window === 'undefined' || isEmbeddedFrame()) {
    return false;
  }
  if (!window.outerWidth || !window.outerHeight) {
    return false;
  }
  const widthGap = window.outerWidth - window.innerWidth;
  const heightGap = window.outerHeight - window.innerHeight;
  return widthGap > DOCKED_WIDTH_GAP || heightGap > DOCKED_HEIGHT_GAP;
}

let largeTableData: Record<string, string>[] | undefined;
let lastConsoleCheckAt = 0;
let lastConsoleOpen = false;
let maxLogPrintTime = 0;

function getLargeTableData(): Record<string, string>[] {
  if (largeTableData) {
    return largeTableData;
  }
  const row: Record<string, string> = {};
  for (let index = 0; index < 500; index++) {
    row[`${index}`] = `${index}`;
  }
  largeTableData = [];
  for (let index = 0; index < 50; index++) {
    largeTableData.push(row);
  }
  return largeTableData;
}

function nowMs(): number {
  return typeof performance !== 'undefined' && typeof performance.now === 'function'
    ? performance.now()
    : Date.now();
}

/**
 * Undocked DevTools still format console.table into DOM, which is slow.
 * Closed DevTools skip that work, so the same call is near 0ms.
 */
function measureConsoleDevToolsOpen(): boolean {
  if (typeof console === 'undefined' || typeof console.table !== 'function') {
    return false;
  }

  const data = getLargeTableData();
  try {
    const tableStart = nowMs();
    console.table(data);
    const tablePrintTime = nowMs() - tableStart;

    if (typeof console.log === 'function') {
      const logStart = nowMs();
      console.log(data);
      maxLogPrintTime = Math.max(maxLogPrintTime, nowMs() - logStart);
    }
    if (typeof console.clear === 'function') {
      console.clear();
    }

    if (tablePrintTime < TABLE_OPEN_MS) {
      return false;
    }
    if (maxLogPrintTime > 0) {
      return tablePrintTime > maxLogPrintTime * 8;
    }
    return true;
  } catch {
    return false;
  }
}

function isConsoleDevToolsOpen(force?: boolean): boolean {
  const t = Date.now();
  if (!force && t - lastConsoleCheckAt < CONSOLE_CHECK_MS) {
    return lastConsoleOpen;
  }
  lastConsoleCheckAt = t;
  lastConsoleOpen = measureConsoleDevToolsOpen();
  return lastConsoleOpen;
}

export function isDevToolsOpen(): boolean {
  return isDockedDevToolsOpen() || isConsoleDevToolsOpen(true);
}

export function subscribeDevToolsOpen(onChange: (open: boolean) => void): () => void {
  if (typeof window === 'undefined') {
    return () => undefined;
  }

  let current = isDockedDevToolsOpen() || isConsoleDevToolsOpen(true);
  let shortcutOpen = current;
  let lastShortcutAt = 0;
  onChange(current);

  const emit = (open: boolean): void => {
    if (open === current) {
      return;
    }
    current = open;
    onChange(current);
  };

  const check = (forceConsole?: boolean): void => {
    if (Date.now() - lastShortcutAt < SHORTCUT_HOLD_MS) {
      return;
    }
    if (isDockedDevToolsOpen()) {
      shortcutOpen = false;
      lastConsoleOpen = true;
      emit(true);
      return;
    }
    if (isConsoleDevToolsOpen(forceConsole)) {
      shortcutOpen = false;
      emit(true);
      return;
    }
    emit(shortcutOpen);
  };

  const onShortcut = (event: KeyboardEvent): void => {
    if (!isDevToolsShortcut(event)) {
      return;
    }
    if (event.type === 'keyup' && Date.now() - lastShortcutAt < 80) {
      return;
    }
    lastShortcutAt = Date.now();
    shortcutOpen = !current;
    emit(shortcutOpen);
  };

  const onResize = (): void => {
    check(true);
  };

  const intervalId = window.setInterval(() => check(false), DEVTOOLS_POLL_MS);
  window.addEventListener('resize', onResize);
  window.addEventListener('keydown', onShortcut, true);
  window.addEventListener('keyup', onShortcut, true);

  return () => {
    window.clearInterval(intervalId);
    window.removeEventListener('resize', onResize);
    window.removeEventListener('keydown', onShortcut, true);
    window.removeEventListener('keyup', onShortcut, true);
  };
}
