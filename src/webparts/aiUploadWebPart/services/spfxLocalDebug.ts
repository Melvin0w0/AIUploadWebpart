declare const DEBUG: boolean;

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
