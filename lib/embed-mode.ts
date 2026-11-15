/**
 * Embed detection: is this window running inside a host page's iframe (or
 * explicitly asked to behave as if it were)?
 *
 * `initEmbedApi` owns the embed *feature* -- the postMessage protocol, the
 * origin allowlist and the `embed-mode` body class. This module is just the
 * cheap, side-effect-free predicate the rest of the app reads to stay out of
 * the host's way: local save routing, the unsaved-changes guard and the
 * autosave history all have to be silent when the document on screen belongs
 * to someone else's page.
 */
let hostMode: 'embedded' | 'deckwise' | undefined;

export function setEditorHostMode(mode: typeof hostMode): void {
  hostMode = mode;
}

export function isEditorModule(): boolean {
  return hostMode !== undefined;
}

const EMBED_QUERY_KEYS = ['embed', 'embedded'];
export const DECKWISE_SAVE_REQUEST_EVENT = 'deckwise:save-requested';

export function isEmbedMode(): boolean {
  if (hostMode) return true;
  if (typeof window === 'undefined') {
    return false;
  }
  if (window.parent !== window) {
    return true;
  }
  const params = new URLSearchParams(window.location.search);
  return EMBED_QUERY_KEYS.some((key) => {
    if (!params.has(key)) return false;
    const value = params.get(key);
    return value === '' || value === '1' || value === 'true';
  });
}

export function isDeckWiseMode(): boolean {
  if (hostMode) return hostMode === 'deckwise';
  if (!isEmbedMode()) return false;
  const value = new URLSearchParams(window.location.search).get('deckwise');
  return value === '' || value === '1' || value === 'true';
}
