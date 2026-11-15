import { DOCS_COMMUNITY_PROTOCOL, serveDocsCommunityReads, type CommunityPluginApi } from './docs-community-plugin';

/** This entry is served on the trusted Host origin, with its matched Docs plugins.js. */
export function installDocsCommunityPluginEntry(
  plugin: CommunityPluginApi & { init?: () => void },
  signal: AbortSignal,
): void {
  signal.throwIfAborted();
  // Electron's outer window has another origin; the trusted launcher is an intermediate ancestor.
  const ancestors = new Set<Window>();
  let ancestor = window.parent;
  for (let depth = 0; depth < 8; depth++) {
    ancestors.add(ancestor);
    if (ancestor === ancestor.parent) break;
    ancestor = ancestor.parent;
  }
  let connected = false;
  const connect = (event: MessageEvent) => {
    if (
      connected ||
      event.origin !== location.origin ||
      !ancestors.has(event.source as Window) ||
      event.data?.protocol !== DOCS_COMMUNITY_PROTOCOL ||
      event.data?.type !== 'connect' ||
      typeof event.data?.sessionId !== 'string' ||
      !/^[\w-]{1,128}$/.test(event.data.sessionId) ||
      event.ports.length !== 1
    )
      return;
    serveDocsCommunityReads(event.ports[0], plugin, event.data.sessionId, signal);
    connected = true;
    window.removeEventListener('message', connect);
  };
  const previous = plugin.init;
  plugin.init = () => {
    previous?.();
    if (signal.aborted) return;
    window.addEventListener('message', connect);
    for (const target of ancestors)
      target.postMessage({ protocol: DOCS_COMMUNITY_PROTOCOL, type: 'ready' }, location.origin);
  };
  signal.addEventListener(
    'abort',
    () => {
      window.removeEventListener('message', connect);
      plugin.init = previous;
    },
    { once: true },
  );
}
