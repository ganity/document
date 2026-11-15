import { createDocsCommunityReadClient, DOCS_COMMUNITY_PROTOCOL } from './docs-community-plugin';

function ownedDescendant(source: MessageEventSource | null, editorWindow: Window): source is Window {
  if (!source || !('parent' in source)) return false;
  try {
    let frame = source as Window;
    for (let depth = 0; depth < 8; depth++) {
      if (frame === editorWindow) return true;
      if (frame === frame.parent) return false;
      frame = frame.parent;
    }
  } catch {
    return false;
  }
  return false;
}

/** Call before the trusted plugin starts; the owned editor window must come from Host. */
export function connectDocsCommunityReads(editorWindow: Window, sessionId: string, signal: AbortSignal) {
  signal.throwIfAborted();
  if (!/^[\w-]{1,128}$/.test(sessionId) || editorWindow === window) throw new Error('Invalid owned Docs window');
  return new Promise<ReturnType<typeof createDocsCommunityReadClient>>((resolve, reject) => {
    const cleanup = () => {
      clearTimeout(timer);
      window.removeEventListener('message', receive);
      signal.removeEventListener('abort', abort);
    };
    const abort = () => {
      cleanup();
      reject(new Error('Docs plugin connection cancelled'));
    };
    const timer = setTimeout(() => {
      cleanup();
      reject(new Error('Docs plugin connection timed out'));
    }, 30000);
    const receive = (event: MessageEvent) => {
      if (
        event.origin !== location.origin ||
        event.data?.protocol !== DOCS_COMMUNITY_PROTOCOL ||
        event.data?.type !== 'ready' ||
        !ownedDescendant(event.source, editorWindow)
      )
        return;
      cleanup();
      const channel = new MessageChannel();
      let client: ReturnType<typeof createDocsCommunityReadClient> | undefined;
      try {
        client = createDocsCommunityReadClient(channel.port1, sessionId, signal);
        event.source.postMessage({ protocol: DOCS_COMMUNITY_PROTOCOL, type: 'connect', sessionId }, location.origin, [
          channel.port2,
        ]);
        resolve(client);
      } catch (error) {
        client?.close();
        channel.port1.close();
        channel.port2.close();
        reject(error);
      }
    };
    signal.addEventListener('abort', abort, { once: true });
    window.addEventListener('message', receive);
  });
}
