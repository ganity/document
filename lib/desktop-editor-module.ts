import { applyDocumentLanguage } from '@ranuts/shared/i18n';
import { mountEditorModule } from './embed-api';

type HostMessage = {
  type: string;
  id?: string;
  message?: { id?: string; type: string; payload: Record<string, unknown> };
  fileName?: string;
  bytes?: ArrayBuffer;
  pageLineage?: unknown;
};
interface EditorHost {
  onCommand(listener: (command: unknown) => void): () => void;
  sendEvent(message: HostMessage): void;
}

declare global {
  interface Window {
    deckwiseEditorHost?: EditorHost;
  }
}

// Only Electron's isolated editor preload supplies this capability. No filesystem or general IPC is exposed.
const host = window.deckwiseEditorHost;
if (!host) throw new Error('The editor module requires an isolated desktop host');
applyDocumentLanguage();
const editor = mountEditorModule({
  deckwise: new URLSearchParams(location.search).get('deckwise') === '1',
  subscribe: (listener) => host.onCommand(listener),
  send(type, payload, id) {
    if (type === 'document:save-requested') {
      host.sendEvent({ type: 'save-requested' });
    } else if (type === 'document:saved') {
      const file = payload.file;
      if (!(file instanceof Blob)) throw new Error('Editor export did not return a file');
      void file
        .arrayBuffer()
        .then((bytes) => {
          host.sendEvent({ type: 'saved', id, bytes, fileName: payload.fileName, pageLineage: payload.pageLineage });
        })
        .catch((error) => {
          host.sendEvent({
            type: 'message',
            message: { type: 'document:error', id, payload: { message: String(error) } },
          });
        });
    } else {
      host.sendEvent({ type: 'message', message: { type, payload, id } });
    }
  },
});
window.addEventListener('pagehide', () => editor.dispose(), { once: true });
