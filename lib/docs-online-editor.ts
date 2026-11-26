import { DOCS_COMMUNITY_PLUGIN_GUID } from './docs-community-plugin';
import { connectDocsCommunityReads } from './docs-community-plugin-host';

/** An isolated Docs host. Signing, authorization and durable save receipts belong to the server. */
export interface DocsOnlineBinding {
  documentKey: string;
  userId: string;
  /** Host-installed asset URL, never derived from the signed config or document content. */
  communityPluginConfigUrl?: string;
}

export interface DocsOnlineConfig {
  documentType: 'slide';
  token: string;
  document: {
    key: string;
    fileType: 'pptx';
    title: string;
    url: string;
    permissions: { edit: boolean };
  };
  editorConfig: {
    mode: 'edit' | 'view';
    callbackUrl: string;
    user: { id: string; name: string };
    coEditing: { mode: 'fast'; change: false };
    customization: { autosave: true };
    plugins?: { autostart: [typeof DOCS_COMMUNITY_PLUGIN_GUID]; pluginsData: [string] };
  };
}

function requireValue(ok: unknown, field: string): asserts ok {
  if (!ok) throw new Error(`Invalid Docs online configuration: ${field}`);
}

function object(value: unknown, keys: string[], field: string): Record<string, unknown> {
  requireValue(value && typeof value === 'object' && !Array.isArray(value), field);
  requireValue(
    Object.keys(value).every((key) => keys.includes(key)),
    field,
  );
  return value as Record<string, unknown>;
}

function text(value: unknown, max: number): value is string {
  return typeof value === 'string' && value.trim().length > 0 && value.length <= max;
}

function remoteUrl(value: unknown, field: string): URL {
  requireValue(text(value, 8192), field);
  let url: URL;
  try {
    url = new URL(value);
  } catch {
    throw new Error(`Invalid Docs online configuration: ${field}`);
  }
  const loopback = ['127.0.0.1', '[::1]', 'localhost'].includes(url.hostname);
  requireValue(url.protocol === 'https:' || (url.protocol === 'http:' && loopback), field);
  requireValue(!url.username && !url.password && !url.hash, field);
  return url;
}

function pinnedPlugin(value: unknown, configUrl: string): void {
  const plugins = object(value, ['autostart', 'pluginsData'], 'plugins');
  requireValue(
    Array.isArray(plugins.autostart) &&
      plugins.autostart.length === 1 &&
      plugins.autostart[0] === DOCS_COMMUNITY_PLUGIN_GUID,
    'plugins.autostart',
  );
  requireValue(
    Array.isArray(plugins.pluginsData) && plugins.pluginsData.length === 1 && plugins.pluginsData[0] === configUrl,
    'plugins.pluginsData',
  );
}

/** Checks shape/binding only; Docs must verify the server-issued JWT. Never re-sign or rewrite it here. */
export function parseDocsOnlineConfig(json: string, binding: DocsOnlineBinding): DocsOnlineConfig {
  requireValue(typeof json === 'string' && json.length <= 65536, 'size');
  let parsed: unknown;
  try {
    parsed = JSON.parse(json);
  } catch {
    throw new Error('Invalid Docs online configuration: JSON');
  }
  const config = object(parsed, ['documentType', 'token', 'document', 'editorConfig'], 'config');
  requireValue(config.documentType === 'slide', 'documentType');
  requireValue(
    typeof config.token === 'string' && /^[A-Za-z0-9_-]+\.[A-Za-z0-9_-]+\.[A-Za-z0-9_-]+$/.test(config.token),
    'token',
  );
  const doc = object(config.document, ['key', 'fileType', 'title', 'url', 'permissions'], 'document');
  requireValue(
    typeof doc.key === 'string' && /^[A-Za-z0-9_.=-]{1,128}$/.test(doc.key) && doc.key === binding.documentKey,
    'document.key',
  );
  requireValue(doc.fileType === 'pptx' && text(doc.title, 128), 'document format');
  remoteUrl(doc.url, 'document.url');
  const permissions = object(doc.permissions, ['edit'], 'permissions');
  requireValue(typeof permissions.edit === 'boolean', 'permissions.edit');
  const editor = object(
    config.editorConfig,
    ['mode', 'callbackUrl', 'user', 'coEditing', 'customization', 'plugins'],
    'editorConfig',
  );
  requireValue(editor.mode === (permissions.edit ? 'edit' : 'view'), 'mode');
  remoteUrl(editor.callbackUrl, 'callbackUrl');
  const user = object(editor.user, ['id', 'name'], 'user');
  requireValue(
    text(user.id, 128) && user.id !== 'local-user' && user.id === binding.userId && text(user.name, 128),
    'user',
  );
  const coEditing = object(editor.coEditing, ['mode', 'change'], 'coEditing');
  requireValue(coEditing.mode === 'fast' && coEditing.change === false, 'coEditing');
  requireValue(object(editor.customization, ['autosave'], 'customization').autosave === true, 'autosave');
  if (binding.communityPluginConfigUrl === undefined) {
    requireValue(editor.plugins === undefined, 'plugins');
  } else {
    const pinned = remoteUrl(binding.communityPluginConfigUrl, 'plugin pin');
    requireValue(
      pinned.href === binding.communityPluginConfigUrl && pinned.pathname.endsWith('/config.json') && !pinned.search,
      'plugin pin',
    );
    pinnedPlugin(editor.plugins, binding.communityPluginConfigUrl);
    let claims: Record<string, unknown>;
    try {
      const payload = config.token.split('.')[1].replaceAll('-', '+').replaceAll('_', '/');
      claims = JSON.parse(
        new TextDecoder('utf-8', { fatal: true }).decode(Uint8Array.from(atob(payload), (c) => c.charCodeAt(0))),
      );
      requireValue(claims && typeof claims === 'object' && !Array.isArray(claims), 'plugin token');
    } catch {
      throw new Error('Invalid Docs online configuration: plugin token');
    }
    const signedEditor = object(
      claims.editorConfig,
      ['mode', 'callbackUrl', 'user', 'coEditing', 'customization', 'plugins'],
      'plugin token',
    );
    // Payload consistency rejects unsigned plugin additions; Docs still verifies the signature.
    pinnedPlugin(signedEditor.plugins, binding.communityPluginConfigUrl);
  }
  return parsed as DocsOnlineConfig;
}

export interface DocsOnlineEvents {
  onDocumentReady?: () => void;
  onDocumentStateChange?: (event: { data: boolean }) => void;
  onError?: (event: unknown) => void;
}

function awaitLoad(
  element: HTMLIFrameElement | HTMLScriptElement,
  append: () => void,
  signal: AbortSignal,
): Promise<void> {
  return new Promise((resolve, reject) => {
    const finish = (error?: Error) => {
      clearTimeout(timer);
      signal.removeEventListener('abort', abort);
      element.onload = element.onerror = null;
      if (error) reject(error);
      else resolve();
    };
    const abort = () => finish(new DOMException('Online editor cancelled', 'AbortError'));
    const timer = setTimeout(() => finish(new Error('Docs resource load timed out')), 30000);
    element.onload = () => finish();
    element.onerror = () => finish(new Error('Docs resource load failed'));
    signal.addEventListener('abort', abort, { once: true });
    try {
      append();
    } catch (error) {
      finish(error instanceof Error ? error : new Error('Docs resource load failed'));
    }
  });
}

/** apiUrl must come from Host's pinned Docs deployment, never from document content. */
export async function mountDocsOnlineEditor(
  container: HTMLElement,
  hostUrl: string,
  apiUrl: string,
  configJson: string,
  binding: DocsOnlineBinding,
  signal: AbortSignal,
  events: DocsOnlineEvents = {},
): Promise<{ destroy: () => void; communityReads?: Awaited<ReturnType<typeof connectDocsCommunityReads>> }> {
  const config = parseDocsOnlineConfig(configJson, binding);
  const requiresCommunityReads = binding.communityPluginConfigUrl !== undefined;
  const url = remoteUrl(apiUrl, 'apiUrl');
  const apiPath = '/web-apps/apps/api/documents/api.js';
  requireValue((url.pathname === apiPath || url.pathname === '/deckwise-office' + apiPath) && !url.search, 'apiUrl');
  const host = remoteUrl(hostUrl, 'hostUrl');
  requireValue(
    host.origin === window.location.origin && host.pathname.endsWith('/docs-online-host.html') && !host.search,
    'hostUrl',
  );
  if (binding.communityPluginConfigUrl !== undefined)
    requireValue(new URL(binding.communityPluginConfigUrl).origin === host.origin, 'plugin origin');
  signal.throwIfAborted();
  requireValue(container.isConnected, 'container');
  const frame = document.createElement('iframe');
  // about:blank produces parentOrigin=null in the native Docs API.
  frame.src = host.href;
  frame.title = 'Online presentation editor';
  frame.style.cssText = 'width:100%;height:100%;border:0';
  let editor: { destroyEditor: () => void } | undefined;
  const lifecycle = new AbortController();
  let communityReads: Awaited<ReturnType<typeof connectDocsCommunityReads>> | undefined;
  let active = true;
  const destroy = () => {
    active = false;
    lifecycle.abort();
    communityReads?.close();
    signal.removeEventListener('abort', destroy);
    const current = editor;
    editor = undefined;
    try {
      current?.destroyEditor();
    } finally {
      frame.remove();
    }
  };
  signal.addEventListener('abort', destroy, { once: true });
  try {
    await awaitLoad(frame, () => container.append(frame), signal);
    signal.throwIfAborted();
    const doc = frame.contentDocument!;
    requireValue(doc.querySelector('meta[name="docs-online-host"]')?.getAttribute('content') === '1', 'host document');
    const placeholder = doc.getElementById('docs-online');
    requireValue(placeholder, 'host placeholder');
    const script = doc.createElement('script');
    script.src = url.href;
    await awaitLoad(script, () => doc.head.append(script), signal);
    signal.throwIfAborted();
    requireValue(frame.isConnected, 'container');
    const api = (
      frame.contentWindow as unknown as {
        DocsAPI?: {
          DocEditor: new (
            id: string,
            config: DocsOnlineConfig & { events: DocsOnlineEvents },
          ) => { destroyEditor: () => void };
        };
      }
    ).DocsAPI;
    requireValue(api?.DocEditor, 'DocsAPI');
    const pendingReads = requiresCommunityReads
      ? connectDocsCommunityReads(frame.contentWindow!, crypto.randomUUID(), lifecycle.signal)
      : undefined;
    // Construction can fail before the connection promise is awaited.
    void pendingReads?.catch(() => {});
    editor = new api.DocEditor(placeholder.id, {
      ...config,
      events: {
        onDocumentReady: () => {
          if (active) events.onDocumentReady?.();
        },
        onDocumentStateChange: (event) => {
          if (active) events.onDocumentStateChange?.(event);
        },
        onError: (event) => {
          if (!active) return;
          // A native error invalidates this read channel; keep Docs visible for recovery.
          communityReads?.close();
          lifecycle.abort();
          events.onError?.(event);
        },
      },
    });
    signal.throwIfAborted();
    communityReads = await pendingReads;
    lifecycle.signal.throwIfAborted();
    signal.throwIfAborted();
    return { destroy, communityReads };
  } catch (error) {
    destroy();
    throw error;
  }
}
