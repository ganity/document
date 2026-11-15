import {
  invokeDocsBarrier,
  isBarrierCapability,
  isBarrierId,
  isBarrierReceipt,
  isBarrierTool,
  isNativeBarrierError,
  type DocsNativeBarrierCapability,
  type DocsNativeBarrierReceipt,
} from './docs-community-barrier';
export const DOCS_COMMUNITY_PROTOCOL = 'deckwise-docs-read/1';
export const DOCS_COMMUNITY_PLUGIN_GUID = 'asc.{D1ECA135-517C-4585-990C-6E0F2B001322}';
const attachedPlugins = new WeakSet<CommunityPluginApi>();

export interface CommunityPluginApi {
  info: { editorType?: string };
  executeMethod(name: string, args: unknown[], callback: (result: unknown) => void): void;
  callCommand(command: () => unknown, close: boolean, recalculate: boolean, callback: (result: unknown) => void): void;
}

interface Presentation {
  GetCurSlideIndex(): number;
  GetSlidesCount(): number;
  GetSlideByIndex(index: number): {
    GetAllShapes(): { GetDocContent(): { GetText(): string } | null }[];
  };
}
declare const Api: { GetPresentation(): Presentation };
declare const Asc: { scope: Record<string, unknown> };

export interface DocsCommunityReadResult {
  page: number;
  totalSlides: number;
  text: string;
  truncated: boolean;
  slidesRead: number;
}

function object(value: unknown): value is Record<string, unknown> {
  return !!value && typeof value === 'object' && !Array.isArray(value);
}

function validResult(value: unknown): value is DocsCommunityReadResult {
  return (
    object(value) &&
    typeof value.text === 'string' &&
    value.text.length <= 8000 &&
    typeof value.truncated === 'boolean' &&
    Number.isSafeInteger(value.page) &&
    Number(value.page) >= 1 &&
    Number.isSafeInteger(value.totalSlides) &&
    Number(value.totalSlides) >= Number(value.page) &&
    Number.isSafeInteger(value.slidesRead) &&
    Number(value.slidesRead) >= 0 &&
    Number(value.slidesRead) <= Number(value.totalSlides)
  );
}

/** Port ownership must be established by Host; a port/session is not server authorization. */
export function serveDocsCommunityReads(
  port: MessagePort,
  plugin: CommunityPluginApi,
  sessionId: string,
  signal: AbortSignal,
): void {
  if (!/^[\w-]{1,128}$/.test(sessionId) || plugin.info.editorType !== 'slide')
    throw new Error('Invalid Docs plugin session');
  if (attachedPlugins.has(plugin)) throw new Error('Docs plugin channel already attached');
  signal.throwIfAborted();
  attachedPlugins.add(plugin);
  let active = true;
  let busy = false;
  const used = new Set<string>();
  const close = () => {
    active = false;
    port.onmessage = null;
    port.close();
    signal.removeEventListener('abort', close);
    attachedPlugins.delete(plugin);
  };
  signal.addEventListener('abort', close, { once: true });
  port.onmessage = (event) => {
    const request: unknown = event.data;
    if (
      !object(request) ||
      request.protocol !== DOCS_COMMUNITY_PROTOCOL ||
      request.sessionId !== sessionId ||
      typeof request.requestId !== 'string' ||
      !/^[\w-]{1,128}$/.test(request.requestId)
    )
      return;
    const reply = (result?: unknown, error?: string) => {
      if (active)
        port.postMessage({ protocol: DOCS_COMMUNITY_PROTOCOL, sessionId, requestId: request.requestId, result, error });
    };
    if (busy) {
      reply(undefined, 'DOCS_PLUGIN_BUSY');
      return;
    }
    if (used.has(request.requestId)) {
      reply(undefined, 'DOCS_PLUGIN_REQUEST_USED');
      return;
    }
    if (used.size >= 4096) {
      reply(undefined, 'DOCS_PLUGIN_SESSION_LIMIT');
      return;
    }
    used.add(request.requestId);
    if (isBarrierTool(request.tool)) {
      const capability = request.tool === 'native_barrier_capability';
      const keys = ['protocol', 'sessionId', 'requestId', 'tool', ...(capability ? [] : ['barrierId'])];
      if (
        Object.keys(request).length !== keys.length ||
        Object.keys(request).some((key) => !keys.includes(key)) ||
        (!capability && !isBarrierId(request.barrierId))
      ) {
        reply(undefined, 'DOCS_BARRIER_INVALID_REQUEST');
        return;
      }
      busy = true;
      let completed = false;
      invokeDocsBarrier(
        plugin,
        request.tool,
        capability ? undefined : (request.barrierId as string),
        (result, error) => {
          if (completed) return;
          completed = true;
          busy = false;
          reply(result, error);
        },
      );
      return;
    }
    if (
      Object.keys(request).some((key) => !['protocol', 'sessionId', 'requestId', 'tool', 'maxChars'].includes(key)) ||
      request.tool !== 'get_presentation_text' ||
      !Number.isInteger(request.maxChars) ||
      Number(request.maxChars) < 1 ||
      Number(request.maxChars) > 8000
    ) {
      reply(undefined, 'DOCS_PLUGIN_UNSUPPORTED_REQUEST');
      return;
    }
    busy = true;
    try {
      // Only data crosses the scope boundary; callers cannot provide executable code.
      Object.assign(pluginScope(), { deckwiseReadMaxChars: request.maxChars });
      plugin.callCommand(
        function () {
          try {
            const max = Asc.scope.deckwiseReadMaxChars;
            if (!Number.isInteger(max) || Number(max) < 1 || Number(max) > 8000) return null;
            const presentation = Api.GetPresentation();
            const totalSlides = presentation.GetSlidesCount();
            const page = presentation.GetCurSlideIndex() + 1;
            if (totalSlides < 1 || totalSlides > 1000 || page < 1 || page > totalSlides) return null;
            let text = '';
            let slidesRead = 0;
            let truncated = false;
            for (let index = 0; index < totalSlides; index++) {
              const shapes = presentation.GetSlideByIndex(index).GetAllShapes();
              if (shapes.length > 4096) return null;
              for (const shape of shapes) {
                const value = shape.GetDocContent()?.GetText();
                if (value === undefined) continue;
                if (typeof value !== 'string') return null;
                const part = `${text ? '\n' : ''}${value}`;
                if (part.length > Number(max) - text.length) {
                  text += part.slice(0, Number(max) - text.length);
                  truncated = true;
                  break;
                }
                text += part;
              }
              slidesRead++;
              if (truncated) break;
            }
            return { page, totalSlides, text, truncated, slidesRead };
          } catch {
            return null;
          }
        },
        false,
        false,
        (value) => {
          busy = false;
          // Docs can skip commands while a collaborative edit holds its native action lock.
          if (value === undefined) reply(undefined, 'DOCS_PLUGIN_UNAVAILABLE');
          else if (validResult(value) && value.text.length <= Number(request.maxChars)) reply(value);
          else reply(undefined, 'DOCS_PLUGIN_INVALID_NATIVE_RESULT');
        },
      );
    } catch {
      busy = false;
      reply(undefined, 'DOCS_PLUGIN_NATIVE_FAILED');
    }
  };
  port.start();
}

function pluginScope(): Record<string, unknown> {
  // This is the plugin's own Asc.scope, not an SDK window escape.
  const asc = (globalThis as unknown as { Asc?: { scope?: Record<string, unknown> } }).Asc;
  if (!asc) throw new Error('Missing native plugin scope');
  return asc.scope ?? (asc.scope = {});
}

/** One owned queue. Read timeout closes; barrier timeout permits explicit status reconciliation. */
export function createDocsCommunityReadClient(port: MessagePort, sessionId: string, signal: AbortSignal) {
  if (!/^[\w-]{1,128}$/.test(sessionId)) throw new Error('Invalid Docs plugin session');
  signal.throwIfAborted();
  let active = true;
  let pending:
    | {
        id: string;
        kind: 'read' | 'barrier';
        validate: (value: unknown) => boolean;
        resolve: (result: unknown) => void;
        reject: (error: Error) => void;
        timer: ReturnType<typeof setTimeout>;
      }
    | undefined;
  const close = () => {
    active = false;
    port.onmessage = null;
    port.close();
    signal.removeEventListener('abort', close);
    if (pending) {
      clearTimeout(pending.timer);
      pending.reject(new Error('Docs plugin channel closed; reconnect explicitly'));
      pending = undefined;
    }
  };
  signal.addEventListener('abort', close, { once: true });
  port.onmessage = (event) => {
    const value: unknown = event.data;
    if (
      !pending ||
      !object(value) ||
      value.protocol !== DOCS_COMMUNITY_PROTOCOL ||
      value.sessionId !== sessionId ||
      value.requestId !== pending.id
    )
      return;
    const current = pending;
    pending = undefined;
    clearTimeout(current.timer);
    const recoverable =
      value.error === 'DOCS_PLUGIN_BUSY' ||
      (current.kind === 'read'
        ? value.error === 'DOCS_PLUGIN_UNAVAILABLE'
        : value.error === 'DOCS_BARRIER_UNAVAILABLE' ||
          value.error === 'DOCS_PLUGIN_UNSUPPORTED_REQUEST' ||
          isNativeBarrierError(value.error));
    if (recoverable && value.result === undefined) {
      current.reject(new Error(String(value.error)));
    } else if (value.error !== undefined || !current.validate(value.result)) {
      current.reject(new Error(current.kind === 'read' ? 'Docs plugin read failed' : 'DOCS_BARRIER_INVALID_RESPONSE'));
      close();
    } else current.resolve(value.result);
  };
  port.start();
  const request = <Result>(
    data: Record<string, unknown>,
    kind: 'read' | 'barrier',
    validate: (value: unknown) => boolean,
  ): Promise<Result> => {
    if (!active) return Promise.reject(new Error('Docs plugin channel closed; reconnect explicitly'));
    if (pending) return Promise.reject(new Error('Docs plugin request in progress'));
    const id = crypto.randomUUID();
    return new Promise((resolve, reject) => {
      const timeout = () => {
        if (kind === 'read') {
          close();
          return;
        }
        if (pending?.id !== id) return;
        pending = undefined;
        reject(new Error('DOCS_BARRIER_ACK_UNKNOWN'));
      };
      pending = {
        id,
        kind,
        validate,
        resolve: (value) => resolve(value as Result),
        reject,
        timer: setTimeout(timeout, 30000),
      };
      try {
        port.postMessage({ protocol: DOCS_COMMUNITY_PROTOCOL, sessionId, requestId: id, ...data });
      } catch {
        close();
      }
    });
  };
  return {
    readPresentation(maxChars = 8000): Promise<DocsCommunityReadResult> {
      if (!Number.isInteger(maxChars) || maxChars < 1 || maxChars > 8000)
        return Promise.reject(new Error('Invalid Docs read limit'));
      return request(
        { tool: 'get_presentation_text', maxChars },
        'read',
        (value) => validResult(value) && value.text.length <= maxChars,
      );
    },
    barrierCapability(): Promise<DocsNativeBarrierCapability> {
      return request({ tool: 'native_barrier_capability' }, 'barrier', isBarrierCapability);
    },
    quiesceAndDrain(barrierId: string): Promise<DocsNativeBarrierReceipt> {
      if (!isBarrierId(barrierId)) return Promise.reject(new Error('DOCS_BARRIER_INVALID_REQUEST'));
      return request({ tool: 'native_barrier_quiesce', barrierId }, 'barrier', (value) =>
        isBarrierReceipt(value, barrierId),
      );
    },
    barrierStatus(barrierId: string): Promise<DocsNativeBarrierReceipt> {
      if (!isBarrierId(barrierId)) return Promise.reject(new Error('DOCS_BARRIER_INVALID_REQUEST'));
      return request({ tool: 'native_barrier_status', barrierId }, 'barrier', (value) =>
        isBarrierReceipt(value, barrierId),
      );
    },
    resume(barrierId: string): Promise<DocsNativeBarrierReceipt> {
      if (!isBarrierId(barrierId)) return Promise.reject(new Error('DOCS_BARRIER_INVALID_REQUEST'));
      return request({ tool: 'native_barrier_resume', barrierId }, 'barrier', (value) =>
        isBarrierReceipt(value, barrierId),
      );
    },
    close,
  };
}
