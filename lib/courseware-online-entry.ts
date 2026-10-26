import { isBarrierCapability, isBarrierId, isBarrierReceipt, isNativeBarrierError } from './docs-community-barrier';
import { mountDocsOnlineEditor, type DocsOnlineBinding } from './docs-online-editor';

export const COURSEWARE_ONLINE_PROTOCOL = 'deckwise-docs-online/1';
type Mount = typeof mountDocsOnlineEditor;
interface Pins {
  parentWindow: Window;
  parentOrigin: string;
  channelId: string;
  apiUrl: string;
  communityPluginConfigUrl?: string;
}
const identifier = (value: unknown): value is string => typeof value === 'string' && /^[\w-]{1,128}$/.test(value);
const object = (value: unknown): value is Record<string, unknown> =>
  !!value && typeof value === 'object' && !Array.isArray(value);
const exact = (value: unknown, keys: string[]): value is Record<string, unknown> =>
  object(value) && Object.keys(value).length === keys.length && keys.every((key) => Object.hasOwn(value, key));

/** Installs only on a Host-selected Office origin; no offline bridge or dynamic code commands. */
export function installCoursewareOnlineEntry(container: HTMLElement, pins: Pins, mount: Mount = mountDocsOnlineEditor) {
  if (pins.parentWindow === window || !identifier(pins.channelId)) throw new Error('Invalid online parent');
  const parent = new URL(pins.parentOrigin);
  if (
    pins.parentOrigin !== `${parent.protocol}//${parent.host}` ||
    !['', '/'].includes(parent.pathname) ||
    parent.search ||
    parent.hash ||
    parent.username ||
    parent.password ||
    !(
      parent.protocol === 'https:' ||
      (parent.protocol === 'http:' && ['127.0.0.1', 'localhost', '[::1]'].includes(parent.hostname)) ||
      (parent.protocol === 'deckwise-editor:' && parent.hostname === 'host' && !parent.port)
    )
  )
    throw new Error('Invalid online parent origin');
  let active = true;
  let generation = 0;
  let current:
    | {
        generation: number;
        abort: AbortController;
        binding: DocsOnlineBinding;
        projectId: string;
        sessionEpoch: number;
        connection?: Awaited<ReturnType<Mount>>;
      }
    | undefined;
  const requests = new Set<string>();
  const post = (message: object) => {
    if (active)
      pins.parentWindow.postMessage(
        { protocol: COURSEWARE_ONLINE_PROTOCOL, channelId: pins.channelId, ...message },
        pins.parentOrigin,
      );
  };
  const stop = () => {
    current?.abort.abort();
    current?.connection?.destroy();
    current = undefined;
  };
  const receive = (event: MessageEvent) => {
    const data: unknown = event.data;
    if (
      !active ||
      event.source !== pins.parentWindow ||
      event.origin !== pins.parentOrigin ||
      !object(data) ||
      data.protocol !== COURSEWARE_ONLINE_PROTOCOL ||
      data.channelId !== pins.channelId ||
      !identifier(data.requestId)
    )
      return;
    if (requests.has(data.requestId) || requests.size >= 4096) {
      post({ type: 'error', requestId: data.requestId, code: 'ONLINE_REQUEST_REUSED' });
      return;
    }
    requests.add(data.requestId);
    const reply = (message: object) => post({ requestId: data.requestId, ...message });
    if (data.type === 'close' && exact(data, ['protocol', 'channelId', 'requestId', 'type', 'generation'])) {
      if (!current || data.generation !== current.generation) {
        reply({ type: 'error', code: 'ONLINE_GENERATION_STALE' });
        return;
      }
      stop();
      reply({ type: 'closed', generation: data.generation });
      return;
    }
    if (data.type === 'read' && exact(data, ['protocol', 'channelId', 'requestId', 'type', 'generation', 'maxChars'])) {
      const state = current;
      if (!state || data.generation !== state.generation) {
        reply({ type: 'error', code: 'ONLINE_GENERATION_STALE' });
        return;
      }
      if (
        !Number.isInteger(data.maxChars) ||
        Number(data.maxChars) < 1 ||
        Number(data.maxChars) > 8000 ||
        !state.connection?.communityReads
      ) {
        reply({ type: 'error', code: 'ONLINE_TEXT_READ_UNAVAILABLE' });
        return;
      }
      void state.connection.communityReads.readPresentation(Number(data.maxChars)).then(
        (result) => {
          if (current === state && !state.abort.signal.aborted)
            reply({ type: 'text', generation: state.generation, result });
        },
        () => {
          if (current === state && !state.abort.signal.aborted)
            reply({ type: 'error', code: 'ONLINE_TEXT_READ_FAILED' });
        },
      );
      return;
    }
    if (
      typeof data.type === 'string' &&
      ['barrier-capability', 'barrier-quiesce', 'barrier-status', 'barrier-resume'].includes(data.type)
    ) {
      const capability = data.type === 'barrier-capability';
      if (
        !exact(data, [
          'protocol',
          'channelId',
          'requestId',
          'type',
          'generation',
          ...(capability ? [] : ['barrierId']),
        ]) ||
        (!capability && !isBarrierId(data.barrierId))
      ) {
        reply({ type: 'error', code: 'ONLINE_INVALID_REQUEST' });
        return;
      }
      const state = current;
      if (!state || data.generation !== state.generation) {
        reply({ type: 'error', code: 'ONLINE_GENERATION_STALE' });
        return;
      }
      const respond = (message: object) => {
        if (current === state && !state.abort.signal.aborted)
          reply({
            generation: state.generation,
            projectId: state.projectId,
            sessionEpoch: state.sessionEpoch,
            documentKey: state.binding.documentKey,
            docsUserId: state.binding.userId,
            ...message,
          });
      };
      const bridge = state.connection?.communityReads;
      if (!bridge) {
        respond(
          capability
            ? { type: 'barrier-capability', result: { available: false } }
            : { type: 'error', code: 'ONLINE_NATIVE_BARRIER_UNAVAILABLE' },
        );
        return;
      }
      void Promise.resolve()
        .then(async (): Promise<unknown> => {
          state.abort.signal.throwIfAborted();
          if (capability) return bridge.barrierCapability();
          if (data.type === 'barrier-quiesce') return bridge.quiesceAndDrain(data.barrierId as string);
          if (data.type === 'barrier-resume') return bridge.resume(data.barrierId as string);
          return bridge.barrierStatus(data.barrierId as string);
        })
        .then(
          (result) => {
            if (!(capability ? isBarrierCapability(result) : isBarrierReceipt(result, data.barrierId as string))) {
              respond({ type: 'error', code: 'ONLINE_NATIVE_BARRIER_FAILED' });
              return;
            }
            respond({ type: capability ? 'barrier-capability' : 'barrier', result });
          },
          (error: unknown) => {
            const message = error instanceof Error ? error.message : '';
            const code = isNativeBarrierError(message)
              ? message
              : ['DOCS_BARRIER_UNAVAILABLE', 'DOCS_PLUGIN_UNSUPPORTED_REQUEST'].includes(message)
                ? 'ONLINE_NATIVE_BARRIER_UNAVAILABLE'
                : message === 'DOCS_BARRIER_ACK_UNKNOWN'
                  ? 'ONLINE_NATIVE_BARRIER_ACK_UNKNOWN'
                  : ['Docs plugin request in progress', 'DOCS_PLUGIN_BUSY'].includes(message)
                    ? 'ONLINE_NATIVE_BARRIER_BUSY'
                    : 'ONLINE_NATIVE_BARRIER_FAILED';
            respond({ type: 'error', code });
          },
        );
      return;
    }
    if (
      data.type !== 'open' ||
      !exact(data, [
        'protocol',
        'channelId',
        'requestId',
        'type',
        'generation',
        'projectId',
        'sessionEpoch',
        'documentKey',
        'docsUserId',
        'configJson',
        'expiresAtEpochMillis',
      ]) ||
      !Number.isSafeInteger(data.generation) ||
      Number(data.generation) <= generation ||
      typeof data.projectId !== 'string' ||
      !/^[a-f0-9]{8}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{12}$/.test(data.projectId) ||
      !Number.isSafeInteger(data.sessionEpoch) ||
      Number(data.sessionEpoch) < 1 ||
      typeof data.documentKey !== 'string' ||
      !/^[A-Za-z0-9_.=-]{1,128}$/.test(data.documentKey) ||
      typeof data.docsUserId !== 'string' ||
      data.docsUserId.length > 128 ||
      data.docsUserId.trim() !== data.docsUserId ||
      !/^tenant-[1-9]\d*-user-[1-9]\d*(?:-client-[a-f0-9]{8}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{12})?$/.test(
        data.docsUserId,
      ) ||
      typeof data.configJson !== 'string' ||
      data.configJson.length > 65536 ||
      !Number.isSafeInteger(data.expiresAtEpochMillis) ||
      Number(data.expiresAtEpochMillis) <= Date.now() ||
      Number(data.expiresAtEpochMillis) > Date.now() + 120000
    ) {
      reply({ type: 'error', code: 'ONLINE_INVALID_REQUEST' });
      return;
    }
    generation = Number(data.generation);
    stop();
    const state = {
      generation,
      abort: new AbortController(),
      projectId: data.projectId,
      sessionEpoch: Number(data.sessionEpoch),
      binding: {
        documentKey: data.documentKey,
        userId: data.docsUserId,
        ...(pins.communityPluginConfigUrl === undefined
          ? {}
          : { communityPluginConfigUrl: pins.communityPluginConfigUrl }),
      },
      connection: undefined as Awaited<ReturnType<Mount>> | undefined,
    };
    current = state;
    const emit = (message: object) => {
      if (current === state && !state.abort.signal.aborted)
        reply({
          generation: state.generation,
          projectId: state.projectId,
          sessionEpoch: state.sessionEpoch,
          documentKey: state.binding.documentKey,
          docsUserId: state.binding.userId,
          ...message,
        });
    };
    void Promise.resolve()
      .then(() => {
        state.abort.signal.throwIfAborted();
        return mount(
          container,
          new URL('docs-online-host.html', location.href).href,
          pins.apiUrl,
          data.configJson as string,
          state.binding,
          state.abort.signal,
          {
            onDocumentReady: () => emit({ type: 'ready' }),
            onDocumentStateChange: (event) => emit({ type: 'state', dirty: event.data }),
            onError: () => {
              emit({ type: 'error', code: 'ONLINE_NATIVE_ERROR' });
              if (current === state) stop();
            },
          },
        );
      })
      .then(
        (connection) => {
          if (current !== state || state.abort.signal.aborted) {
            connection.destroy();
            return;
          }
          state.connection = connection;
          emit({ type: 'mounted', canReadText: !!connection.communityReads });
        },
        () => {
          emit({ type: 'error', code: 'ONLINE_MOUNT_FAILED' });
          if (current === state) stop();
        },
      );
  };
  window.addEventListener('message', receive);
  post({ type: 'entry-ready' });
  return () => {
    active = false;
    stop();
    window.removeEventListener('message', receive);
  };
}
