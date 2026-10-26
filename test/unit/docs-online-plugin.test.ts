import { afterEach, describe, expect, it, vi } from 'vitest';
import { DOCS_COMMUNITY_PLUGIN_GUID } from '../../lib/docs-community-plugin';
import { connectDocsCommunityReads } from '../../lib/docs-community-plugin-host';

const barrierMethods = () => ({
  barrierCapability: vi.fn(async () => ({ available: false })),
  quiesceAndDrain: vi.fn(async () => {
    throw new Error('DOCS_BARRIER_UNAVAILABLE');
  }),
  barrierStatus: vi.fn(async () => {
    throw new Error('DOCS_BARRIER_UNAVAILABLE');
  }),
  resume: vi.fn(async () => {
    throw new Error('DOCS_BARRIER_UNAVAILABLE');
  }),
});
import { mountDocsOnlineEditor, parseDocsOnlineConfig, type DocsOnlineBinding } from '../../lib/docs-online-editor';
import pluginConfig from '../../public/docs-community-plugin/config.json';

vi.mock('../../lib/docs-community-plugin-host', () => ({ connectDocsCommunityReads: vi.fn() }));
const pluginUrl = new URL('/office-modules/config.json', location.href).href;
const binding = { documentKey: 'team-key', userId: 'member-a', communityPluginConfigUrl: pluginUrl };
const apiUrl = 'https://docs.example.test/web-apps/apps/api/documents/api.js';
function config() {
  return {
    documentType: 'slide',
    token: '',
    document: {
      key: binding.documentKey,
      fileType: 'pptx',
      title: 'Lesson.pptx',
      url: 'https://storage.example.test/source',
      permissions: { edit: true },
    },
    editorConfig: {
      mode: 'edit',
      callbackUrl: 'https://storage.example.test/callback',
      user: { id: binding.userId, name: 'Member A' },
      coEditing: { mode: 'fast', change: false },
      customization: { autosave: true },
      plugins: { autostart: [DOCS_COMMUNITY_PLUGIN_GUID], pluginsData: [pluginUrl] },
    },
  };
}
function signed(value = config()) {
  value.token = `fixture.${Buffer.from(JSON.stringify({ documentType: value.documentType, document: value.document, editorConfig: value.editorConfig, exp: 2000000000 })).toString('base64url')}.signature`;
  return value;
}
const parse = (value: unknown, hostBinding: DocsOnlineBinding = binding) =>
  parseDocsOnlineConfig(JSON.stringify(value), hostBinding);
afterEach(() => {
  document.body.replaceChildren();
  vi.resetAllMocks();
  vi.useRealTimers();
});

describe('Host-pinned native plugin configuration', () => {
  it('preserves exact signed fields and uses the installed plugin GUID', () => {
    const value = signed();
    expect(pluginConfig.guid).toBe(DOCS_COMMUNITY_PLUGIN_GUID);
    expect(parse(value)).toEqual(value);
  });
  it('accepts viewer mode without adding edit capability', () => {
    const value = config();
    value.document.permissions.edit = false;
    value.editorConfig.mode = 'view';
    expect(parse(signed(value)).document.permissions.edit).toBe(false);
  });
  it('decodes UTF-8 claims without rewriting the token', () => {
    const value = config();
    value.document.title = '\u8bfe\u7a0b.pptx';
    value.editorConfig.user.name = '\u8001\u5e08';
    const input = signed(value);
    expect(parse(input).token).toBe(input.token);
  });
  it('requires a separate Host pin, not a URL copied out of config', () => {
    expect(() => parse(signed(), { documentKey: binding.documentKey, userId: binding.userId })).toThrow();
  });
  it.each([
    ['extra GUID', (v: ReturnType<typeof config>) => v.editorConfig.plugins.autostart.push('asc.other')],
    ['wrong GUID', (v: ReturnType<typeof config>) => (v.editorConfig.plugins.autostart[0] = 'asc.other')],
    ['missing autostart', (v: ReturnType<typeof config>) => (v.editorConfig.plugins.autostart = [])],
    ['extra URL', (v: ReturnType<typeof config>) => v.editorConfig.plugins.pluginsData.push(pluginUrl)],
    [
      'wrong URL',
      (v: ReturnType<typeof config>) => (v.editorConfig.plugins.pluginsData[0] = 'https://untrusted.test/config.json'),
    ],
    ['missing URL', (v: ReturnType<typeof config>) => (v.editorConfig.plugins.pluginsData = [])],
    [
      'unknown plugin field',
      (v: ReturnType<typeof config>) => Object.assign(v.editorConfig.plugins, { url: pluginUrl }),
    ],
    ['null plugins', (v: ReturnType<typeof config>) => Object.assign(v.editorConfig, { plugins: null })],
    ['array plugins', (v: ReturnType<typeof config>) => Object.assign(v.editorConfig, { plugins: [] })],
    [
      'string autostart',
      (v: ReturnType<typeof config>) =>
        Object.assign(v.editorConfig.plugins, { autostart: DOCS_COMMUNITY_PLUGIN_GUID }),
    ],
  ])('rejects signed but unapproved %s', (_label, mutate) => {
    const value = config();
    mutate(value);
    expect(() => parse(signed(value))).toThrow(/^Invalid Docs online configuration:/);
  });
  it('rejects missing plugins when Host requires the bridge', () => {
    const value = signed();
    Reflect.deleteProperty(value.editorConfig, 'plugins');
    expect(() => parse(value)).toThrow();
  });
  it.each([
    '',
    'javascript:alert(1)',
    'http://remote.test/config.json',
    `${pluginUrl}?x=1`,
    `${pluginUrl}#x`,
    pluginUrl.replace('config.json', 'index.html'),
  ])('rejects an invalid trusted pin: %s', (pin) => {
    expect(() => parse(signed(), { ...binding, communityPluginConfigUrl: pin })).toThrow();
  });
  it('rejects plugin additions after signing', () => {
    const value = config();
    Reflect.deleteProperty(value.editorConfig, 'plugins');
    signed(value);
    value.editorConfig.plugins = config().editorConfig.plugins;
    expect(() => parse(value)).toThrow();
  });
  it('rejects plugin replacement after signing even when the outer config matches the Host pin', () => {
    const value = config();
    value.editorConfig.plugins.pluginsData[0] = 'https://old.test/config.json';
    signed(value);
    value.editorConfig.plugins.pluginsData[0] = pluginUrl;
    expect(() => parse(value)).toThrow();
  });
  it.each(['bad', 'null', '[]', '"text"', '{"editorConfig":null}'])(
    'rejects malformed claims without echoing the token: %s',
    (claims) => {
      const value = config();
      value.token = `fixture.${Buffer.from(claims).toString('base64url')}.signature`;
      expect(() => parse(value)).toThrow();
      try {
        parse(value);
      } catch (error) {
        expect(String(error)).not.toContain(value.token);
      }
    },
  );
});

async function launch(
  controller = new AbortController(),
  pin = pluginUrl,
  events = {},
  hostBinding: DocsOnlineBinding = { ...binding, communityPluginConfigUrl: pin },
) {
  const container = document.createElement('div');
  document.body.append(container);
  const value = config();
  value.editorConfig.plugins.pluginsData[0] = pin;
  const mounted = mountDocsOnlineEditor(
    container,
    new URL('/docs-online-host.html', location.href).href,
    apiUrl,
    JSON.stringify(signed(value)),
    hostBinding,
    controller.signal,
    events,
  );
  void mounted.catch(() => {});
  const frame = container.querySelector('iframe');
  if (frame) {
    frame.contentDocument!.write(
      '<!doctype html><meta name="docs-online-host" content="1"><div id="docs-online"></div>',
    );
    frame.dispatchEvent(new Event('load'));
    await Promise.resolve();
  }
  return { mounted, frame, script: frame?.contentDocument?.querySelector('script') };
}
function api(frame: HTMLIFrameElement, construct = vi.fn()) {
  const destroy = vi.fn();
  Object.assign(frame.contentWindow!, {
    DocsAPI: {
      DocEditor: class {
        destroyEditor = destroy;
        constructor(id: string, value: unknown) {
          construct(id, value);
        }
      },
    },
  });
  return destroy;
}

describe('online mount owns the Community read lifecycle', () => {
  it.each(['construction', 'settled handshake'])(
    'rejects native errors during %s before exposing a mount',
    async (phase) => {
      const close = vi.fn();
      let connected!: (client: Awaited<ReturnType<typeof connectDocsCommunityReads>>) => void;
      vi.mocked(connectDocsCommunityReads).mockImplementation(
        () =>
          new Promise((resolve) => {
            connected = resolve;
          }),
      );
      const onError = vi.fn();
      const run = await launch(undefined, pluginUrl, { onError });
      let native!: { onError(event: unknown): void };
      const destroy = api(
        run.frame!,
        vi.fn((_id, value) => {
          native = value.events;
          if (phase === 'construction') native.onError({ data: { errorCode: -21 } });
        }),
      );
      run.script!.dispatchEvent(new Event('load'));
      await Promise.resolve();
      connected({ ...barrierMethods(), readPresentation: vi.fn(), close });
      if (phase === 'settled handshake') native.onError({ data: { errorCode: -21 } });
      await expect(run.mounted).rejects.toMatchObject({ name: 'AbortError' });
      expect(close).toHaveBeenCalledTimes(1);
      expect(onError).toHaveBeenCalledTimes(1);
      expect(destroy).toHaveBeenCalledTimes(1);
      expect(run.frame!.isConnected).toBe(false);
    },
  );
  it('retains the validated plugin requirement across resource loading', async () => {
    const hostBinding: DocsOnlineBinding = { ...binding };
    const close = vi.fn();
    vi.mocked(connectDocsCommunityReads).mockResolvedValue({ ...barrierMethods(), readPresentation: vi.fn(), close });
    const run = await launch(undefined, pluginUrl, {}, hostBinding);
    hostBinding.communityPluginConfigUrl = undefined;
    api(run.frame!);
    run.script!.dispatchEvent(new Event('load'));
    const mounted = await run.mounted;
    expect(connectDocsCommunityReads).toHaveBeenCalledTimes(1);
    expect(mounted.communityReads).toBeDefined();
    mounted.destroy();
    expect(close).toHaveBeenCalledTimes(1);
  });
  it('closes the read channel before forwarding a native Docs error without discarding the editor', async () => {
    const close = vi.fn();
    vi.mocked(connectDocsCommunityReads).mockResolvedValue({ ...barrierMethods(), readPresentation: vi.fn(), close });
    const onError = vi.fn(() => expect(close).toHaveBeenCalledTimes(1));
    const run = await launch(undefined, pluginUrl, { onError });
    let native: { onError(event: unknown): void };
    const destroy = api(
      run.frame!,
      vi.fn((_id, value) => {
        native = value.events;
      }),
    );
    run.script!.dispatchEvent(new Event('load'));
    const mounted = await run.mounted;
    native!.onError({ data: { errorCode: -18 } });
    expect(onError).toHaveBeenCalledExactlyOnceWith({ data: { errorCode: -18 } });
    expect(destroy).not.toHaveBeenCalled();
    expect(run.frame!.isConnected).toBe(true);
    mounted.destroy();
    expect(destroy).toHaveBeenCalledTimes(1);
  });
  it('aborts a pending plugin handshake when native Docs rejects the session', async () => {
    vi.mocked(connectDocsCommunityReads).mockImplementation(
      (_window, _session, signal) =>
        new Promise((_resolve, reject) =>
          signal.addEventListener('abort', () => reject(new Error('Session unavailable')), { once: true }),
        ),
    );
    const onError = vi.fn();
    const run = await launch(undefined, pluginUrl, { onError });
    const destroy = api(
      run.frame!,
      vi.fn((_id, value) => value.events.onError({ data: { errorCode: -20 } })),
    );
    run.script!.dispatchEvent(new Event('load'));
    await expect(run.mounted).rejects.toThrow('Session unavailable');
    expect(onError).toHaveBeenCalledTimes(1);
    expect(destroy).toHaveBeenCalledTimes(1);
    expect(run.frame!.isConnected).toBe(false);
  });
  it('installs connection discovery before native autostart and preserves signed config', async () => {
    const readPresentation = vi.fn();
    const barriers = barrierMethods();
    const close = vi.fn();
    vi.mocked(connectDocsCommunityReads).mockResolvedValue({ ...barriers, readPresentation, close });
    const run = await launch();
    let connectionSignal: AbortSignal;
    const destroy = api(
      run.frame!,
      vi.fn((_id, value) => {
        expect(connectDocsCommunityReads).toHaveBeenCalledTimes(1);
        const [owned, sessionId, signal] = vi.mocked(connectDocsCommunityReads).mock.calls[0];
        expect(owned).toBe(run.frame!.contentWindow);
        expect(sessionId).toMatch(/^[\w-]+$/);
        connectionSignal = signal;
        expect(value.token).toBe(signed().token);
        expect(value.editorConfig).toEqual(config().editorConfig);
      }),
    );
    run.script!.dispatchEvent(new Event('load'));
    const mounted = await run.mounted;
    expect(mounted.communityReads?.readPresentation).toBe(readPresentation);
    expect(mounted.communityReads?.barrierCapability).toBe(barriers.barrierCapability);
    expect(mounted.communityReads?.quiesceAndDrain).toBe(barriers.quiesceAndDrain);
    expect(mounted.communityReads?.barrierStatus).toBe(barriers.barrierStatus);
    expect(mounted.communityReads?.resume).toBe(barriers.resume);
    expect(connectionSignal!.aborted).toBe(false);
    mounted.destroy();
    expect(connectionSignal!.aborted).toBe(true);
    expect(close).toHaveBeenCalled();
    expect(destroy).toHaveBeenCalledTimes(1);
    expect(run.frame!.isConnected).toBe(false);
  });
  it('waits for the plugin and cleans up connection failure', async () => {
    let fail!: (error: Error) => void;
    vi.mocked(connectDocsCommunityReads).mockImplementation(
      () =>
        new Promise((_resolve, reject) => {
          fail = reject;
        }),
    );
    const run = await launch();
    const destroy = api(run.frame!);
    run.script!.dispatchEvent(new Event('load'));
    await Promise.resolve();
    fail(new Error('Plugin timed out'));
    await expect(run.mounted).rejects.toThrow('Plugin timed out');
    expect(destroy).toHaveBeenCalledTimes(1);
    expect(run.frame!.isConnected).toBe(false);
  });
  it('aborts discovery and destroys the native editor when cancelled during handshake', async () => {
    vi.mocked(connectDocsCommunityReads).mockImplementation(
      (_window, _id, signal) =>
        new Promise((_resolve, reject) =>
          signal.addEventListener('abort', () => reject(new Error('Cancelled')), { once: true }),
        ),
    );
    const controller = new AbortController();
    const run = await launch(controller);
    const destroy = api(run.frame!);
    run.script!.dispatchEvent(new Event('load'));
    await Promise.resolve();
    controller.abort();
    await expect(run.mounted).rejects.toThrow();
    expect(destroy).toHaveBeenCalledTimes(1);
    expect(run.frame!.isConnected).toBe(false);
  });
  it('cancels discovery when native construction fails', async () => {
    let cancelled = false;
    vi.mocked(connectDocsCommunityReads).mockImplementation(
      (_window, _id, signal) =>
        new Promise((_resolve, reject) =>
          signal.addEventListener(
            'abort',
            () => {
              cancelled = true;
              reject(new Error('Cancelled'));
            },
            { once: true },
          ),
        ),
    );
    const run = await launch();
    api(
      run.frame!,
      vi.fn(() => {
        throw new Error('Construction failed');
      }),
    );
    run.script!.dispatchEvent(new Event('load'));
    await expect(run.mounted).rejects.toThrow('Construction failed');
    expect(cancelled).toBe(true);
    expect(run.frame!.isConnected).toBe(false);
  });
  it('closes a connected client on later cancellation', async () => {
    const close = vi.fn();
    vi.mocked(connectDocsCommunityReads).mockResolvedValue({ ...barrierMethods(), readPresentation: vi.fn(), close });
    const controller = new AbortController();
    const run = await launch(controller);
    const destroy = api(run.frame!);
    run.script!.dispatchEvent(new Event('load'));
    await run.mounted;
    controller.abort();
    expect(close).toHaveBeenCalled();
    expect(destroy).toHaveBeenCalledTimes(1);
  });
  it('rejects even a signed, pinned plugin hosted outside the top-level Host origin', async () => {
    const run = await launch(undefined, 'https://other.example.test/config.json');
    await expect(run.mounted).rejects.toThrow('plugin origin');
    expect(run.frame).toBe(null);
    expect(connectDocsCommunityReads).not.toHaveBeenCalled();
  });
});
