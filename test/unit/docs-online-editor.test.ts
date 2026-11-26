import { afterEach, describe, expect, it, vi } from 'vitest';
import { mountDocsOnlineEditor, parseDocsOnlineConfig } from '../../lib/docs-online-editor';

const binding = { documentKey: 'team-project-r7-session3', userId: 'member-a' };
const apiUrl = 'https://docs.example.test/web-apps/apps/api/documents/api.js';
const config = () => ({
  documentType: 'slide',
  token: 'fixture.header.signature',
  document: {
    key: binding.documentKey,
    fileType: 'pptx',
    title: 'Lesson.pptx',
    url: 'https://storage.example.test/source?download=fixture',
    permissions: { edit: true },
  },
  editorConfig: {
    mode: 'edit',
    callbackUrl: 'https://storage.example.test/docs/callback?session=fixture',
    user: { id: binding.userId, name: 'Member A' },
    coEditing: { mode: 'fast', change: false },
    customization: { autosave: true },
  },
});

afterEach(() => {
  document.body.replaceChildren();
  vi.useRealTimers();
});

describe('Docs online launch shape, not JWT verification', () => {
  it('preserves the entire signed config and shares one server key between distinct users', () => {
    const a = config();
    const b = config();
    b.editorConfig.user = { id: 'member-b', name: 'Member B' };
    expect(parseDocsOnlineConfig(JSON.stringify(a), binding)).toEqual(a);
    expect(parseDocsOnlineConfig(JSON.stringify(b), { ...binding, userId: 'member-b' }).document.key).toBe(
      a.document.key,
    );
  });

  it('preserves a registered client suffix and rejects another window even for the same tenant user', () => {
    const value = config();
    const userId = 'tenant-1-user-2-client-d5820881-b533-4d0b-aafe-d0129c660fb8';
    value.editorConfig.user.id = userId;
    expect(parseDocsOnlineConfig(JSON.stringify(value), { ...binding, userId })).toEqual(value);
    expect(() =>
      parseDocsOnlineConfig(JSON.stringify(value), {
        ...binding,
        userId: 'tenant-1-user-2-client-d5820881-b533-4d0b-aafe-d0129c660fb9',
      }),
    ).toThrow('Invalid Docs online configuration: user');
    expect(() => parseDocsOnlineConfig(JSON.stringify(value), { ...binding, userId: 'tenant-1-user-2' })).toThrow(
      'Invalid Docs online configuration: user',
    );
  });

  it('keeps viewer permission and mode', () => {
    const value = config();
    value.document.permissions.edit = false;
    value.editorConfig.mode = 'view';
    expect(parseDocsOnlineConfig(JSON.stringify(value), binding).document.permissions.edit).toBe(false);
  });

  it.each([
    [
      'wrong document',
      (v: ReturnType<typeof config>) => {
        v.document.key = 'other';
      },
    ],
    [
      'wrong user',
      (v: ReturnType<typeof config>) => {
        v.editorConfig.user.id = 'other';
      },
    ],
    [
      'wrong mode',
      (v: ReturnType<typeof config>) => {
        v.editorConfig.mode = 'view';
      },
    ],
    [
      'unsigned',
      (v: ReturnType<typeof config>) => {
        v.token = '';
      },
    ],
    [
      'blob source',
      (v: ReturnType<typeof config>) => {
        v.document.url = 'blob:https://host.test/source';
      },
    ],
    [
      'unencrypted source',
      (v: ReturnType<typeof config>) => {
        v.document.url = 'http://storage.example.test/source';
      },
    ],
    [
      'callback credentials',
      (v: ReturnType<typeof config>) => {
        v.editorConfig.callbackUrl = 'https://secret:password@host.test/callback';
      },
    ],
    [
      'script callback',
      (v: ReturnType<typeof config>) => {
        v.editorConfig.callbackUrl = 'javascript:alert(1)';
      },
    ],
    [
      'wrong format',
      (v: ReturnType<typeof config>) => {
        v.document.fileType = 'docx';
      },
    ],
    [
      'bad key',
      (v: ReturnType<typeof config>) => {
        v.document.key = 'invalid/key';
      },
    ],
    [
      'empty name',
      (v: ReturnType<typeof config>) => {
        v.editorConfig.user.name = ' ';
      },
    ],
    [
      'wrong coediting mode',
      (v: ReturnType<typeof config>) => {
        v.editorConfig.coEditing.mode = 'strict';
      },
    ],
    [
      'null document',
      (v: ReturnType<typeof config>) => {
        Object.assign(v, { document: null });
      },
    ],
    [
      'local save event',
      (v: ReturnType<typeof config>) => {
        Object.assign(v, { events: { onSave: 'local' } });
      },
    ],
    [
      'local coauthoring patch',
      (v: ReturnType<typeof config>) => {
        Object.assign(v.editorConfig, { canCoAuthoring: false });
      },
    ],
    [
      'disabled autosave',
      (v: ReturnType<typeof config>) => {
        Object.assign(v.editorConfig.customization, { autosave: false });
      },
    ],
  ])('rejects %s without echoing secrets', (_label, mutate) => {
    const value = config();
    mutate(value);
    expect(() => parseDocsOnlineConfig(JSON.stringify(value), binding)).toThrow(/^Invalid Docs online configuration:/);
    try {
      parseDocsOnlineConfig(JSON.stringify(value), binding);
    } catch (error) {
      expect(String(error)).not.toContain('password');
      if (value.token) expect(String(error)).not.toContain(value.token);
    }
  });

  it.each(['null', '[]', '{', ' '.repeat(65537)])('rejects malformed or oversized input', (json) => {
    expect(() => parseDocsOnlineConfig(json, binding)).toThrow();
  });

  it('rejects the offline default identity even with a matching binding', () => {
    const value = config();
    value.editorConfig.user.id = 'local-user';
    expect(() => parseDocsOnlineConfig(JSON.stringify(value), { ...binding, userId: 'local-user' })).toThrow();
  });
});

async function launch(signal = new AbortController().signal, url = apiUrl) {
  const container = document.createElement('div');
  document.body.append(container);
  const ready = vi.fn();
  const dirty = vi.fn();
  const error = vi.fn();
  const mounted = mountDocsOnlineEditor(
    container,
    new URL('/docs-online-host.html', window.location.href).href,
    url,
    JSON.stringify(config()),
    binding,
    signal,
    {
      onDocumentReady: ready,
      onDocumentStateChange: dirty,
      onError: error,
    },
  );
  const frame = container.querySelector('iframe');
  void mounted.catch(() => {});
  if (frame) {
    frame.contentDocument!.write(
      '<!doctype html><meta name="docs-online-host" content="1"><body><div id="docs-online"></div></body>',
    );
    frame.dispatchEvent(new Event('load'));
    await Promise.resolve();
  }
  const script = frame?.contentDocument?.querySelector('script');
  return { mounted, frame, script, ready, dirty, error };
}

function fakeApi(frame: HTMLIFrameElement, construct = vi.fn()) {
  const destroyEditor = vi.fn();
  Object.assign(frame.contentWindow!, {
    DocsAPI: {
      DocEditor: class {
        destroyEditor = destroyEditor;
        constructor(id: string, value: unknown) {
          construct(id, value);
        }
      },
    },
  });
  return { destroyEditor, construct };
}

describe('isolated online editor lifecycle', () => {
  it.each([
    'https://other.example.test/docs-online-host.html',
    'blob:https://host.test/page',
    new URL('/index.html', window.location.href).href,
  ])('rejects an invalid host document URL before mounting %s', async (hostUrl) => {
    const container = document.createElement('div');
    document.body.append(container);
    await expect(
      mountDocsOnlineEditor(
        container,
        hostUrl,
        apiUrl,
        JSON.stringify(config()),
        binding,
        new AbortController().signal,
      ),
    ).rejects.toThrow();
    expect(container.children).toHaveLength(0);
  });

  it.each(['error', 'timeout', 'missing marker', 'abort'])('cleans up a failed host load: %s', async (failure) => {
    vi.useFakeTimers();
    const container = document.createElement('div');
    document.body.append(container);
    const controller = new AbortController();
    const mounted = mountDocsOnlineEditor(
      container,
      new URL('/docs-online-host.html', window.location.href).href,
      apiUrl,
      JSON.stringify(config()),
      binding,
      controller.signal,
    );
    const rejected = expect(mounted).rejects.toThrow();
    const frame = container.querySelector('iframe')!;
    if (failure === 'timeout') await vi.advanceTimersByTimeAsync(30000);
    else if (failure === 'abort') controller.abort();
    else frame.dispatchEvent(new Event(failure === 'error' ? 'error' : 'load'));
    await rejected;
    expect(frame.isConnected).toBe(false);
    expect(vi.getTimerCount()).toBe(0);
  });

  it('uses only the child Docs API, leaves offline globals alone and disposes once', async () => {
    const offline = { DocEditor: vi.fn() };
    const previous = window.DocsAPI;
    Object.assign(window, { DocsAPI: offline });
    try {
      const run = await launch();
      expect(run.script!.src).toBe(apiUrl);
      const fake = fakeApi(run.frame!);
      run.script!.dispatchEvent(new Event('load'));
      const mounted = await run.mounted;
      expect(fake.construct).toHaveBeenCalledWith('docs-online', {
        ...config(),
        events: {
          onDocumentReady: expect.any(Function),
          onDocumentStateChange: expect.any(Function),
          onError: expect.any(Function),
        },
      });
      expect(offline.DocEditor).not.toHaveBeenCalled();
      expect(window.DocsAPI).toBe(offline);
      expect(run.ready).not.toHaveBeenCalled();
      mounted.destroy();
      mounted.destroy();
      expect(fake.destroyEditor).toHaveBeenCalledTimes(1);
      expect(run.frame!.isConnected).toBe(false);
    } finally {
      window.DocsAPI = previous;
    }
  });

  it('loads the Docs API through the DeckWise reverse proxy path', async () => {
    const url = 'https://trainwise.efrobot.com/deckwise-office/web-apps/apps/api/documents/api.js';
    const run = await launch(undefined, url);
    expect(run.script!.src).toBe(url);
    const native = fakeApi(run.frame!);
    run.script!.dispatchEvent(new Event('load'));
    (await run.mounted).destroy();
    expect(native.destroyEditor).toHaveBeenCalledTimes(1);
  });

  it('cleans an aborted load and never creates a late editor', async () => {
    const controller = new AbortController();
    const run = await launch(controller.signal);
    const fake = fakeApi(run.frame!);
    controller.abort();
    await expect(run.mounted).rejects.toThrow();
    run.script!.dispatchEvent(new Event('load'));
    expect(fake.construct).not.toHaveBeenCalled();
    expect(run.frame!.isConnected).toBe(false);
  });

  it('cleans the editor when abort occurs inside construction', async () => {
    const controller = new AbortController();
    const run = await launch(controller.signal);
    const fake = fakeApi(
      run.frame!,
      vi.fn(() => controller.abort()),
    );
    run.script!.dispatchEvent(new Event('load'));
    await expect(run.mounted).rejects.toThrow();
    expect(fake.destroyEditor).toHaveBeenCalledTimes(1);
    expect(run.frame!.isConnected).toBe(false);
  });

  it('destroys on a later abort', async () => {
    const controller = new AbortController();
    const run = await launch(controller.signal);
    const fake = fakeApi(run.frame!);
    run.script!.dispatchEvent(new Event('load'));
    await run.mounted;
    controller.abort();
    expect(fake.destroyEditor).toHaveBeenCalledTimes(1);
    expect(run.frame!.isConnected).toBe(false);
  });

  it('ignores late readiness, dirty and error events after destruction', async () => {
    const run = await launch();
    let callbacks: {
      onDocumentReady: () => void;
      onDocumentStateChange: (event: { data: boolean }) => void;
      onError: (event: unknown) => void;
    };
    fakeApi(
      run.frame!,
      vi.fn((_id, value) => {
        callbacks = value.events;
      }),
    );
    run.script!.dispatchEvent(new Event('load'));
    const mounted = await run.mounted;
    callbacks!.onDocumentReady();
    callbacks!.onDocumentStateChange({ data: true });
    callbacks!.onError({ errorCode: 'current' });
    expect(run.ready).toHaveBeenCalledTimes(1);
    expect(run.dirty).toHaveBeenCalledExactlyOnceWith({ data: true });
    expect(run.error).toHaveBeenCalledExactlyOnceWith({ errorCode: 'current' });
    mounted.destroy();
    callbacks!.onDocumentReady();
    callbacks!.onDocumentStateChange({ data: false });
    callbacks!.onError({ errorCode: 'late' });
    expect(run.ready).toHaveBeenCalledTimes(1);
    expect(run.dirty).toHaveBeenCalledTimes(1);
    expect(run.error).toHaveBeenCalledTimes(1);
  });

  it('removes the frame even when native destruction throws', async () => {
    const run = await launch();
    const fake = fakeApi(run.frame!);
    fake.destroyEditor.mockImplementation(() => {
      throw new Error('Destroy failed');
    });
    run.script!.dispatchEvent(new Event('load'));
    const mounted = await run.mounted;
    expect(mounted.destroy).toThrow('Destroy failed');
    expect(run.frame!.isConnected).toBe(false);
    mounted.destroy();
    expect(fake.destroyEditor).toHaveBeenCalledTimes(1);
  });

  it('does not mount an already cancelled launch', async () => {
    const controller = new AbortController();
    controller.abort();
    const run = await launch(controller.signal);
    await expect(run.mounted).rejects.toThrow();
    expect(run.frame).toBe(null);
  });

  it.each(['error', 'missing API', 'timeout', 'constructor failure'])('cleans up after %s', async (failure) => {
    vi.useFakeTimers();
    const run = await launch();
    const rejected = expect(run.mounted).rejects.toThrow();
    if (failure === 'timeout') await vi.advanceTimersByTimeAsync(30000);
    else {
      if (failure === 'constructor failure')
        fakeApi(
          run.frame!,
          vi.fn(() => {
            throw new Error('SDK failed');
          }),
        );
      run.script!.dispatchEvent(new Event(failure === 'error' ? 'error' : 'load'));
    }
    await rejected;
    expect(run.frame!.isConnected).toBe(false);
    expect(vi.getTimerCount()).toBe(0);
  });

  it.each(['https://docs.test/evil.js', 'http://docs.test/web-apps/apps/api/documents/api.js', 'javascript:alert(1)'])(
    'rejects unapproved API paths/schemes before mounting %s',
    async (url) => {
      const run = await launch(undefined, url);
      await expect(run.mounted).rejects.toThrow();
      expect(run.frame).toBe(null);
    },
  );
});
