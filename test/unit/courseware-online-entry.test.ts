import { afterEach, describe, expect, it, vi } from 'vitest';
import { COURSEWARE_ONLINE_PROTOCOL, installCoursewareOnlineEntry } from '../../lib/courseware-online-entry';
import type { mountDocsOnlineEditor } from '../../lib/docs-online-editor';

const disposers: (() => void)[] = [];
const flush = async () => {
  await Promise.resolve();
  await Promise.resolve();
  await Promise.resolve();
};
afterEach(() => {
  for (const dispose of disposers.splice(0)) dispose();
  document.body.replaceChildren();
});
function setup() {
  const parent = { postMessage: vi.fn() } as unknown as Window;
  const parentOrigin = 'deckwise-editor://host';
  const container = document.createElement('div');
  document.body.append(container);
  const mounts: {
    args: Parameters<typeof mountDocsOnlineEditor>;
    resolve: (value: Awaited<ReturnType<typeof mountDocsOnlineEditor>>) => void;
  }[] = [];
  const mount: typeof mountDocsOnlineEditor = (...args) =>
    new Promise((resolve) => {
      mounts.push({ args, resolve });
    });
  const dispose = installCoursewareOnlineEntry(
    container,
    {
      parentWindow: parent,
      parentOrigin,
      channelId: 'owned-channel',
      apiUrl: 'https://docs.test/web-apps/apps/api/documents/api.js',
    },
    mount,
  );
  disposers.push(dispose);
  const open = (generation = 1) => ({
    protocol: COURSEWARE_ONLINE_PROTOCOL,
    channelId: 'owned-channel',
    requestId: `open-${generation}`,
    type: 'open',
    generation,
    projectId: 'e53fe140-6154-4aa2-8d04-25348952b700',
    sessionEpoch: generation,
    documentKey: `fixed-server-key-${generation}`,
    docsUserId: 'tenant-1-user-2',
    configJson: 'signed-original-config',
    expiresAtEpochMillis: Date.now() + 60000,
  });
  const send = (data: unknown, origin = parentOrigin, source = parent) =>
    window.dispatchEvent(new MessageEvent('message', { data, origin, source }));
  const replies = () => (parent.postMessage as ReturnType<typeof vi.fn>).mock.calls.map((call) => call[0]);
  return { parent, parentOrigin, container, mounts, dispose, open, send, replies };
}

describe('owned online entry handshake, not a Docs engine substitute', () => {
  it('advertises to the exact pinned parent and forwards untouched signed config and fixed identities', async () => {
    const f = setup();
    expect(f.replies()).toEqual([
      { protocol: COURSEWARE_ONLINE_PROTOCOL, channelId: 'owned-channel', type: 'entry-ready' },
    ]);
    expect(f.parent.postMessage).toHaveBeenCalledWith(expect.anything(), f.parentOrigin);
    f.send(f.open());
    await flush();
    expect(f.mounts).toHaveLength(1);
    const [container, hostUrl, apiUrl, config, binding, signal, events] = f.mounts[0].args;
    expect(container).toBe(f.container);
    expect(hostUrl).toBe(`${location.origin}/docs-online-host.html`);
    expect(apiUrl).toBe('https://docs.test/web-apps/apps/api/documents/api.js');
    expect(config).toBe('signed-original-config');
    expect(binding).toEqual({ documentKey: 'fixed-server-key-1', userId: 'tenant-1-user-2' });
    expect(signal.aborted).toBe(false);
    events?.onDocumentReady?.();
    events?.onDocumentStateChange?.({ data: true });
    expect(f.replies().at(-1)).toMatchObject({
      type: 'state',
      dirty: true,
      generation: 1,
      sessionEpoch: 1,
      documentKey: binding.documentKey,
    });
    f.mounts[0].resolve({ destroy: vi.fn() });
    await flush();
    expect(f.replies().at(-1)).toMatchObject({ type: 'mounted', canReadText: false });
    expect(f.replies().some((value) => value.type === 'saved')).toBe(false);
  });

  it('ignores foreign origins, foreign windows and wrong channels before consuming the request', async () => {
    const f = setup();
    f.send(f.open(), 'https://foreign.test');
    f.send(f.open(), f.parentOrigin, {} as Window);
    f.send({ ...f.open(), channelId: 'foreign-channel' });
    await flush();
    expect(f.mounts).toHaveLength(0);
    expect(f.replies()).toHaveLength(1);
    f.send(f.open());
    await flush();
    expect(f.mounts).toHaveLength(1);
  });

  it('rejects request replay, stale generations, injected commands and expired launch data', async () => {
    const f = setup();
    f.send(f.open());
    await flush();
    f.send(f.open());
    expect(f.replies().at(-1)).toMatchObject({ code: 'ONLINE_REQUEST_REUSED' });
    f.send({ ...f.open(), requestId: 'old-generation' });
    expect(f.replies().at(-1)).toMatchObject({ code: 'ONLINE_INVALID_REQUEST' });
    f.send({ ...f.open(2), script: 'arbitrary-code' });
    expect(f.replies().at(-1)).toMatchObject({ code: 'ONLINE_INVALID_REQUEST' });
    f.send({ ...f.open(3), expiresAtEpochMillis: Date.now() - 1 });
    expect(f.replies().at(-1)).toMatchObject({ code: 'ONLINE_INVALID_REQUEST' });
    expect(f.mounts).toHaveLength(1);
  });

  it('aborts replaced mounts, destroys late results and suppresses stale engine events', async () => {
    const f = setup();
    f.send(f.open());
    await flush();
    const old = f.mounts[0];
    f.send(f.open(2));
    await flush();
    expect(old.args[5].aborted).toBe(true);
    expect(f.mounts).toHaveLength(2);
    const destroy = vi.fn();
    old.resolve({ destroy });
    await flush();
    expect(destroy).toHaveBeenCalledOnce();
    const count = f.replies().length;
    old.args[6]?.onDocumentReady?.();
    expect(f.replies()).toHaveLength(count);
    f.send({
      protocol: COURSEWARE_ONLINE_PROTOCOL,
      channelId: 'owned-channel',
      requestId: 'stale-close',
      type: 'close',
      generation: 1,
    });
    expect(f.replies().at(-1)).toMatchObject({ code: 'ONLINE_GENERATION_STALE' });
    expect(f.mounts[1].args[5].aborted).toBe(false);
  });

  it('closes only the owned view, and disposes without ending the shared server session', async () => {
    const f = setup();
    f.send(f.open());
    await flush();
    const destroy = vi.fn();
    f.mounts[0].resolve({ destroy });
    await flush();
    f.send({
      protocol: COURSEWARE_ONLINE_PROTOCOL,
      channelId: 'owned-channel',
      requestId: 'close-1',
      type: 'close',
      generation: 1,
    });
    expect(destroy).toHaveBeenCalledOnce();
    expect(f.replies().at(-1)).toMatchObject({ type: 'closed', generation: 1 });
    f.dispose();
    const count = f.replies().length;
    f.send(f.open(2));
    await flush();
    expect(f.replies()).toHaveLength(count);
  });

  it('permits only bounded Community text reads and invalidates their late replies on epoch change', async () => {
    const f = setup();
    f.send(f.open());
    await flush();
    let finish: (value: unknown) => void = () => {};
    const read = vi.fn(
      () =>
        new Promise((resolve) => {
          finish = resolve;
        }),
    );
    f.mounts[0].resolve({
      destroy: vi.fn(),
      communityReads: { ...barrierBridge(), close: vi.fn(), readPresentation: read },
    } as Awaited<ReturnType<typeof mountDocsOnlineEditor>>);
    await flush();
    const request = {
      protocol: COURSEWARE_ONLINE_PROTOCOL,
      channelId: 'owned-channel',
      requestId: 'read-1',
      type: 'read',
      generation: 1,
      maxChars: 8000,
    };
    f.send({ ...request, maxChars: 8001 });
    expect(read).not.toHaveBeenCalled();
    f.send({ ...request, requestId: 'read-2' });
    expect(read).toHaveBeenCalledWith(8000);
    f.send(f.open(2));
    await flush();
    const count = f.replies().length;
    finish({ text: 'late old epoch' });
    await flush();
    expect(f.replies()).toHaveLength(count);
  });
});

const barrierId = 'd5820881-b533-4d0b-aafe-d0129c660fb8';
const held = { barrierId, state: 'held' as const, serverIndex: 12 };
function barrierBridge() {
  return {
    close: vi.fn(),
    readPresentation: vi.fn(async () => ({
      page: 1,
      totalSlides: 1,
      text: 'original text',
      truncated: false,
      slidesRead: 1,
    })),
    barrierCapability: vi.fn(async () => ({ available: true })),
    quiesceAndDrain: vi.fn(async () => ({ barrierId, state: 'draining' as const })),
    barrierStatus: vi.fn(async () => held),
    resume: vi.fn(async () => ({ ...held, state: 'resumed' as const })),
  };
}
const barrierRequest = (type: string, extra = {}) => ({
  protocol: COURSEWARE_ONLINE_PROTOCOL,
  channelId: 'owned-channel',
  requestId: type,
  type,
  generation: 1,
  ...(type === 'barrier-capability' ? {} : { barrierId }),
  ...extra,
});

describe('registered Docs window user identity', () => {
  const clientId = 'd5820881-b533-4d0b-aafe-d0129c660fb8';
  it.each(['tenant-1-user-2', `tenant-1-user-2-client-${clientId}`])(
    'preserves the exact authorized identity %s',
    async (docsUserId) => {
      const f = setup();
      f.send({ ...f.open(), docsUserId });
      await flush();
      expect(f.mounts).toHaveLength(1);
      expect(f.mounts[0].args[4].userId).toBe(docsUserId);
      f.mounts[0].resolve({ destroy: vi.fn(), communityReads: barrierBridge() });
      await flush();
      expect(f.replies().at(-1)).toMatchObject({ type: 'mounted', docsUserId });
    },
  );
  it.each([
    `tenant-1-user-2-client-${clientId.toUpperCase()}`,
    'tenant-1-user-2-client-',
    `tenant-0-user-2-client-${clientId}`,
    `tenant-1-user-02-client-${clientId}`,
    `tenant-1-user-2-client-${clientId}-extra`,
    `tenant-1-user-2-client-${clientId}\n`,
    `tenant-1-user-2-client-${clientId.replaceAll('-', '')}`,
    'local-user',
    `tenant-${'1'.repeat(100)}-user-2-client-${clientId}`,
  ])('rejects malformed or oversized identity before mounting: %j', async (docsUserId) => {
    const f = setup();
    f.send({ ...f.open(), docsUserId });
    await flush();
    expect(f.mounts).toHaveLength(0);
    expect(f.replies().at(-1)).toMatchObject({ code: 'ONLINE_INVALID_REQUEST' });
  });
});

describe('owned fixed native barrier entry, without a publication UI', () => {
  it('forwards only fixed operations with the current mount identities; dirty=false is not an ACK', async () => {
    const f = setup();
    f.send(f.open());
    await flush();
    const bridge = barrierBridge(),
      destroy = vi.fn();
    f.mounts[0].resolve({ destroy, communityReads: bridge });
    await flush();
    f.mounts[0].args[6]?.onDocumentStateChange?.({ data: false });
    expect(f.replies().some((value) => value.type === 'barrier')).toBe(false);
    for (const [type, method, result] of [
      ['barrier-capability', bridge.barrierCapability, { available: true }],
      ['barrier-quiesce', bridge.quiesceAndDrain, { barrierId, state: 'draining' }],
      ['barrier-status', bridge.barrierStatus, held],
      ['barrier-resume', bridge.resume, { ...held, state: 'resumed' }],
    ] as const) {
      f.send(barrierRequest(type));
      await flush();
      await flush();
      expect(method).toHaveBeenCalledTimes(1);
      expect(f.replies().at(-1)).toMatchObject({
        requestId: type,
        type: type === 'barrier-capability' ? type : 'barrier',
        generation: 1,
        projectId: f.open().projectId,
        sessionEpoch: 1,
        documentKey: 'fixed-server-key-1',
        docsUserId: 'tenant-1-user-2',
        result,
      });
    }
    expect(bridge.quiesceAndDrain).toHaveBeenCalledWith(barrierId);
    expect(bridge.readPresentation).not.toHaveBeenCalled();
    expect(destroy).not.toHaveBeenCalled();
    expect(bridge.close).not.toHaveBeenCalled();
  });

  it('reports old-engine unavailability and retains the existing read connection', async () => {
    const f = setup();
    f.send(f.open());
    await flush();
    const bridge = barrierBridge(),
      destroy = vi.fn();
    bridge.barrierCapability.mockResolvedValue({ available: false });
    bridge.quiesceAndDrain.mockRejectedValue(new Error('DOCS_BARRIER_UNAVAILABLE'));
    f.mounts[0].resolve({ destroy, communityReads: bridge });
    await flush();
    f.send(barrierRequest('barrier-capability'));
    await flush();
    await flush();
    expect(f.replies().at(-1)).toMatchObject({ type: 'barrier-capability', result: { available: false } });
    f.send(barrierRequest('barrier-quiesce'));
    await flush();
    await flush();
    expect(f.replies().at(-1)).toMatchObject({ type: 'error', code: 'ONLINE_NATIVE_BARRIER_UNAVAILABLE' });
    f.send({
      protocol: COURSEWARE_ONLINE_PROTOCOL,
      channelId: 'owned-channel',
      requestId: 'still-read',
      type: 'read',
      generation: 1,
      maxChars: 8000,
    });
    await flush();
    expect(f.replies().at(-1)).toMatchObject({ type: 'text', result: { text: 'original text' } });
    expect(destroy).not.toHaveBeenCalled();
    expect(bridge.close).not.toHaveBeenCalled();
  });

  it.each([
    { barrierId: 'invalid' },
    { barrierId: barrierId.toUpperCase() },
    { script: 'arbitrary code' },
    { method: 'RemoveSlide' },
  ])('refuses malformed barrier payload before dispatch: %j', async (extra) => {
    const f = setup();
    f.send(f.open());
    await flush();
    const bridge = barrierBridge();
    f.mounts[0].resolve({ destroy: vi.fn(), communityReads: bridge });
    await flush();
    f.send(barrierRequest('barrier-quiesce', extra));
    await flush();
    expect(f.replies().at(-1)).toMatchObject({ code: 'ONLINE_INVALID_REQUEST' });
    expect(bridge.quiesceAndDrain).not.toHaveBeenCalled();
  });

  it('preserves origin/window/channel and request identity checks for barriers', async () => {
    const f = setup();
    f.send(f.open());
    await flush();
    const bridge = barrierBridge();
    f.mounts[0].resolve({ destroy: vi.fn(), communityReads: bridge });
    await flush();
    f.send(barrierRequest('barrier-quiesce'), 'https://foreign.test');
    f.send(barrierRequest('barrier-quiesce'), f.parentOrigin, {} as Window);
    f.send(barrierRequest('barrier-quiesce', { channelId: 'other' }));
    await flush();
    expect(bridge.quiesceAndDrain).not.toHaveBeenCalled();
    f.send(barrierRequest('barrier-quiesce'));
    await flush();
    await flush();
    f.send(barrierRequest('barrier-quiesce'));
    await flush();
    expect(f.replies().at(-1)).toMatchObject({ code: 'ONLINE_REQUEST_REUSED' });
    expect(bridge.quiesceAndDrain).toHaveBeenCalledTimes(1);
  });

  it('suppresses a late receipt on generation change and refuses stale requests', async () => {
    const f = setup();
    f.send(f.open());
    await flush();
    const bridge = barrierBridge();
    let finish: (value: typeof held) => void = () => {};
    bridge.barrierStatus.mockImplementation(
      () =>
        new Promise((resolve) => {
          finish = resolve;
        }),
    );
    f.mounts[0].resolve({ destroy: vi.fn(), communityReads: bridge });
    await flush();
    f.send(barrierRequest('barrier-status'));
    await flush();
    f.send(f.open(2));
    await flush();
    const count = f.replies().length;
    finish(held);
    await flush();
    expect(f.replies()).toHaveLength(count);
    f.send(barrierRequest('barrier-status', { requestId: 'stale' }));
    expect(f.replies().at(-1)).toMatchObject({ code: 'ONLINE_GENERATION_STALE' });
  });

  it.each([
    ['NATIVE_BARRIER_ACK_UNKNOWN', 'NATIVE_BARRIER_ACK_UNKNOWN'],
    ['DOCS_BARRIER_ACK_UNKNOWN', 'ONLINE_NATIVE_BARRIER_ACK_UNKNOWN'],
    ['DOCS_PLUGIN_BUSY', 'ONLINE_NATIVE_BARRIER_BUSY'],
    ['secret token and URL', 'ONLINE_NATIVE_BARRIER_FAILED'],
  ])('fails only this request for %s and retains the mounted editor', async (failure, code) => {
    const f = setup();
    f.send(f.open());
    await flush();
    const bridge = barrierBridge(),
      destroy = vi.fn();
    bridge.barrierStatus.mockRejectedValue(new Error(failure));
    f.mounts[0].resolve({ destroy, communityReads: bridge });
    await flush();
    f.send(barrierRequest('barrier-status'));
    await flush();
    await flush();
    expect(f.replies().at(-1)).toMatchObject({ type: 'error', requestId: 'barrier-status', generation: 1, code });
    expect(JSON.stringify(f.replies())).not.toContain('secret token');
    expect(destroy).not.toHaveBeenCalled();
    expect(bridge.close).not.toHaveBeenCalled();
    expect(f.mounts[0].args[5].aborted).toBe(false);
  });

  it('does not turn a native uncertain-held response into a ready/resumed response', async () => {
    const f = setup();
    f.send(f.open());
    await flush();
    const bridge = barrierBridge();
    bridge.barrierStatus.mockResolvedValue({
      ...held,
      state: 'uncertain-held',
      error: 'NATIVE_BARRIER_STATE_CHANGED',
    } as unknown as typeof held);
    f.mounts[0].resolve({ destroy: vi.fn(), communityReads: bridge });
    await flush();
    f.send(barrierRequest('barrier-status'));
    await flush();
    await flush();
    expect(f.replies().at(-1)).toMatchObject({
      type: 'barrier',
      result: { state: 'uncertain-held', error: 'NATIVE_BARRIER_STATE_CHANGED' },
    });
  });
});
