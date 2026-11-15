import { afterEach, describe, expect, it, vi } from 'vitest';
import {
  createDocsCommunityReadClient,
  serveDocsCommunityReads,
  DOCS_COMMUNITY_PROTOCOL,
  type CommunityPluginApi,
} from '../../lib/docs-community-plugin';

class Port {
  onmessage: ((event: { data: unknown }) => void) | null = null;
  other?: Port;
  closed = false;
  sent: unknown[] = [];
  postMessage(data: unknown) {
    if (this.closed) throw new Error('Closed');
    this.sent.push(data);
    queueMicrotask(() => {
      if (!this.other?.closed) this.other?.onmessage?.({ data });
    });
  }
  start() {}
  close() {
    this.closed = true;
  }
}
const asPort = (port: Port) => port as unknown as MessagePort;
function channel() {
  const a = new Port();
  const b = new Port();
  a.other = b;
  b.other = a;
  return { a, b };
}
const request = (extra = {}) => ({
  protocol: DOCS_COMMUNITY_PROTOCOL,
  sessionId: 'session-a',
  requestId: 'request-a',
  tool: 'get_presentation_text',
  maxChars: 8000,
  ...extra,
});
const valid = { page: 1, totalSlides: 2, text: 'Hello', truncated: false, slidesRead: 2 };

function setup(delay = false) {
  const { a, b } = channel();
  const controller = new AbortController();
  const shape = (text: string) => ({ GetDocContent: () => ({ GetText: () => text }) });
  const slides = [[shape('Hello'), shape('World')], [shape('Next')]];
  vi.stubGlobal('Asc', { scope: {} });
  vi.stubGlobal('Api', {
    GetPresentation: () => ({
      GetCurSlideIndex: () => 0,
      GetSlidesCount: () => slides.length,
      GetSlideByIndex: (index: number) => ({ GetAllShapes: () => slides[index] }),
    }),
  });
  let completion: (() => void) | undefined;
  const plugin: CommunityPluginApi = {
    info: { editorType: 'slide' },
    executeMethod: vi.fn(),
    callCommand: vi.fn((code, close, recalculate, callback) => {
      expect(close).toBe(false);
      expect(recalculate).toBe(false);
      completion = () => callback(code());
      if (!delay) completion();
    }),
  };
  serveDocsCommunityReads(asPort(b), plugin, 'session-a', controller.signal);
  return { a, b, controller, plugin, slides, complete: () => completion?.() };
}
afterEach(() => {
  vi.unstubAllGlobals();
  vi.useRealTimers();
});

describe('finite Community plugin native reads', () => {
  it('reads actual public presentation methods without a connector or mutation', async () => {
    const run = setup();
    const client = createDocsCommunityReadClient(asPort(run.a), 'session-a', run.controller.signal);
    expect(await client.readPresentation()).toEqual({
      page: 1,
      totalSlides: 2,
      text: 'Hello\nWorld\nNext',
      slidesRead: 2,
      truncated: false,
    });
    expect(run.plugin.callCommand).toHaveBeenCalledTimes(1);
    expect(run.plugin.executeMethod).not.toHaveBeenCalled();
    client.close();
    run.controller.abort();
  });

  it('bounds text and refreshes native page data on each request', async () => {
    const run = setup();
    const client = createDocsCommunityReadClient(asPort(run.a), 'session-a', run.controller.signal);
    expect(await client.readPresentation(6)).toEqual({
      page: 1,
      totalSlides: 2,
      text: 'Hello\n',
      slidesRead: 1,
      truncated: true,
    });
    run.slides.splice(1);
    expect((await client.readPresentation()).totalSlides).toBe(1);
    client.close();
    run.controller.abort();
  });

  it('rejects a skipped native read without closing or automatically replaying it', async () => {
    const run = setup();
    vi.mocked(run.plugin.callCommand).mockImplementationOnce((_fn, _close, _recalc, callback) => callback(undefined));
    const client = createDocsCommunityReadClient(asPort(run.a), 'session-a', run.controller.signal);
    await expect(client.readPresentation()).rejects.toThrow('DOCS_PLUGIN_UNAVAILABLE');
    expect(run.plugin.callCommand).toHaveBeenCalledTimes(1);
    expect(run.a.closed).toBe(false);
    expect((await client.readPresentation()).text).toBe('Hello\nWorld\nNext');
    expect(run.plugin.callCommand).toHaveBeenCalledTimes(2);
    expect((run.a.sent[0] as { requestId: string }).requestId).not.toBe(
      (run.a.sent[1] as { requestId: string }).requestId,
    );
    run.controller.abort();
  });

  it('returns document script text as inert data and never executes it', async () => {
    const run = setup();
    const text = '"; Api.RemoveSlide(0); throw new Error("injected"); //';
    run.slides[0][0] = { GetDocContent: () => ({ GetText: () => text }) };
    const client = createDocsCommunityReadClient(asPort(run.a), 'session-a', run.controller.signal);
    expect((await client.readPresentation()).text).toContain(text);
    expect(run.slides).toHaveLength(2);
    client.close();
    run.controller.abort();
  });

  it.each([
    { tool: 'add_slide_text' },
    { tool: 'add_slide_text', text: 'Injected' },
    { tool: 'execute_js', code: 'throw Error()' },
    { script: 'Api.RemoveSlide(0)' },
    { maxChars: 0 },
    { maxChars: 8001 },
    { maxChars: 1.5 },
    { maxChars: '8000' },
  ])('refuses unsupported or malformed requests: %j', async (extra) => {
    const run = setup();
    run.a.postMessage(request(extra));
    await Promise.resolve();
    expect(run.plugin.callCommand).not.toHaveBeenCalled();
    expect(run.b.sent).toEqual([expect.objectContaining({ error: 'DOCS_PLUGIN_UNSUPPORTED_REQUEST' })]);
    run.controller.abort();
  });

  it.each([null, [], { protocol: 'wrong' }, request({ sessionId: 'other' }), request({ requestId: '../bad' })])(
    'ignores another session or malformed envelope %j',
    async (value) => {
      const run = setup();
      run.a.postMessage(value);
      await Promise.resolve();
      expect(run.plugin.callCommand).not.toHaveBeenCalled();
      expect(run.b.sent).toHaveLength(0);
      run.controller.abort();
    },
  );

  it('rejects duplicate IDs and busy requests without another native call', async () => {
    const run = setup(true);
    run.a.postMessage(request());
    await Promise.resolve();
    run.a.postMessage(request({ requestId: 'second' }));
    await Promise.resolve();
    expect(run.b.sent).toEqual([expect.objectContaining({ error: 'DOCS_PLUGIN_BUSY' })]);
    run.complete();
    run.a.postMessage(request());
    await Promise.resolve();
    expect(run.b.sent.at(-1)).toEqual(expect.objectContaining({ error: 'DOCS_PLUGIN_REQUEST_USED' }));
    expect(run.plugin.callCommand).toHaveBeenCalledTimes(1);
    run.controller.abort();
  });

  it('suppresses a native result arriving after disposal', async () => {
    const run = setup(true);
    run.a.postMessage(request());
    await Promise.resolve();
    run.controller.abort();
    run.complete();
    expect(run.b.sent).toHaveLength(0);
    expect(run.b.closed).toBe(true);
  });

  it('allows only one port per plugin scope', () => {
    const run = setup();
    expect(() =>
      serveDocsCommunityReads(asPort(channel().b), run.plugin, 'session-b', new AbortController().signal),
    ).toThrow('already attached');
    run.controller.abort();
    expect(() =>
      serveDocsCommunityReads(asPort(channel().b), run.plugin, 'session-b', new AbortController().signal),
    ).not.toThrow();
  });

  it.each(['invalid result', 'native exception', 'scope missing', 'empty deck', 'oversized deck'])(
    'fails closed on %s',
    async (failure) => {
      const run = setup();
      if (failure === 'invalid result')
        vi.mocked(run.plugin.callCommand).mockImplementation((_fn, _close, _recalc, cb) => cb({ ...valid, page: 0 }));
      if (failure === 'native exception')
        vi.mocked(run.plugin.callCommand).mockImplementation(() => {
          throw new Error('secret');
        });
      if (failure === 'scope missing') vi.stubGlobal('Asc', undefined);
      if (failure === 'empty deck') run.slides.length = 0;
      if (failure === 'oversized deck') run.slides.length = 1001;
      const client = createDocsCommunityReadClient(asPort(run.a), 'session-a', run.controller.signal);
      await expect(client.readPresentation()).rejects.toThrow('read failed');
      expect(run.a.closed).toBe(true);
      run.controller.abort();
    },
  );
});

describe('owned read channel lifecycle', () => {
  it.each([0, 8001, 1.5, NaN])('rejects invalid read limits %s before sending', async (max) => {
    const { a } = channel();
    const client = createDocsCommunityReadClient(asPort(a), 's', new AbortController().signal);
    await expect(client.readPresentation(max)).rejects.toThrow('Invalid Docs read limit');
    expect(a.sent).toHaveLength(0);
    client.close();
  });

  it('ignores uncorrelated messages, rejects concurrent calls, then resolves the matching result', async () => {
    const { a, b } = channel();
    const client = createDocsCommunityReadClient(asPort(a), 'session-a', new AbortController().signal);
    const pending = client.readPresentation();
    const sent = a.sent[0] as ReturnType<typeof request>;
    await expect(client.readPresentation()).rejects.toThrow('in progress');
    for (const value of [
      null,
      {},
      { ...sent, sessionId: 'other', result: valid },
      { ...sent, requestId: 'other', result: valid },
      { ...sent, protocol: 'other', result: valid },
    ])
      b.postMessage(value);
    await Promise.resolve();
    expect(a.closed).toBe(false);
    b.postMessage({ ...sent, result: valid });
    expect(await pending).toEqual(valid);
    client.close();
  });

  it.each(['timeout', 'abort', 'post failure'])('closes after %s and never retries', async (failure) => {
    vi.useFakeTimers();
    const { a } = channel();
    const controller = new AbortController();
    const client = createDocsCommunityReadClient(asPort(a), 'session-a', controller.signal);
    if (failure === 'post failure') a.closed = true;
    const rejected = expect(client.readPresentation()).rejects.toThrow('closed');
    if (failure === 'timeout') await vi.advanceTimersByTimeAsync(30000);
    if (failure === 'abort') controller.abort();
    await rejected;
    await expect(client.readPresentation()).rejects.toThrow('closed');
    expect(a.sent.length).toBeLessThanOrEqual(1);
    expect(vi.getTimerCount()).toBe(0);
    client.close();
  });

  it.each([
    { ...valid, text: 'x'.repeat(8001) },
    { ...valid, totalSlides: 0 },
    { ...valid, slidesRead: 3 },
    { ...valid, truncated: 'no' },
    null,
  ])('rejects bad native responses %j', async (result) => {
    const { a, b } = channel();
    const client = createDocsCommunityReadClient(asPort(a), 's', new AbortController().signal);
    const rejected = expect(client.readPresentation()).rejects.toThrow('read failed');
    b.postMessage({ ...(a.sent[0] as object), result });
    await rejected;
    expect(a.closed).toBe(true);
  });

  it('closes on an unavailable error that also carries a result', async () => {
    const { a, b } = channel();
    const client = createDocsCommunityReadClient(asPort(a), 's', new AbortController().signal);
    const rejected = expect(client.readPresentation()).rejects.toThrow('read failed');
    b.postMessage({ ...(a.sent[0] as object), error: 'DOCS_PLUGIN_UNAVAILABLE', result: valid });
    await rejected;
    expect(a.closed).toBe(true);
  });

  it('enforces the requested output limit on the receiving side', async () => {
    const { a, b } = channel();
    const client = createDocsCommunityReadClient(asPort(a), 's', new AbortController().signal);
    const rejected = expect(client.readPresentation(1)).rejects.toThrow('read failed');
    b.postMessage({ ...(a.sent[0] as object), result: valid });
    await rejected;
  });

  it('rejects invalid session identity and pre-aborted channels', () => {
    const { a } = channel();
    const controller = new AbortController();
    controller.abort();
    expect(() => createDocsCommunityReadClient(asPort(a), '../bad', new AbortController().signal)).toThrow('Invalid');
    expect(() => createDocsCommunityReadClient(asPort(a), 's', controller.signal)).toThrow();
    const plugin: CommunityPluginApi = { info: { editorType: 'word' }, executeMethod: vi.fn(), callCommand: vi.fn() };
    expect(() => serveDocsCommunityReads(asPort(a), plugin, 's', new AbortController().signal)).toThrow('Invalid');
    expect(() =>
      serveDocsCommunityReads(asPort(a), { ...plugin, info: { editorType: 'slide' } }, 's', controller.signal),
    ).toThrow();
  });
});

const barrierId = 'd5820881-b533-4d0b-aafe-d0129c660fb8';
const held = { barrierId, state: 'held', serverIndex: 12 };
describe('fixed native barrier calls over the same owned port', () => {
  it('detects the extension using only fixed status and keeps old-engine reads available', async () => {
    const run = setup();
    vi.mocked(run.plugin.executeMethod).mockImplementation((_name, _args, callback) => callback(undefined));
    const client = createDocsCommunityReadClient(asPort(run.a), 'session-a', run.controller.signal);
    expect(await client.barrierCapability()).toEqual({ available: false });
    expect(run.plugin.executeMethod).toHaveBeenCalledWith(
      'DeckwiseBarrierStatus',
      [expect.stringMatching(/^[a-f0-9-]{36}$/)],
      expect.any(Function),
    );
    expect(run.plugin.callCommand).not.toHaveBeenCalled();
    await expect(client.quiesceAndDrain(barrierId)).rejects.toThrow('DOCS_BARRIER_UNAVAILABLE');
    expect(run.a.closed).toBe(false);
    expect((await client.readPresentation()).text).toBe('Hello\nWorld\nNext');
    run.controller.abort();
  });

  it('uses exactly the three fixed native methods and forwards finite receipts', async () => {
    const run = setup();
    vi.mocked(run.plugin.executeMethod).mockImplementation((name, args, callback) =>
      callback(
        args[0] !== barrierId
          ? { error: 'NATIVE_BARRIER_NOT_FOUND' }
          : name === 'DeckwiseQuiesceAndDrain'
            ? { barrierId, state: 'draining' }
            : name === 'DeckwiseResume'
              ? { ...held, state: 'resumed' }
              : held,
      ),
    );
    const client = createDocsCommunityReadClient(asPort(run.a), 'session-a', run.controller.signal);
    expect(await client.barrierCapability()).toEqual({ available: true });
    expect(await client.quiesceAndDrain(barrierId)).toEqual({ barrierId, state: 'draining' });
    expect(await client.barrierStatus(barrierId)).toEqual(held);
    expect(await client.resume(barrierId)).toEqual({ ...held, state: 'resumed' });
    expect(vi.mocked(run.plugin.executeMethod).mock.calls.map((call) => call[0])).toEqual([
      'DeckwiseBarrierStatus',
      'DeckwiseQuiesceAndDrain',
      'DeckwiseBarrierStatus',
      'DeckwiseResume',
    ]);
    expect(run.plugin.callCommand).not.toHaveBeenCalled();
    run.controller.abort();
  });

  it('returns legal native refusal without closing the read channel or fabricating a receipt', async () => {
    const run = setup();
    vi.mocked(run.plugin.executeMethod).mockImplementation((_name, _args, callback) =>
      callback({ error: 'NATIVE_BARRIER_ACK_UNKNOWN' }),
    );
    const client = createDocsCommunityReadClient(asPort(run.a), 'session-a', run.controller.signal);
    await expect(client.resume(barrierId)).rejects.toThrow('NATIVE_BARRIER_ACK_UNKNOWN');
    expect(run.a.closed).toBe(false);
    expect((await client.readPresentation()).text).toContain('Hello');
    run.controller.abort();
  });

  it('preserves uncertain-held with its native error and old index', async () => {
    const run = setup();
    const result = { ...held, state: 'uncertain-held', error: 'NATIVE_BARRIER_STATE_CHANGED' };
    vi.mocked(run.plugin.executeMethod).mockImplementation((_name, _args, callback) => callback(result));
    const client = createDocsCommunityReadClient(asPort(run.a), 'session-a', run.controller.signal);
    expect(await client.barrierStatus(barrierId)).toEqual(result);
    run.controller.abort();
  });

  it.each([
    { tool: 'native_barrier_quiesce', barrierId: '../bad' },
    { tool: 'native_barrier_quiesce', barrierId: barrierId.toUpperCase() },
    { tool: 'native_barrier_resume', barrierId, method: 'RemoveSlide' },
    { tool: 'native_barrier_status', barrierId, script: 'Api.RemoveSlide(0)' },
    { tool: 'native_barrier_capability', barrierId },
    { tool: { toString: 'bad' } },
  ])('rejects malformed barrier requests before native dispatch: %j', async (extra) => {
    const run = setup();
    const { maxChars: _, ...base } = request();
    run.a.postMessage({ ...base, ...extra });
    await Promise.resolve();
    expect(run.plugin.executeMethod).not.toHaveBeenCalled();
    expect(run.plugin.callCommand).not.toHaveBeenCalled();
    expect(run.b.sent).toHaveLength(1);
    run.controller.abort();
  });

  it.each([
    null,
    { ...held, barrierId: 'other' },
    { ...held, serverIndex: -1 },
    { ...held, serverIndex: 1.2 },
    { ...held, state: 'saved' },
    { barrierId, state: 'held' },
    { ...held, token: 'secret' },
    { barrierId, state: 'uncertain-held', error: 'secret-error' },
    { error: 'secret-error' },
    { error: 'NATIVE_BARRIER_BUSY', config: 'secret' },
    { barrierId, state: 'draining', serverIndex: 1 },
  ])('rejects malformed native receipts without forwarding native detail: %j', async (value) => {
    const run = setup();
    vi.mocked(run.plugin.executeMethod).mockImplementation((_name, _args, callback) => callback(value));
    const client = createDocsCommunityReadClient(asPort(run.a), 'session-a', run.controller.signal);
    await expect(client.barrierStatus(barrierId)).rejects.toThrow('DOCS_BARRIER_INVALID_RESPONSE');
    expect(JSON.stringify(run.b.sent)).not.toContain('secret');
    expect(run.a.closed).toBe(true);
    run.controller.abort();
  });

  it('validates UUID before sending and ignores a late native result after abort', async () => {
    const run = setup();
    let finish: (value: unknown) => void = () => {};
    vi.mocked(run.plugin.executeMethod).mockImplementation((_name, _args, callback) => {
      finish = callback;
    });
    const client = createDocsCommunityReadClient(asPort(run.a), 'session-a', run.controller.signal);
    await expect(client.quiesceAndDrain('invalid')).rejects.toThrow('DOCS_BARRIER_INVALID_REQUEST');
    expect(run.a.sent).toHaveLength(0);
    const rejection = expect(client.quiesceAndDrain(barrierId)).rejects.toThrow('closed');
    await Promise.resolve();
    run.controller.abort();
    await rejection;
    finish(held);
    expect(run.b.sent).toHaveLength(0);
  });

  it('times out only the barrier request, permits explicit status recovery and never auto-replays', async () => {
    vi.useFakeTimers();
    const run = setup();
    let finish: (value: unknown) => void = () => {};
    vi.mocked(run.plugin.executeMethod)
      .mockImplementationOnce((_name, _args, callback) => {
        finish = callback;
      })
      .mockImplementation((_name, _args, callback) => callback(held));
    const client = createDocsCommunityReadClient(asPort(run.a), 'session-a', run.controller.signal);
    const rejection = expect(client.quiesceAndDrain(barrierId)).rejects.toThrow('DOCS_BARRIER_ACK_UNKNOWN');
    await Promise.resolve();
    await vi.advanceTimersByTimeAsync(30000);
    await rejection;
    expect(run.a.closed).toBe(false);
    expect(run.plugin.executeMethod).toHaveBeenCalledTimes(1);
    await expect(client.readPresentation()).rejects.toThrow('DOCS_PLUGIN_BUSY');
    expect(run.a.closed).toBe(false);
    await expect(client.barrierStatus(barrierId)).rejects.toThrow('DOCS_PLUGIN_BUSY');
    expect(run.a.closed).toBe(false);
    finish({ barrierId, state: 'draining' });
    await Promise.resolve();
    expect(await client.barrierStatus(barrierId)).toEqual(held);
    expect(run.plugin.executeMethod).toHaveBeenCalledTimes(2);
    run.controller.abort();
  });
});
