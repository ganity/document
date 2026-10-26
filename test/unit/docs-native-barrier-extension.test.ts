import { afterEach, describe, expect, it, vi } from 'vitest';
import { installDocsNativeBarrier } from '../../lib/docs-native-barrier-extension';
import { installDocsNativeBarrier930138 } from '../../lib/docs-native-barrier-9.3.0-138';

function fixture() {
  class Editor {
    restriction = 4;
    connected = true;
    clean = false;
    idle = true;
    saveAccepted = true;
  }
  let ack: (index: number) => void = () => {};
  let fail: () => void = () => {};
  let blocked: () => boolean = () => false;
  const save = vi.fn((editor: Editor) => editor.saveAccepted);
  const roundTrip = vi.fn((_editor: Editor, failed: () => void) => {
    fail = failed;
  });
  installDocsNativeBarrier(
    Editor.prototype,
    {
      supported: () => true,
      connected: (editor) => editor.connected,
      restrictions: (editor) => editor.restriction,
      addView: (editor) => {
        editor.restriction |= 128;
      },
      removeView: (editor) => {
        editor.restriction &= ~128;
      },
      idle: (editor) => editor.idle,
      clean: (editor) => editor.clean,
      save,
      roundTrip,
      observeAck: (_editor, received) => {
        ack = received;
        return vi.fn();
      },
      blockPluginWrites: (_editor, callback) => {
        blocked = callback;
        return vi.fn();
      },
    },
    { timeoutMs: 100, retryMs: 10 },
  );
  const editor = new Editor() as Editor & {
    pluginMethod_DeckwiseQuiesceAndDrain(id: string): { state: string; serverIndex?: number; error?: string };
    pluginMethod_DeckwiseBarrierStatus(id: string): { state: string; serverIndex?: number; error?: string };
    pluginMethod_DeckwiseResume(id: string): { state: string };
  };
  return { editor, save, roundTrip, ack: (index = 5) => ack(index), fail: () => fail(), blocked: () => blocked() };
}
const id = '00000000-0000-0000-0000-000000000001';
afterEach(() => {
  vi.useRealTimers();
});

describe('fixed native quiesce and drain', () => {
  it('freezes synchronously but requires the actual ACK and no remaining changes', () => {
    vi.useFakeTimers();
    const f = fixture();
    expect(f.editor.pluginMethod_DeckwiseQuiesceAndDrain(id).state).toBe('draining');
    expect(f.editor.restriction).toBe(132);
    expect(f.blocked()).toBe(true);
    expect(f.save).toHaveBeenCalledOnce();
    f.ack();
    expect(f.editor.pluginMethod_DeckwiseBarrierStatus(id).state).toBe('draining');
    f.editor.clean = true;
    expect(f.editor.pluginMethod_DeckwiseBarrierStatus(id).state).toBe('draining');
    f.ack(7);
    expect(f.editor.pluginMethod_DeckwiseBarrierStatus(id)).toMatchObject({ state: 'held', serverIndex: 7 });
    expect(f.editor.pluginMethod_DeckwiseResume(id).state).toBe('resumed');
    expect(f.editor.restriction).toBe(4);
    expect(f.blocked()).toBe(false);
  });
  it('does not fabricate a clean-document ACK and performs a native empty round trip', () => {
    vi.useFakeTimers();
    const f = fixture();
    f.editor.clean = true;
    f.editor.saveAccepted = false;
    expect(f.editor.pluginMethod_DeckwiseQuiesceAndDrain(id).state).toBe('draining');
    expect(f.roundTrip).toHaveBeenCalledOnce();
    vi.advanceTimersByTime(60);
    expect(f.editor.pluginMethod_DeckwiseBarrierStatus(id).state).toBe('draining');
    f.ack();
    expect(f.editor.pluginMethod_DeckwiseBarrierStatus(id).state).toBe('held');
  });
  it('keeps the view held on unknown ACK and only reconciles from a real late ACK', () => {
    vi.useFakeTimers();
    const f = fixture();
    f.editor.pluginMethod_DeckwiseQuiesceAndDrain(id);
    vi.advanceTimersByTime(110);
    expect(f.editor.pluginMethod_DeckwiseBarrierStatus(id)).toMatchObject({
      state: 'uncertain-held',
      error: 'NATIVE_BARRIER_ACK_UNKNOWN',
    });
    expect(f.editor.pluginMethod_DeckwiseResume(id)).toMatchObject({ error: 'NATIVE_BARRIER_ACK_UNKNOWN' });
    expect(f.editor.restriction & 128).toBe(128);
    f.editor.clean = true;
    f.ack();
    expect(f.editor.pluginMethod_DeckwiseResume(id).state).toBe('resumed');
  });
  it('retries waiting for the current native save and does not accept its unrelated ACK', () => {
    vi.useFakeTimers();
    const f = fixture();
    f.editor.idle = false;
    f.editor.pluginMethod_DeckwiseQuiesceAndDrain(id);
    f.editor.clean = true;
    f.ack();
    expect(f.save).not.toHaveBeenCalled();
    expect(f.editor.pluginMethod_DeckwiseBarrierStatus(id).state).toBe('draining');
    f.editor.idle = true;
    vi.advanceTimersByTime(10);
    expect(f.save).toHaveBeenCalledOnce();
    f.ack();
    expect(f.editor.pluginMethod_DeckwiseBarrierStatus(id).state).toBe('held');
  });
  it('preserves an existing view restriction and returns an idempotent receipt', () => {
    vi.useFakeTimers();
    const f = fixture();
    f.editor.restriction = 132;
    f.editor.pluginMethod_DeckwiseQuiesceAndDrain(id);
    f.editor.pluginMethod_DeckwiseQuiesceAndDrain(id);
    expect(f.save).toHaveBeenCalledOnce();
    expect(f.editor.pluginMethod_DeckwiseQuiesceAndDrain('00000000-0000-0000-0000-000000000002')).toMatchObject({
      error: 'NATIVE_BARRIER_BUSY',
    });
    f.editor.clean = true;
    f.ack();
    f.editor.pluginMethod_DeckwiseResume(id);
    expect(f.editor.restriction).toBe(132);
  });
  it('refuses disconnected start and invalid identifiers', () => {
    const f = fixture();
    f.editor.connected = false;
    expect(f.editor.pluginMethod_DeckwiseQuiesceAndDrain(id)).toMatchObject({ error: 'NATIVE_BARRIER_UNAVAILABLE' });
    expect(f.editor.pluginMethod_DeckwiseQuiesceAndDrain('not-an-id')).toMatchObject({
      error: 'NATIVE_BARRIER_INVALID',
    });
  });
  it('does not accept stale indexes, lost restriction or disconnected completion', () => {
    vi.useFakeTimers();
    const f = fixture();
    f.editor.pluginMethod_DeckwiseQuiesceAndDrain(id);
    f.editor.clean = true;
    f.ack(-1);
    expect(f.editor.pluginMethod_DeckwiseBarrierStatus(id).state).toBe('draining');
    f.editor.restriction = 4;
    f.ack();
    expect(f.editor.pluginMethod_DeckwiseBarrierStatus(id)).toMatchObject({
      state: 'uncertain-held',
      error: 'NATIVE_BARRIER_RESTRICTION_LOST',
    });
  });
});

it('leaves unsupported/offline SDK installation inert', () => {
  const prototype = {};
  installDocsNativeBarrier930138({ isOffline: true, Asc: { asc_docs_api: { prototype } } });
  installDocsNativeBarrier930138({});
  expect(Object.keys(prototype)).toEqual([]);
  expect(Object.getOwnPropertyNames(prototype)).toEqual([]);
});
