/** Fixed native plugin methods. This module must run inside the online slide SDK. */
export interface NativeBarrierReceipt {
  barrierId: string;
  state: 'draining' | 'held' | 'uncertain-held' | 'resumed';
  serverIndex?: number;
  error?: string;
}

export interface NativeBarrierAdapter<Editor extends object> {
  supported(editor: Editor): boolean;
  connected(editor: Editor): boolean;
  restrictions(editor: Editor): number;
  addView(editor: Editor): void;
  removeView(editor: Editor): void;
  idle(editor: Editor): boolean;
  clean(editor: Editor): boolean;
  save(editor: Editor): boolean;
  roundTrip(editor: Editor, failed: () => void): void;
  observeAck(editor: Editor, received: (index: number) => void): () => void;
  blockPluginWrites(editor: Editor, blocked: () => boolean): () => void;
}

const ids = /^[a-f0-9]{8}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{12}$/;

/** Timers only retry or report uncertainty; only a native server ACK can produce held. */
export function installDocsNativeBarrier<Editor extends object>(
  prototype: object,
  adapter: NativeBarrierAdapter<Editor>,
  options: { timeoutMs?: number; retryMs?: number } = {},
): void {
  const timeoutMs = options.timeoutMs ?? 30_000;
  const retryMs = options.retryMs ?? 50;
  const entries = new WeakMap<
    Editor,
    {
      receipt: NativeBarrierReceipt;
      ownsView: boolean;
      requested: boolean;
      deadline: number;
      timer?: ReturnType<typeof setTimeout>;
      stopAck: () => void;
      stopPlugins: () => void;
    }
  >();
  const copy = (entry: { receipt: NativeBarrierReceipt }) => ({ ...entry.receipt });
  const fail = (entry: { receipt: NativeBarrierReceipt }, error: string) => {
    entry.receipt.state = 'uncertain-held';
    entry.receipt.error = error;
  };
  const verify = (value: unknown): string => {
    if (typeof value !== 'string' || !ids.test(value)) throw new Error('NATIVE_BARRIER_INVALID');
    return value;
  };
  // Native plugin dispatch does not catch method exceptions. Always return a finite error.
  const safe = (run: (this: Editor, value: unknown) => NativeBarrierReceipt) =>
    function (this: Editor, value: unknown) {
      try {
        return run.call(this, value);
      } catch (error) {
        return {
          error:
            error instanceof Error && /^NATIVE_BARRIER_[A-Z_]+$/.test(error.message)
              ? error.message
              : 'NATIVE_BARRIER_UNAVAILABLE',
        };
      }
    };
  Object.defineProperties(prototype, {
    pluginMethod_DeckwiseQuiesceAndDrain: {
      configurable: false,
      value: safe(function (this: Editor, value: unknown) {
        const barrierId = verify(value);
        const existing = entries.get(this);
        if (existing && existing.receipt.state !== 'resumed') {
          if (existing.receipt.barrierId !== barrierId) throw new Error('NATIVE_BARRIER_BUSY');
          return copy(existing);
        }
        if (!adapter.supported(this) || !adapter.connected(this)) throw new Error('NATIVE_BARRIER_UNAVAILABLE');
        const before = adapter.restrictions(this);
        if (!Number.isSafeInteger(before) || before < 0) throw new Error('NATIVE_BARRIER_UNAVAILABLE');
        const receipt: NativeBarrierReceipt = { barrierId, state: 'draining' };
        const entry = {
          receipt,
          ownsView: (before & 128) === 0,
          requested: false,
          deadline: Date.now() + timeoutMs,
          timer: undefined as ReturnType<typeof setTimeout> | undefined,
          stopAck: () => {},
          stopPlugins: () => {},
        };
        entries.set(this, entry);
        const retry = () => {
          if (entry.receipt.state !== 'draining') return;
          if (!adapter.connected(this)) {
            fail(entry, 'NATIVE_BARRIER_DISCONNECTED');
            return;
          }
          if (Date.now() >= entry.deadline) {
            fail(entry, 'NATIVE_BARRIER_ACK_UNKNOWN');
            return;
          }
          if (!entry.requested && adapter.idle(this)) {
            // A clean document still needs a real lock/unlock round trip, not a synthetic success.
            entry.requested = true;
            try {
              if (!adapter.save(this)) {
                if (!adapter.clean(this)) {
                  fail(entry, 'NATIVE_BARRIER_SAVE_UNAVAILABLE');
                  return;
                }
                adapter.roundTrip(this, () => fail(entry, 'NATIVE_BARRIER_SAVE_UNKNOWN'));
              }
            } catch {
              fail(entry, 'NATIVE_BARRIER_SAVE_UNKNOWN');
              return;
            }
          }
          entry.timer = setTimeout(retry, retryMs);
        };
        try {
          // This synchronous section cannot interleave with native keyboard/plugin mutation.
          entry.stopPlugins = adapter.blockPluginWrites(this, () => entry.receipt.state !== 'resumed');
          adapter.addView(this);
          if ((adapter.restrictions(this) & 128) !== 128) throw new Error('NATIVE_BARRIER_RESTRICTION_FAILED');
          entry.stopAck = adapter.observeAck(this, (index) => {
            if (entry.receipt.state === 'resumed' || !entry.requested) return;
            if (!Number.isSafeInteger(index) || index < 0 || !adapter.connected(this) || !adapter.clean(this)) return;
            if ((adapter.restrictions(this) & 128) !== 128) {
              fail(entry, 'NATIVE_BARRIER_RESTRICTION_LOST');
              return;
            }
            if (entry.timer) clearTimeout(entry.timer);
            entry.receipt = { barrierId, state: 'held', serverIndex: index };
          });
          retry();
        } catch {
          fail(entry, 'NATIVE_BARRIER_SETUP_FAILED');
        }
        return copy(entry);
      }),
    },
    pluginMethod_DeckwiseBarrierStatus: {
      configurable: false,
      value: safe(function (this: Editor, value: unknown) {
        const barrierId = verify(value),
          entry = entries.get(this);
        if (!entry || entry.receipt.barrierId !== barrierId) throw new Error('NATIVE_BARRIER_NOT_FOUND');
        if (
          entry.receipt.state === 'held' &&
          (!adapter.connected(this) || !adapter.clean(this) || (adapter.restrictions(this) & 128) !== 128)
        )
          fail(entry, 'NATIVE_BARRIER_STATE_CHANGED');
        return copy(entry);
      }),
    },
    pluginMethod_DeckwiseResume: {
      configurable: false,
      value: safe(function (this: Editor, value: unknown) {
        const barrierId = verify(value),
          entry = entries.get(this);
        if (!entry || entry.receipt.barrierId !== barrierId) throw new Error('NATIVE_BARRIER_NOT_FOUND');
        if (entry.receipt.state === 'resumed') return copy(entry);
        // Unknown submission must first reconcile its actual ACK, never discard data on resume.
        if (entry.receipt.state !== 'held' || !adapter.connected(this) || !adapter.clean(this)) {
          throw new Error('NATIVE_BARRIER_ACK_UNKNOWN');
        }
        if (entry.ownsView) adapter.removeView(this);
        if (entry.timer) clearTimeout(entry.timer);
        entry.stopAck();
        entry.stopPlugins();
        entry.receipt.state = 'resumed';
        return copy(entry);
      }),
    },
  });
}
