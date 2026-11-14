import { afterEach, expect, it, vi } from 'vitest';

const { initialize, mount, setOpen, context } = vi.hoisted(() => ({
  initialize: vi.fn(),
  setOpen: vi.fn(),
  context: vi.fn(),
  mount: vi.fn(),
}));
vi.mock('../../lib/deckwise-native-teaching', () => ({ initializeNativeTeaching: initialize }));
vi.mock('../../lib/deckwise-teaching-sidebar', () => ({ mountDeckwiseTeachingSidebar: mount }));
vi.mock('../../lib/deckwise-editor-bridge', () => ({ deckwiseEditorBridge: { getContext: context, reset() {} } }));

let dispose: (() => void) | undefined;
afterEach(() => {
  dispose?.();
  dispose = undefined;
  vi.clearAllMocks();
});

it.each([true, false])(
  'acknowledges native readiness with teachingOpen=%s and no sidebar DOM',
  async (teachingOpen) => {
    mount.mockReturnValue({ setOpen, dispose() {} });
    const { mountEditorModule } = await import('../../lib/embed-api');
    const send = vi.fn();
    let command!: (value: unknown) => void;
    const module = mountEditorModule({
      deckwise: true,
      send,
      subscribe(listener) {
        command = listener;
        return () => {};
      },
    });
    dispose = () => module.dispose();
    command({ type: 'document:configure-ui', payload: { teachingOpen } });
    command({
      type: 'document:teaching-enable',
      payload: { binding: { projectId: 'p', documentId: 'd', sessionId: 's' }, nativeTeaching: true },
    });
    await vi.waitFor(() => expect(send).toHaveBeenCalledWith('document:teaching-ready', {}, undefined));
    expect(initialize).toHaveBeenCalledOnce();
    expect(setOpen).toHaveBeenCalledWith(teachingOpen);
  },
);

it('does not acknowledge a native initialization failure', async () => {
  initialize.mockImplementationOnce(() => {
    throw new Error('Native SDK unavailable');
  });
  const { mountEditorModule } = await import('../../lib/embed-api');
  const send = vi.fn();
  let command!: (value: unknown) => void;
  const module = mountEditorModule({
    deckwise: true,
    send,
    subscribe(listener) {
      command = listener;
      return () => {};
    },
  });
  dispose = () => module.dispose();
  command({ id: 'enable', type: 'document:teaching-enable', payload: { binding: {}, nativeTeaching: true } });
  await vi.waitFor(() =>
    expect(send).toHaveBeenCalledWith('document:error', { message: 'Native SDK unavailable' }, 'enable'),
  );
  expect(send.mock.calls.some((call) => call[0] === 'document:teaching-ready')).toBe(false);
});
