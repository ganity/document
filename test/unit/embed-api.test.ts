import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { webcrypto, createHash } from 'node:crypto';
import { File as NodeFile } from 'node:buffer';
import { DOCUMENT_TYPE_MAP } from '@ranuts/shared/document-utils';

const mockOpenDocumentFromUrl = vi.fn().mockResolvedValue(undefined);
const mockLoadEditorApi = vi.fn().mockResolvedValue(undefined);
const mockHandleDocumentOperation = vi.fn().mockResolvedValue(undefined);
const mockDestroyEditorInstance = vi.fn();
const mockGetReadonlyMode = vi.fn().mockReturnValue(false);
const mockSetReadonlyMode = vi.fn();
const mockRequestSaveDocument = vi.fn();
const mockHasPendingEmbeddedSaveRequest = vi.fn().mockReturnValue(false);
const mockRenderDeckWiseCurrentSlidePreview = vi.fn();

vi.mock('../../lib/document', () => ({ openDocumentFromUrl: mockOpenDocumentFromUrl }));
vi.mock('../../lib/deckwise-slide-preview', () => ({
  renderDeckWiseCurrentSlidePreview: mockRenderDeckWiseCurrentSlidePreview,
}));
vi.mock('../../lib/converter', () => ({
  loadEditorApi: mockLoadEditorApi,
  handleDocumentOperation: mockHandleDocumentOperation,
}));
vi.mock('../../lib/onlyoffice-editor', () => ({
  destroyEditorInstance: mockDestroyEditorInstance,
  getReadonlyMode: mockGetReadonlyMode,
  setReadonlyMode: mockSetReadonlyMode,
  requestSaveDocument: mockRequestSaveDocument,
}));

vi.mock('../../lib/onlyoffice/save-stream', async (importOriginal) => ({
  ...(await importOriginal<typeof import('../../lib/onlyoffice/save-stream')>()),
  hasPendingEmbeddedSaveRequest: mockHasPendingEmbeddedSaveRequest,
}));

async function dispatchMessage(
  data: unknown,
  origin = 'https://parent.example.com',
  source: MessageEventSource | null = window.parent,
) {
  window.dispatchEvent(new MessageEvent('message', { data, origin, source }));
  await new Promise((r) => setTimeout(r, 0));
}

// Verify a specific message type + id was posted (content-based, not count-based).
// Using unique IDs per test prevents false positives from accumulated old listeners.
function expectMessagePosted(
  spy: ReturnType<typeof vi.spyOn>,
  type: string,
  id: string,
  payloadMatch?: Record<string, unknown>,
) {
  const found = spy.mock.calls.find((call: unknown[]) => {
    const msg = call[0] as { type?: string; id?: string };
    return msg?.type === type && msg?.id === id;
  });
  expect(found, `Expected message type="${type}" id="${id}" to have been posted`).toBeDefined();
  if (payloadMatch) {
    const msg = found![0] as { payload?: Record<string, unknown> };
    expect(msg.payload).toMatchObject(payloadMatch);
  }
}

// Stale listeners from earlier module instances answer the same id with their
// own state, so look for the current instance's answer among all of them.
type PostedMessage = { type?: string; id?: string; payload?: Record<string, unknown> };

function postedMessages(spy: ReturnType<typeof vi.spyOn>): PostedMessage[] {
  return (spy.mock.calls as unknown[][]).map((call) => call[0] as PostedMessage);
}

function postedPayloads(spy: ReturnType<typeof vi.spyOn>, type: string, id: string) {
  return postedMessages(spy)
    .filter((msg) => msg?.type === type && msg?.id === id)
    .map((msg) => msg.payload);
}

function expectMessageNotPosted(spy: ReturnType<typeof vi.spyOn>, id: string) {
  const found = spy.mock.calls.find((call: unknown[]) => {
    const msg = call[0] as { id?: string };
    return msg?.id === id;
  });
  expect(found, `Expected no message with id="${id}" to be posted`).toBeUndefined();
}

describe('embed-api', () => {
  let postMessageSpy: ReturnType<typeof vi.spyOn>;
  let listenerSpy: ReturnType<typeof vi.spyOn>;

  beforeEach(() => {
    vi.resetModules();
    vi.clearAllMocks();
    mockHasPendingEmbeddedSaveRequest.mockReturnValue(false);
    postMessageSpy = vi.spyOn(window, 'postMessage');
    listenerSpy = vi.spyOn(window, 'addEventListener');
    document.body.classList.remove('embed-mode');
    window.history.pushState({}, '', '/');
  });

  afterEach(() => {
    for (const [type, listener, options] of listenerSpy.mock.calls) window.removeEventListener(type, listener, options);
    listenerSpy.mockRestore();
    postMessageSpy.mockRestore();
    delete (window as any).editor;
  });

  describe('desktop module lifecycle', () => {
    it('uses a direct capability, validates UI updates and detaches on disposal', async () => {
      const { mountEditorModule } = await import('../../lib/embed-api');
      const { isEditorModule, isEmbedMode, isDeckWiseMode } = await import('../../lib/embed-mode');
      const send = vi.fn(), unsubscribe = vi.fn();
      let command!: (value: unknown) => void;
      const editor = mountEditorModule({ deckwise: true, send, subscribe(listener) { command = listener; return unsubscribe; } });
      expect(send).toHaveBeenCalledWith('document:ready', {}, undefined);
      expect([isEditorModule(), isEmbedMode(), isDeckWiseMode()]).toEqual([true, true, true]);
      expect(() => mountEditorModule({ deckwise: true, send, subscribe: () => unsubscribe })).toThrow('already mounted');
      await dispatchMessage({ type: 'document:configure-ui', payload: { theme: 'dark' } });
      expect(send).not.toHaveBeenCalledWith('document:ui-configured', expect.anything(), expect.anything());
      command({ type: 'document:configure-ui', payload: { theme: 'dark', teachingOpen: false }, id: 'theme' });
      expect(document.documentElement.getAttribute('data-ran-theme')).toBe('dark');
      expect(send).toHaveBeenCalledWith('document:ui-configured', { theme: 'dark', teachingOpen: false }, 'theme');
      command({ type: 'document:configure-ui', payload: { theme: 'evil', path: '/etc/passwd' }, id: 'invalid' });
      expect(send).toHaveBeenCalledWith('document:error', { message: 'Invalid editor UI configuration' }, 'invalid');
      editor.dispose(); editor.dispose();
      expect(unsubscribe).toHaveBeenCalledTimes(1);
      expect(mockDestroyEditorInstance).toHaveBeenCalledTimes(1);
      expect(isEditorModule()).toBe(false);
      expect(document.body.classList.contains('deckwise-mode')).toBe(false);
      document.documentElement.removeAttribute('data-ran-theme');
    });

    it('cannot send a late open result into a replacement module', async () => {
      const { mountEditorModule } = await import('../../lib/embed-api');
      let release!: () => void;
      mockLoadEditorApi.mockReturnValueOnce(new Promise<void>(resolve => { release = resolve; }));
      let command!: (value: unknown) => void;
      const firstSend = vi.fn(), secondSend = vi.fn();
      const first = mountEditorModule({ deckwise: false, send: firstSend, subscribe(listener) { command = listener; return () => {}; } });
      command({ type: 'document:open-buffer', id: 'old-open', payload: { fileName: 'test.docx', buffer: new Uint8Array([1]).buffer } });
      first.dispose();
      const second = mountEditorModule({ deckwise: false, send: secondSend, subscribe: () => () => {} });
      release(); await new Promise(resolve => setTimeout(resolve, 0));
      expect(secondSend.mock.calls.map(call => call[0])).toEqual(['document:ready']);
      expect(mockHandleDocumentOperation).not.toHaveBeenCalled();
      second.dispose();
    });
  });

  describe('DeckWise command messages', () => {
    const owner = { projectId: 'project', documentId: 'document', sessionId: 'session' };
    it('renders a page preview only from the bound parent and outside save transitions', async () => {
      window.history.pushState({}, '', '/?embed=1&deckwise=1&embedOrigin=https%3A%2F%2Fparent.example.com');
      const png = new Uint8Array([137, 80, 78, 71]);
      mockRenderDeckWiseCurrentSlidePreview.mockReturnValue({ page: 2, png });
      const { initEmbedApi } = await import('../../lib/embed-api');
      initEmbedApi();
      const payload = { binding: owner, editGeneration: 8 };
      await dispatchMessage(
        { id: 'preview-foreign', type: 'document:render-page-preview', payload },
        'https://evil.example',
      );
      expect(mockRenderDeckWiseCurrentSlidePreview).not.toHaveBeenCalled();
      await dispatchMessage({ id: 'preview-current', type: 'document:render-page-preview', payload });
      expect(mockRenderDeckWiseCurrentSlidePreview).toHaveBeenCalledWith(owner, 8, undefined);
      expectMessagePosted(postMessageSpy, 'document:page-preview', 'preview-current', { page: 2 });
      await dispatchMessage({
        id: 'preview-textless',
        type: 'document:render-page-preview',
        payload: { ...payload, omitNativeText: true },
      });
      expect(mockRenderDeckWiseCurrentSlidePreview).toHaveBeenCalledWith(owner, 8, true);
      await dispatchMessage({
        id: 'preview-page-two',
        type: 'document:render-page-preview',
        payload: { ...payload, page: 2 },
      });
      expect(mockRenderDeckWiseCurrentSlidePreview).toHaveBeenCalledWith(owner, 8, undefined, 2);
      mockRenderDeckWiseCurrentSlidePreview.mockClear();
      for (const page of [0, -1, 1.5, '2', 5001]) {
        await dispatchMessage({
          id: 'preview-bad-page',
          type: 'document:render-page-preview',
          payload: { ...payload, page },
        });
      }
      expect(mockRenderDeckWiseCurrentSlidePreview).not.toHaveBeenCalled();
      await dispatchMessage({
        id: 'preview-malformed',
        type: 'document:render-page-preview',
        payload: { ...payload, omitNativeText: 'true' },
      });
      expect(mockRenderDeckWiseCurrentSlidePreview).not.toHaveBeenCalled();
      expectMessagePosted(postMessageSpy, 'document:error', 'preview-malformed', {
        message: 'Invalid page preview request',
      });
      await dispatchMessage({
        id: 'preview-unknown-field',
        type: 'document:render-page-preview',
        payload: { ...payload, path: '/tmp/source.pptx' },
      });
      expectMessagePosted(postMessageSpy, 'document:error', 'preview-unknown-field', {
        message: 'Invalid page preview request',
      });
      mockHasPendingEmbeddedSaveRequest.mockReturnValueOnce(true);
      await dispatchMessage({ id: 'preview-saving', type: 'document:render-page-preview', payload });
      expectMessagePosted(postMessageSpy, 'document:error', 'preview-saving', {
        message: 'Finish or recover the current save before rendering a page preview',
      });
      mockHasPendingEmbeddedSaveRequest.mockReturnValue(false);
    });
    it('routes Host-only fade review/execution and rejects dirty, saving or foreign messages', async () => {
      window.history.pushState({}, '', '/?embed=1&deckwise=1&embedOrigin=https%3A%2F%2Fparent.example.com');
      const { deckwiseEditorBridge } = await import('../../lib/deckwise-editor-bridge');
      const { markDocumentDirty, markDocumentSaved } = await import('../../lib/unsaved-guard');
      markDocumentSaved();
      const review = vi.spyOn(deckwiseEditorBridge, 'reviewFade').mockReturnValue({ ...owner } as never);
      const execute = vi.spyOn(deckwiseEditorBridge, 'executeFade').mockResolvedValue({ ...owner } as never);
      const revoke = vi.spyOn(deckwiseEditorBridge, 'revokeFade');
      const { initEmbedApi } = await import('../../lib/embed-api');
      initEmbedApi();
      const authorization = {
        editGeneration: 1,
        page: 1,
        sourceHash: 'sha256:' + 'a'.repeat(64),
        slideHash: 'sha256:' + 'b'.repeat(64),
      };
      const command = { ...owner, tool: 'set_slide_fade', editGeneration: 1, requestId: 'fade-execute', input: {} };
      const payload = { binding: owner, authorization };
      await dispatchMessage({ id: 'fade-foreign', type: 'document:review-fade', payload }, 'https://evil.example');
      expect(review).not.toHaveBeenCalled();
      await dispatchMessage({ id: 'fade-review', type: 'document:review-fade', payload });
      expect(review).toHaveBeenCalledWith(owner, authorization);
      expectMessagePosted(postMessageSpy, 'document:context', 'fade-review');
      await dispatchMessage({ id: 'fade-execute', type: 'document:execute-fade', payload: { command, authorization } });
      expect(execute).toHaveBeenCalledWith(command, authorization, 'fade-execute');
      expectMessagePosted(postMessageSpy, 'document:command-result', 'fade-execute');
      review.mockClear();
      execute.mockClear();
      for (const type of ['document:review-fade', 'document:execute-fade']) {
        const correct = type === 'document:review-fade' ? payload : { command, authorization };
        await dispatchMessage({
          id: 'fade-extra-' + type.replace(':', '-'),
          type,
          payload: { ...correct, script: 'bad' },
        });
        expectMessagePosted(postMessageSpy, 'document:error', 'fade-extra-' + type.replace(':', '-'));
        markDocumentDirty();
        await dispatchMessage({ id: 'fade-dirty-' + type.replace(':', '-'), type, payload: correct });
        expectMessagePosted(postMessageSpy, 'document:error', 'fade-dirty-' + type.replace(':', '-'), {
          message: 'Save the latest PPTX before fade inspection',
        });
        markDocumentSaved();
        mockHasPendingEmbeddedSaveRequest.mockReturnValue(true);
        await dispatchMessage({ id: 'fade-saving-' + type.replace(':', '-'), type, payload: correct });
        expectMessagePosted(postMessageSpy, 'document:error', 'fade-saving-' + type.replace(':', '-'));
        mockHasPendingEmbeddedSaveRequest.mockReturnValue(false);
      }
      expect(review).not.toHaveBeenCalled();
      expect(execute).not.toHaveBeenCalled();
      expect(revoke).toHaveBeenCalled();
      await dispatchMessage({ id: 'fade-revoke-readonly', type: 'document:set-readonly', payload: { readonly: true } });
      expect(revoke).toHaveBeenCalled();
      markDocumentSaved();
    });
    it('routes target-bound image review only from the trusted parent with a distinct response', async () => {
      window.history.pushState({}, '', '/?embed=1&deckwise=1&embedOrigin=https%3A%2F%2Fparent.example.com');
      const { deckwiseEditorBridge } = await import('../../lib/deckwise-editor-bridge');
      const context = { ...owner, editGeneration: 9, page: 1 };
      const review = vi.spyOn(deckwiseEditorBridge, 'reviewImageTarget').mockReturnValue(context as never);
      const legacy = vi.spyOn(deckwiseEditorBridge, 'reviewImage');
      const { initEmbedApi } = await import('../../lib/embed-api');
      initEmbedApi();
      const asset = { assetId: 'reviewed_png_1' },
        target = { page: 1, editGeneration: 8, operation: 'replace' };
      const payload = { binding: owner, asset, target };
      await dispatchMessage(
        { id: 'target-foreign', type: 'document:review-image-target', payload },
        'https://evil.example',
      );
      await dispatchMessage(
        { id: 'target-sibling', type: 'document:review-image-target', payload },
        'https://parent.example.com',
        null,
      );
      expect(review).not.toHaveBeenCalled();
      expectMessageNotPosted(postMessageSpy, 'target-foreign');
      expectMessageNotPosted(postMessageSpy, 'target-sibling');
      await dispatchMessage({ id: 'target-current', type: 'document:review-image-target', payload });
      expect(review).toHaveBeenCalledWith(owner, asset, target);
      expectMessagePosted(postMessageSpy, 'document:image-target-reviewed', 'target-current', context);
      expect(postedPayloads(postMessageSpy, 'document:context', 'target-current')).toEqual([]);
      expect(legacy).not.toHaveBeenCalled();
      expect(mockRequestSaveDocument).not.toHaveBeenCalled();
      review.mockClear();
      await dispatchMessage({
        id: 'target-extra',
        type: 'document:review-image-target',
        payload: { ...payload, bytes: new Uint8Array(33) },
      });
      expectMessagePosted(postMessageSpy, 'document:error', 'target-extra', {
        message: 'Invalid image target review payload',
      });
      await dispatchMessage({
        id: 'target-missing',
        type: 'document:review-image-target',
        payload: { binding: owner, asset },
      });
      expectMessagePosted(postMessageSpy, 'document:error', 'target-missing', {
        message: 'Invalid image target review payload',
      });
      mockHasPendingEmbeddedSaveRequest.mockReturnValue(true);
      await dispatchMessage({ id: 'target-saving', type: 'document:review-image-target', payload });
      expectMessagePosted(postMessageSpy, 'document:error', 'target-saving', {
        message: 'Finish or recover the current save before using editor tools',
      });
      expect(review).not.toHaveBeenCalled();
      mockHasPendingEmbeddedSaveRequest.mockReturnValue(false);
      review.mockImplementationOnce(() => {
        throw new Error('stale image target');
      });
      await dispatchMessage({ id: 'target-stale', type: 'document:review-image-target', payload });
      expectMessagePosted(postMessageSpy, 'document:error', 'target-stale', { message: 'stale image target' });
      expect(postedPayloads(postMessageSpy, 'document:image-target-reviewed', 'target-stale')).toEqual([]);
    });
    it('routes reviewed image messages from the bound parent only', async () => {
      window.history.pushState({}, '', '/?embed=1&deckwise=1&embedOrigin=https%3A%2F%2Fparent.example.com');
      const { deckwiseEditorBridge } = await import('../../lib/deckwise-editor-bridge');
      const review = vi.spyOn(deckwiseEditorBridge, 'reviewImage').mockReturnValue({ ...owner } as never);
      const execute = vi.spyOn(deckwiseEditorBridge, 'executeImage').mockResolvedValue({ ...owner } as never);
      const { initEmbedApi } = await import('../../lib/embed-api');
      initEmbedApi();
      await dispatchMessage(
        { id: 'image-foreign', type: 'document:review-image', payload: { binding: owner, asset: null } },
        'https://evil.example',
      );
      expect(review).not.toHaveBeenCalled();
      await dispatchMessage({
        id: 'image-review',
        type: 'document:review-image',
        payload: { binding: owner, asset: null },
      });
      expect(review).toHaveBeenCalledWith(owner, null);
      expectMessagePosted(postMessageSpy, 'document:context', 'image-review');
      const command = { tool: 'replace_slide_image_from_asset' };
      const bytes = new Uint8Array(33);
      await dispatchMessage({ id: 'image-execute', type: 'document:execute-image', payload: { command, bytes } });
      expect(execute).toHaveBeenCalledWith(command, bytes, 'image-execute');
      expectMessagePosted(postMessageSpy, 'document:command-result', 'image-execute');
      await dispatchMessage({
        id: 'image-private',
        type: 'document:review-image',
        payload: { binding: owner, asset: null, token: 'invalid' },
      });
      expectMessagePosted(postMessageSpy, 'document:error', 'image-private');
    });
    it('refuses a host save while an editor command is awaiting asset bytes', async () => {
      window.history.pushState({}, '', '/?embed=1&deckwise=1');
      const { deckwiseEditorBridge } = await import('../../lib/deckwise-editor-bridge');
      const executing = vi.spyOn(deckwiseEditorBridge, 'isExecuting', 'get').mockReturnValue(true);
      const { initEmbedApi } = await import('../../lib/embed-api');
      initEmbedApi();
      await dispatchMessage({ id: 'save-during-command', type: 'document:save', payload: {} });
      expectMessagePosted(postMessageSpy, 'document:error', 'save-during-command', {
        message: 'Finish the current editor command before saving',
      });
      expect(mockRequestSaveDocument).not.toHaveBeenCalled();
      executing.mockRestore();
    });
    it('routes context and command results only for the bound parent source and origin', async () => {
      window.history.pushState({}, '', '/?embed=1&deckwise=1&embedOrigin=https%3A%2F%2Fparent.example.com');
      const { deckwiseEditorBridge } = await import('../../lib/deckwise-editor-bridge');
      const context = vi.spyOn(deckwiseEditorBridge, 'getContext').mockReturnValue({
        ...owner,
        editGeneration: 1,
        kind: 'slide',
        page: 1,
        selectedText: '',
        selectionCharacters: 0,
        readOnly: false,
        capabilities: [],
      });
      const execute = vi.spyOn(deckwiseEditorBridge, 'execute').mockResolvedValue({
        ...owner,
        editGeneration: 1,
        requestId: 'command-good',
        status: 'sent',
        tool: 'get_presentation_text',
        result: { text: 'lesson' },
      });
      const { initEmbedApi } = await import('../../lib/embed-api');
      initEmbedApi();
      await dispatchMessage(
        { id: 'context-bad-origin', type: 'document:get-context', payload: owner },
        'https://evil.example',
      );
      await dispatchMessage(
        { id: 'context-bad-source', type: 'document:get-context', payload: owner },
        'https://parent.example.com',
        null,
      );
      await dispatchMessage(
        { id: 'command-bad-origin', type: 'document:execute', payload: owner },
        'https://evil.example',
      );
      expect(context).not.toHaveBeenCalled();
      expect(execute).not.toHaveBeenCalled();
      await dispatchMessage({ id: 'context-good', type: 'document:get-context', payload: owner });
      expect(context).toHaveBeenCalledWith(owner);
      expectMessagePosted(postMessageSpy, 'document:context', 'context-good', { ...owner, editGeneration: 1 });
      await dispatchMessage({ id: 'command-good', type: 'document:execute', payload: owner });
      expect(execute).toHaveBeenCalledWith(owner, 'command-good');
      expectMessagePosted(postMessageSpy, 'document:command-result', 'command-good', { requestId: 'command-good' });
      context.mockRestore();
      execute.mockRestore();
    });
    it('checks the current editor generation before routing narration writes and preserves existing callers', async () => {
      window.history.pushState({}, '', '/?embed=1&deckwise=1&embedOrigin=https%3A%2F%2Fparent.example.com');
      const { deckwiseEditorBridge } = await import('../../lib/deckwise-editor-bridge');
      const router = await import('../../lib/deckwise-native-teaching-router');
      const context = vi.spyOn(deckwiseEditorBridge, 'getContext').mockReturnValue({
        ...owner, editGeneration: 4, kind: 'slide', page: 1, selectedText: '', selectionCharacters: 0, readOnly: false, capabilities: [],
      });
      const route = vi.spyOn(router, 'routeNativeTeaching').mockResolvedValue({ value: { saved: true } });
      try {
        const { initEmbedApi } = await import('../../lib/embed-api');
        initEmbedApi();
        for (const expectedEditGeneration of [3, null, '4', 4.5]) {
          const id = 'native-stale-' + String(expectedEditGeneration);
          await dispatchMessage({ id, type: 'document:native-teaching', payload: {
            binding: owner, operation: 'route', routeOperation: 'write', page: 1, args: { expectedEditGeneration },
          } });
          expectMessagePosted(postMessageSpy, 'document:error', id, { message: 'Native editor generation changed before narration write' });
        }
        expect(route).not.toHaveBeenCalled();
        for (const args of [{ expectedEditGeneration: 4 }, {}]) {
          await dispatchMessage({ id: 'native-current-' + route.mock.calls.length, type: 'document:native-teaching', payload: {
            binding: owner, operation: 'route', routeOperation: 'write', page: 1, args,
          } });
        }
        expect(route).toHaveBeenCalledTimes(2);
      } finally { context.mockRestore(); route.mockRestore(); }
    });
    it('rejects command messages outside DeckWise mode', async () => {
      window.history.pushState({}, '', '/?embed=1');
      const { initEmbedApi } = await import('../../lib/embed-api');
      initEmbedApi();
      await dispatchMessage({ id: 'not-deckwise', type: 'document:get-context', payload: owner });
      expectMessagePosted(postMessageSpy, 'document:error', 'not-deckwise', {
        message: 'Editor commands require DeckWise mode',
      });
      await dispatchMessage({ id: 'target-not-deckwise', type: 'document:review-image-target', payload: {} });
      expectMessagePosted(postMessageSpy, 'document:error', 'target-not-deckwise', {
        message: 'Editor commands require DeckWise mode',
      });
    });
  });

  describe('detectEmbedMode via URL params', () => {
    it.each([
      ['?embed=', true],
      ['?embed=1', true],
      ['?embed=true', true],
      ['?embedded=true', true],
      ['?embedded=1', true],
      ['', false],
      ['?other=1', false],
    ])('URL "%s" → embed mode = %s', async (search, expected) => {
      window.history.pushState({}, '', `/${search}`);
      const { initEmbedApi } = await import('../../lib/embed-api');
      initEmbedApi();
      expect(document.body.classList.contains('embed-mode')).toBe(expected);
    });
  });

  describe('initEmbedApi', () => {
    it('is idempotent - second call does not throw', async () => {
      window.history.pushState({}, '', '/?embed=1');
      const { initEmbedApi } = await import('../../lib/embed-api');
      expect(() => initEmbedApi()).not.toThrow();
      expect(() => initEmbedApi()).not.toThrow();
      expect(document.body.classList.contains('embed-mode')).toBe(true);
    });

    it('posts document:ready when the load event fires', async () => {
      window.history.pushState({}, '', '/?embed=1');
      const { initEmbedApi } = await import('../../lib/embed-api');
      initEmbedApi();

      window.dispatchEvent(new Event('load'));

      expect(postMessageSpy).toHaveBeenCalledWith(
        expect.objectContaining({ type: 'document:ready' }),
        expect.any(String),
      );
    });

    it('forwards the OnlyOffice save button to the DeckWise host', async () => {
      window.history.pushState({}, '', '/?embed=1&deckwise=1');
      const { initEmbedApi } = await import('../../lib/embed-api');
      initEmbedApi();

      window.dispatchEvent(new Event('deckwise:save-requested'));

      expect(postMessageSpy).toHaveBeenCalledWith(
        expect.objectContaining({ type: 'document:save-requested' }),
        expect.any(String),
      );
      expect(document.body.classList.contains('deckwise-mode')).toBe(true);
    });
  });

  describe('message handling', () => {
    it('ignores commands from a sibling frame even when its origin matches', async () => {
      window.history.pushState({}, '', '/?embed=1');
      const frame = document.createElement('iframe');
      document.body.appendChild(frame);
      try {
        const { initEmbedApi } = await import('../../lib/embed-api');
        initEmbedApi();
        await dispatchMessage(
          { type: 'document:get-state', id: 'sibling-forgery' },
          'https://parent.example.com',
          frame.contentWindow,
        );
        expectMessageNotPosted(postMessageSpy, 'sibling-forgery');
      } finally {
        frame.remove();
      }
    });

    it('keeps replies bound to the first parent origin', async () => {
      window.history.pushState({}, '', '/?embed=1');
      const { initEmbedApi } = await import('../../lib/embed-api');
      initEmbedApi();
      await dispatchMessage({ type: 'document:get-state', id: 'bind-parent' });
      await dispatchMessage({ type: 'document:get-state', id: 'parent-origin-changed' }, 'https://other.example.com');
      expectMessagePosted(postMessageSpy, 'document:state', 'bind-parent');
      expectMessageNotPosted(postMessageSpy, 'parent-origin-changed');
    });

    it('ignores malformed message types without an unhandled rejection', async () => {
      window.history.pushState({}, '', '/?embed=1');
      const { initEmbedApi } = await import('../../lib/embed-api');
      initEmbedApi();
      await dispatchMessage({ type: 123, id: 'malformed-type' });
      expectMessageNotPosted(postMessageSpy, 'malformed-type');
    });

    it('ignores messages that lack a document: prefix', async () => {
      window.history.pushState({}, '', '/?embed=1');
      const { initEmbedApi } = await import('../../lib/embed-api');
      initEmbedApi();

      await dispatchMessage({ type: 'other:ping', id: 'ignore-1' });
      await dispatchMessage(null);
      await dispatchMessage('plain string');

      expectMessageNotPosted(postMessageSpy, 'ignore-1');
    });

    it('ignores messages from disallowed origins when embedOrigin is set', async () => {
      window.history.pushState({}, '', '/?embed=1&embedOrigin=https://allowed.example.com');
      const { initEmbedApi } = await import('../../lib/embed-api');
      initEmbedApi();

      // The explicit origin allowlist rejects the command before it is handled.
      await dispatchMessage({ type: 'document:get-state', id: 'origin-block-1' }, 'https://evil.example.com');

      expectMessageNotPosted(postMessageSpy, 'origin-block-1');
    });

    it('accepts messages from a matching embedOrigin', async () => {
      window.history.pushState({}, '', '/?embed=1&embedOrigin=https://allowed.example.com');
      const { initEmbedApi } = await import('../../lib/embed-api');
      initEmbedApi();

      await dispatchMessage({ type: 'document:get-state', id: 'origin-allow-1' }, 'https://allowed.example.com');

      expectMessagePosted(postMessageSpy, 'document:state', 'origin-allow-1');
    });

    it('responds to document:get-state with readonly and hasDocument flags', async () => {
      window.history.pushState({}, '', '/?embed=1');
      mockGetReadonlyMode.mockReturnValue(false);
      const { initEmbedApi } = await import('../../lib/embed-api');
      initEmbedApi();

      await dispatchMessage({ type: 'document:get-state', id: 'state-1' });

      expectMessagePosted(postMessageSpy, 'document:state', 'state-1', {
        readonly: false,
        hasDocument: false,
      });
    });

    it('reports the unsaved-changes flag in document:state', async () => {
      window.history.pushState({}, '', '/?embed=1');
      const { initEmbedApi } = await import('../../lib/embed-api');
      const { markDocumentDirty } = await import('../../lib/unsaved-guard');
      initEmbedApi();

      markDocumentDirty();
      await dispatchMessage({ type: 'document:get-state', id: 'state-dirty-1' });

      expect(postedPayloads(postMessageSpy, 'document:state', 'state-dirty-1')).toContainEqual(
        expect.objectContaining({ dirty: true }),
      );
    });

    it('pushes document:dirty-changed when the flag flips', async () => {
      window.history.pushState({}, '', '/?embed=1');
      const { initEmbedApi } = await import('../../lib/embed-api');
      const { markDocumentDirty, markDocumentSaved } = await import('../../lib/unsaved-guard');
      initEmbedApi();

      markDocumentDirty();
      markDocumentDirty();
      markDocumentSaved();

      const dirtyMessages = postedMessages(postMessageSpy)
        .filter((msg) => msg?.type === 'document:dirty-changed')
        .map((msg) => msg.payload?.dirty);
      expect(dirtyMessages).toEqual([true, false]);
    });

    it('keeps the document dirty after a save it cannot pin to a history point', async () => {
      // No editor frame in jsdom: the save point cannot be captured, so the
      // host must not be told its copy is current.
      window.history.pushState({}, '', '/?embed=1');
      mockRequestSaveDocument.mockResolvedValue(new File(['x'], 'a.docx'));
      const { initEmbedApi } = await import('../../lib/embed-api');
      const { markDocumentDirty } = await import('../../lib/unsaved-guard');
      initEmbedApi();

      markDocumentDirty();
      await dispatchMessage({ type: 'document:save', id: 'save-dirty-1', payload: {} });
      mockRequestSaveDocument.mockReset();

      expect(postedPayloads(postMessageSpy, 'document:saved', 'save-dirty-1')).toContainEqual(
        expect.objectContaining({ dirty: true }),
      );
    });

    it('responds to document:set-readonly and updates readonly mode', async () => {
      window.history.pushState({}, '', '/?embed=1');
      const { initEmbedApi } = await import('../../lib/embed-api');
      initEmbedApi();

      await dispatchMessage({ type: 'document:set-readonly', id: 'ro-1', payload: { readonly: true } });

      expect(mockSetReadonlyMode).toHaveBeenCalledWith(true);
      expectMessagePosted(postMessageSpy, 'document:readonly-changed', 'ro-1');
    });

    it('posts document:error when a handler throws', async () => {
      window.history.pushState({}, '', '/?embed=1');
      mockRequestSaveDocument.mockRejectedValueOnce(new Error('Save failed'));
      const { initEmbedApi } = await import('../../lib/embed-api');
      initEmbedApi();

      await dispatchMessage({ type: 'document:save', id: 'err-1', payload: {} });

      expectMessagePosted(postMessageSpy, 'document:error', 'err-1', { message: 'Save failed' });
    });

    it.each(Object.keys(DOCUMENT_TYPE_MAP))('saves %s in its natural format in DeckWise mode', async (ext) => {
      window.history.pushState({}, '', '/?embed=1&deckwise=1');
      const legacy: Record<string, string> = { doc: 'docx', xls: 'xlsx', ppt: 'pptx' };
      const target = (legacy[ext] || ext).toUpperCase();
      const { setDocmentObj } = await import('@ranuts/shared/store');
      setDocmentObj({ fileName: `course.${ext}` });
      const savedFile = new File(['saved bytes'], `course.${target.toLowerCase()}`);
      mockRequestSaveDocument.mockResolvedValueOnce(savedFile);
      const { initEmbedApi } = await import('../../lib/embed-api');
      initEmbedApi();

      await dispatchMessage({ type: 'document:save', id: 'deckwise-save', payload: {} });
      const otherFormat = target === 'PDF' ? 'PPTX' : 'PDF';
      await dispatchMessage({ type: 'document:save', id: 'deckwise-export', payload: { targetExt: otherFormat } });

      expect(mockRequestSaveDocument).toHaveBeenCalledWith(
        target,
        expect.objectContaining({ returnOriginalOnTimeout: false }),
      );
      expectMessagePosted(postMessageSpy, 'document:saved', 'deckwise-save', { fileName: savedFile.name });
      expectMessagePosted(postMessageSpy, 'document:error', 'deckwise-export', {
        message: `DeckWise only permits ${target} save for this document`,
      });
    });

    it.each(Object.keys(DOCUMENT_TYPE_MAP))('opens %s buffers for preview in DeckWise mode', async (ext) => {
      window.history.pushState({}, '', '/?embed=1&deckwise=1');
      const { initEmbedApi } = await import('../../lib/embed-api');
      initEmbedApi();

      const fileName = `course.${ext.toUpperCase()}`;
      await dispatchMessage({
        type: 'document:open-buffer',
        id: 'deckwise-preview',
        payload: { fileName, buffer: new ArrayBuffer(4), readonly: true },
      });

      expect(mockHandleDocumentOperation).toHaveBeenCalledWith(expect.objectContaining({ fileName, readonly: true }));
      expectMessagePosted(postMessageSpy, 'document:opened', 'deckwise-preview');
    });

    it.each(['unknown', 'constructor', 'toString'])('rejects unsupported %s buffers in DeckWise mode', async (ext) => {
      window.history.pushState({}, '', '/?embed=1&deckwise=1');
      const { initEmbedApi } = await import('../../lib/embed-api');
      initEmbedApi();

      await dispatchMessage({
        type: 'document:open-buffer',
        id: 'deckwise-unsupported',
        payload: { fileName: `course.${ext}`, buffer: new ArrayBuffer(4) },
      });
      expect(mockHandleDocumentOperation).not.toHaveBeenCalled();
      expectMessagePosted(postMessageSpy, 'document:error', 'deckwise-unsupported', {
        message: 'Unsupported document format',
      });
    });

    it.each(['document:open-url', 'document:open-buffer'])(
      'rejects URL payloads via %s in DeckWise mode',
      async (type) => {
        window.history.pushState({}, '', '/?embed=1&deckwise=1');
        const { initEmbedApi } = await import('../../lib/embed-api');
        initEmbedApi();

        await dispatchMessage({
          type,
          id: 'deckwise-url',
          payload: { fileName: 'course.pptx', url: 'https://example.com/course.pptx' },
        });

        expect(mockOpenDocumentFromUrl).not.toHaveBeenCalled();
        expectMessagePosted(postMessageSpy, 'document:error', 'deckwise-url', {
          message: 'DeckWise only accepts document:open-buffer',
        });
      },
    );

    // GitHub #4 "save to remote server": the parent page's documented recovery path
    // is listening for this success payload and uploading `file` itself.
    it('posts document:saved with the saved File and its metadata (#4)', async () => {
      window.history.pushState({}, '', '/?embed=1');
      const savedFile = new File(['xlsx bytes'], 'report.xlsx', {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      });
      mockRequestSaveDocument.mockResolvedValueOnce(savedFile);
      const { initEmbedApi } = await import('../../lib/embed-api');
      initEmbedApi();

      await dispatchMessage({ type: 'document:save', id: 'save-1', payload: { targetExt: 'xlsx' } });

      expect(mockRequestSaveDocument).toHaveBeenCalledWith(
        'xlsx',
        expect.objectContaining({ returnOriginalOnTimeout: false }),
      );
      expectMessagePosted(postMessageSpy, 'document:saved', 'save-1', {
        file: savedFile,
        fileName: 'report.xlsx',
        mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        size: savedFile.size,
      });
    });

    it('opens a document from url payload', async () => {
      window.history.pushState({}, '', '/?embed=1');
      const { initEmbedApi } = await import('../../lib/embed-api');
      initEmbedApi();

      await dispatchMessage({
        type: 'document:open-url',
        id: 'open-1',
        payload: { url: 'https://example.com/test.xlsx' },
      });

      expect(mockOpenDocumentFromUrl).toHaveBeenCalledWith(
        'https://example.com/test.xlsx',
        undefined,
        expect.objectContaining({ readonly: false }),
      );
      expectMessagePosted(postMessageSpy, 'document:opened', 'open-1');
    });
  });

  it('captures an opted-in open baseline, returns the real save hash and withholds raced exports', async () => {
    vi.stubGlobal('crypto', webcrypto);
    vi.stubGlobal('File', NodeFile);
    window.history.pushState({}, '', '/?embed=1&deckwise=1');
    const { initEmbedApi } = await import('../../lib/embed-api');
    const { deckwisePageLineage } = await import('../../lib/deckwise-page-lineage');
    initEmbedApi();
    const history = { Index: 0, Points: [{ Items: [{}] }], Reset_SavedIndex: vi.fn() };
    mockHandleDocumentOperation.mockImplementationOnce(async ({ onContentReady }) => {
      const iframe = document.createElement('iframe');
      iframe.name = 'frameEditor';
      document.body.append(iframe);
      const api = {
        isDocumentLoadComplete: true,
        isLoadFullApi: true,
        WordControl: { m_oLogicDocument: { Slides: [{}, {}] } },
        CheckChangedDocument: vi.fn(),
      };
      Object.assign(iframe.contentWindow!, { editor: api, Asc: { editor: api }, AscCommon: { History: history } });
      onContentReady();
    });
    try {
      const binding = { projectId: 'project', documentId: 'document', sessionId: 'lineage-session' };
      await dispatchMessage({
        type: 'document:open-buffer',
        id: 'lineage-open',
        payload: {
          file: new NodeFile(['source'], 'lesson.pptx'),
          editorBinding: binding,
          trackPageLineage: true,
        },
      });
      await vi.waitFor(() => expect(deckwisePageLineage.context()?.pageTokens.length).toBe(2));
      expectMessagePosted(postMessageSpy, 'document:opened', 'lineage-open');
      const file = new NodeFile(['output'], 'lesson.pptx');
      mockRequestSaveDocument.mockResolvedValueOnce(file);
      await dispatchMessage({ type: 'document:save', id: 'lineage-save', payload: { requireHostCommit: true } });
      await vi.waitFor(() =>
        expectMessagePosted(postMessageSpy, 'document:saved', 'lineage-save', {
          file,
          pageLineage: {
            ...binding,
            protocol: 'deckwise-office-page-lineage/2',
            saveRequestId: 'lineage-save',
            sourceSha256: createHash('sha256').update('source').digest('hex'),
            outputSha256: createHash('sha256').update('output').digest('hex'),
          },
        }),
      );
      await dispatchMessage({ type: 'document:commit-save', id: 'lineage-save', payload: { success: false } });
      mockRequestSaveDocument.mockImplementationOnce(async () => {
        history.Points[0].Items.push({});
        return file;
      });
      await dispatchMessage({ type: 'document:save', id: 'lineage-race', payload: { requireHostCommit: true } });
      await vi.waitFor(() => expectMessagePosted(postMessageSpy, 'document:saved', 'lineage-race'));
      expect(postedPayloads(postMessageSpy, 'document:saved', 'lineage-race').at(-1)?.pageLineage).toBeUndefined();
    } finally {
      deckwisePageLineage.reset();
      document.querySelectorAll('iframe').forEach((frame) => frame.remove());
      mockHandleDocumentOperation.mockResolvedValue(undefined);
      vi.unstubAllGlobals();
    }
  });

  describe('host-managed save confirmation', () => {
    async function setup() {
      window.history.pushState({}, '', '/?embed=1&deckwise=1&embedOrigin=https%3A%2F%2Fparent.example.com');
      const iframe = document.createElement('iframe');
      document.body.appendChild(iframe);
      const history = { ForceSave: false, Index: 0, Points: [{ Items: [{}] }], Reset_SavedIndex: vi.fn() };
      Object.assign(iframe.contentWindow as object, {
        AscCommon: { History: history },
        Asc: { editor: { CheckChangedDocument: vi.fn() } },
      });
      const { setDocmentObj } = await import('@ranuts/shared/store');
      setDocmentObj({ fileName: 'lesson.pptx' });
      const { initEmbedApi } = await import('../../lib/embed-api');
      const guard = await import('../../lib/unsaved-guard');
      initEmbedApi();
      guard.markDocumentDirty();
      mockRequestSaveDocument.mockResolvedValue(new File(['pptx'], 'lesson.pptx'));
      return { history, guard, iframe };
    }
    const save = (id: string) => dispatchMessage({ type: 'document:save', id, payload: { requireHostCommit: true } });
    const ack = (id: string, success: boolean) =>
      dispatchMessage({ type: 'document:commit-save', id, payload: { success } });

    afterEach(() => document.querySelectorAll('iframe').forEach((frame) => frame.remove()));

    it.each(['document:get-context', 'document:execute', 'document:open-buffer', 'document:save'])(
      'blocks %s during a legacy DeckWise SDK export without creating a host save',
      async (type) => {
        const { guard } = await setup();
        const { deckwiseEditorBridge } = await import('../../lib/deckwise-editor-bridge');
        const context = vi.spyOn(deckwiseEditorBridge, 'getContext');
        const execute = vi.spyOn(deckwiseEditorBridge, 'execute');
        const reset = vi.spyOn(deckwiseEditorBridge, 'reset');
        let release!: () => void;
        mockRequestSaveDocument.mockImplementationOnce(() => {
          mockHasPendingEmbeddedSaveRequest.mockReturnValue(true);
          return new Promise<File>((resolve) => {
            release = () => {
              mockHasPendingEmbeddedSaveRequest.mockReturnValue(false);
              resolve(new File(['pptx'], 'lesson.pptx'));
            };
          });
        });
        try {
          await dispatchMessage({ type: 'document:save', id: 'legacy-export', payload: {} });
          await dispatchMessage({ type, id: 'during-export', payload: { requireHostCommit: true } });
          expectMessagePosted(postMessageSpy, 'document:error', 'during-export', {
            message:
              type === 'document:save'
                ? 'A save request is already in progress'
                : type === 'document:open-buffer'
                  ? 'Finish the current save before opening another document'
                  : 'Finish or recover the current save before using editor tools',
          });
          expect(context).not.toHaveBeenCalled();
          expect(execute).not.toHaveBeenCalled();
          expect(reset).not.toHaveBeenCalled();
          expect(mockHandleDocumentOperation).not.toHaveBeenCalled();
          expect(mockRequestSaveDocument).toHaveBeenCalledOnce();
          release();
          await vi.waitFor(() => expectMessagePosted(postMessageSpy, 'document:saved', 'legacy-export'));
          expect(guard.hasUnsavedChanges()).toBe(false);
          await save('after-legacy');
          expectMessagePosted(postMessageSpy, 'document:saved', 'after-legacy');
          await ack('after-legacy', true);
          expectMessagePosted(postMessageSpy, 'document:save-committed', 'after-legacy', { dirty: false });
        } finally {
          release?.();
          context.mockRestore();
          execute.mockRestore();
          reset.mockRestore();
        }
      },
    );

    it('automatically requests a Host save after edits settle, and schedules the next dirty edit', async () => {
      vi.useFakeTimers();
      try {
        window.history.pushState({}, '', '/?embed=1&deckwise=1');
        const { initEmbedApi } = await import('../../lib/embed-api');
        const guard = await import('../../lib/unsaved-guard');
        initEmbedApi();
        const count = () => postedMessages(postMessageSpy).filter(message => message.type === 'document:save-requested').length;
        const before = count();
        guard.markDocumentDirty();
        await vi.advanceTimersByTimeAsync(1000);
        guard.markDocumentDirty();
        await vi.advanceTimersByTimeAsync(1499);
        expect(count()).toBe(before);
        await vi.advanceTimersByTimeAsync(1);
        expect(count()).toBe(before + 1);
        await vi.advanceTimersByTimeAsync(5000);
        expect(count()).toBe(before + 1);
        guard.markDocumentDirty();
        await vi.advanceTimersByTimeAsync(1500);
        expect(count()).toBe(before + 2);
      } finally { vi.useRealTimers(); }
    });

    it('keeps exported bytes dirty until the matching host confirms persistence', async () => {
      const { guard, history } = await setup();
      await save('persist-1');
      expect(guard.hasUnsavedChanges()).toBe(true);
      expect(history.Reset_SavedIndex).not.toHaveBeenCalled();
      await ack('wrong-id', true);
      expect(guard.hasUnsavedChanges()).toBe(true);
      await ack('persist-1', true);
      expect(guard.hasUnsavedChanges()).toBe(false);
      expectMessagePosted(postMessageSpy, 'document:save-committed', 'persist-1', { success: true, dirty: false });
    });

    it('keeps edits made after export dirty, even inside the same undo point', async () => {
      const { guard, history } = await setup();
      await save('race');
      history.Points[0].Items.push({});
      await ack('race', true);
      expect(guard.hasUnsavedChanges()).toBe(true);
      expect(history.Reset_SavedIndex).not.toHaveBeenCalled();
      expectMessagePosted(postMessageSpy, 'document:save-committed', 'race', { dirty: true });
    });

    it('stays dirty when undo returns to the old SDK saved point before host confirmation', async () => {
      const { guard, history } = await setup();
      await save('undo-race');
      history.Index = -1;
      guard.markDocumentSaved();
      await ack('undo-race', true);
      expect(guard.hasUnsavedChanges()).toBe(true);
      expect(history.ForceSave).toBe(true);
    });

    it('allows retry after a failed disk write and ignores a late old acknowledgement', async () => {
      const { guard } = await setup();
      await save('failed');
      await ack('failed', false);
      expect(guard.hasUnsavedChanges()).toBe(true);
      await save('retry');
      await ack('failed', true);
      expect(guard.hasUnsavedChanges()).toBe(true);
      await ack('retry', true);
      expect(guard.hasUnsavedChanges()).toBe(false);
    });

    it('rejects another save until the host acknowledges the first export', async () => {
      await setup();
      const { deckwiseEditorBridge } = await import('../../lib/deckwise-editor-bridge');
      const reset = vi.spyOn(deckwiseEditorBridge, 'reset');
      await save('one');
      await dispatchMessage({ type: 'document:open-buffer', id: 'blocked-open', payload: {} });
      expect(reset).not.toHaveBeenCalled();
      await save('two');
      expect(mockRequestSaveDocument).toHaveBeenCalledOnce();
      expectMessagePosted(postMessageSpy, 'document:error', 'two', { message: 'A host save is still pending' });
      await ack('one', false);
    });

    it('blocks context and commands during host save and after an uncertain export', async () => {
      await setup();
      const { deckwiseEditorBridge } = await import('../../lib/deckwise-editor-bridge');
      const context = vi.spyOn(deckwiseEditorBridge, 'getContext');
      const execute = vi.spyOn(deckwiseEditorBridge, 'execute');
      const blocked = async (suffix: string) => {
        for (const type of ['document:get-context', 'document:execute']) {
          const id = type.endsWith('execute') ? `command-${suffix}` : `context-${suffix}`;
          await dispatchMessage({
            type,
            id,
            payload: { projectId: 'project', documentId: 'document', sessionId: 'session' },
          });
          expectMessagePosted(postMessageSpy, 'document:error', id, {
            message: 'Finish or recover the current save before using editor tools',
          });
        }
      };
      await save('tools-save');
      await blocked('pending');
      await ack('tools-save', false);
      mockRequestSaveDocument.mockRejectedValueOnce(new Error('Export timed out'));
      await save('tools-timeout');
      await blocked('uncertain');
      expect(context).not.toHaveBeenCalled();
      expect(execute).not.toHaveBeenCalled();
      context.mockRestore();
      execute.mockRestore();
    });

    it('ignores acknowledgement before export completion and from another source', async () => {
      const { guard } = await setup();
      let resolve!: (file: File) => void;
      mockRequestSaveDocument.mockReturnValueOnce(
        new Promise<File>((done) => {
          resolve = done;
        }),
      );
      await save('early');
      await ack('early', true);
      expect(guard.hasUnsavedChanges()).toBe(true);
      resolve(new File(['pptx'], 'lesson.pptx'));
      await new Promise((done) => setTimeout(done, 0));
      await dispatchMessage(
        { type: 'document:commit-save', id: 'early', payload: { success: true } },
        'https://parent.example.com',
        null,
      );
      expect(guard.hasUnsavedChanges()).toBe(true);
      await ack('early', true);
      expect(guard.hasUnsavedChanges()).toBe(false);
    });

    it('requires reopening after uncertain export failure so a late SDK stream cannot satisfy a retry', async () => {
      const { guard } = await setup();
      mockRequestSaveDocument.mockRejectedValueOnce(new Error('Export timed out'));
      await save('timeout');
      await save('unsafe-retry');
      expect(mockRequestSaveDocument).toHaveBeenCalledOnce();
      expect(guard.hasUnsavedChanges()).toBe(true);
      expectMessagePosted(postMessageSpy, 'document:error', 'unsafe-retry', {
        message: 'Export outcome is uncertain; reopen before saving again',
      });
    });
  });

  describe('makeFileFromPayload branches', () => {
    // makeFileFromPayload is internal; we exercise it via document:open-buffer messages.
    // mockHandleDocumentOperation lets us inspect the File that was constructed.

    async function openWithPayload(payload: Record<string, unknown>, id: string) {
      window.history.pushState({}, '', '/?embed=1');
      const { initEmbedApi } = await import('../../lib/embed-api');
      initEmbedApi();
      await dispatchMessage({ type: 'document:open-buffer', id, payload });
    }

    it('passes a File payload through unchanged', async () => {
      const file = new File([new Uint8Array([1, 2, 3])], 'original.xlsx', {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      });

      await openWithPayload({ file }, 'file-1');

      const [callArg] = mockHandleDocumentOperation.mock.calls.at(-1)!;
      expect(callArg.fileName).toBe('original.xlsx');
      expectMessagePosted(postMessageSpy, 'document:opened', 'file-1');
    });

    it('wraps a Blob payload into a File', async () => {
      const blob = new Blob([new Uint8Array([10, 20])], { type: 'text/csv' });

      await openWithPayload({ blob, fileName: 'data.csv' }, 'blob-1');

      const [callArg] = mockHandleDocumentOperation.mock.calls.at(-1)!;
      expect(callArg.fileName).toBe('data.csv');
      expectMessagePosted(postMessageSpy, 'document:opened', 'blob-1');
    });

    it('wraps an ArrayBuffer payload into a File', async () => {
      const buffer = new Uint8Array([7, 8, 9]).buffer;

      await openWithPayload({ buffer, fileName: 'report.docx' }, 'arraybuffer-1');

      const [callArg] = mockHandleDocumentOperation.mock.calls.at(-1)!;
      expect(callArg.fileName).toBe('report.docx');
      expectMessagePosted(postMessageSpy, 'document:opened', 'arraybuffer-1');
    });

    it('wraps a Uint8Array payload (via "bytes" key) into a File', async () => {
      const bytes = new Uint8Array([0xff, 0xfe]);

      await openWithPayload({ bytes, fileName: 'slide.pptx' }, 'uint8-1');

      const [callArg] = mockHandleDocumentOperation.mock.calls.at(-1)!;
      expect(callArg.fileName).toBe('slide.pptx');
      expectMessagePosted(postMessageSpy, 'document:opened', 'uint8-1');
    });

    it('decodes a base64 string payload (via "data" key) into a File', async () => {
      const original = new Uint8Array([0x50, 0x4b, 0x03, 0x04, 1, 2, 3]);
      const base64 = btoa(String.fromCharCode(...original));

      await openWithPayload({ data: base64, fileName: 'from-base64.docx' }, 'base64-1');

      const [callArg] = mockHandleDocumentOperation.mock.calls.at(-1)!;
      expect(callArg.fileName).toBe('from-base64.docx');
      const bytes = new Uint8Array(await callArg.file.arrayBuffer());
      expect(Array.from(bytes)).toEqual(Array.from(original));
      expectMessagePosted(postMessageSpy, 'document:opened', 'base64-1');
    });

    it('strips a data-URL prefix before decoding a base64 payload', async () => {
      const original = new Uint8Array([1, 2, 3, 4]);
      const base64 = btoa(String.fromCharCode(...original));

      await openWithPayload({ data: `data:application/octet-stream;base64,${base64}`, fileName: 'x.docx' }, 'base64-2');

      const [callArg] = mockHandleDocumentOperation.mock.calls.at(-1)!;
      const bytes = new Uint8Array(await callArg.file.arrayBuffer());
      expect(Array.from(bytes)).toEqual(Array.from(original));
    });

    it('uses default filename "document.xlsx" when no name is supplied', async () => {
      const buffer = new Uint8Array([1]).buffer;

      await openWithPayload({ buffer }, 'default-name-1');

      const [callArg] = mockHandleDocumentOperation.mock.calls.at(-1)!;
      expect(callArg.fileName).toBe('document.xlsx');
    });

    it('posts document:error for an empty payload (no file/blob/buffer/url)', async () => {
      await openWithPayload({}, 'invalid-1');

      expectMessagePosted(postMessageSpy, 'document:error', 'invalid-1', {
        message: expect.stringContaining('document:open requires'),
      });
    });
  });
});
