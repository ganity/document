import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
vi.mock('../../lib/agent-plugin/slide-image-asset', () => ({
  replaceSlideImageFromAsset: vi.fn(),
  insertSlideImageFromAsset: vi.fn(),
}));
vi.mock('../../lib/agent-plugin/slide-image-replacement', () => ({ canReplaceSlideImage: vi.fn() }));
import { DeckWiseEditorBridge } from '../../lib/deckwise-editor-bridge';
import { agentTools } from '../../lib/agent-plugin/tools';
import { replaceSlideImageFromAsset, insertSlideImageFromAsset } from '../../lib/agent-plugin/slide-image-asset';
import { canReplaceSlideImage } from '../../lib/agent-plugin/slide-image-replacement';
import { canInsertSlideImage } from '../../lib/agent-plugin/slide-image-insertion';
vi.mock('../../lib/agent-plugin/slide-image-insertion', () => ({ canInsertSlideImage: vi.fn() }));
import { setReadonlyState } from '../../lib/onlyoffice/readonly';

const owner = { projectId: 'project', documentId: 'document', sessionId: 'session' };
const reviewedImage = {
  assetId: 'reviewed_png_1',
  sha256: 'a'.repeat(64),
  mime: 'image/png' as const,
  byteLength: 33,
  width: 1,
  height: 1,
};
const command = (editGeneration: number, requestId = 'request-1') => ({
  ...owner,
  editGeneration,
  requestId,
  tool: 'get_presentation_text',
  input: { maxChars: 8000 },
});

function mount() {
  const history = { Index: -1, Points: [] as Array<{ Items: unknown[] }> };
  let page = 0;
  const selectedObjects: unknown[] = [];
  const controller = { selectedObjects };
  const logic = {
    Slides: [{ cSld: { spTree: [{ getText: () => 'Hello lesson' }] } }],
    GetSelectedSlides: () => [page],
    GetCurrentController: () => controller,
  };
  const api = {
    isDocumentLoadComplete: true,
    isLoadFullApi: true,
    WordControl: { m_oLogicDocument: logic },
    getCurrentPage: () => page,
  };
  const frame = document.createElement('iframe');
  frame.name = 'frameEditor';
  document.body.append(frame);
  Object.assign(frame.contentWindow!, { editor: api, AscCommon: { History: history } });
  return { history, frame, selectedObjects, navigate: () => page++ };
}

describe('DeckWise native editor command bridge', () => {
  let bridge: DeckWiseEditorBridge;
  beforeEach(() => {
    setReadonlyState(false);
    vi.mocked(canReplaceSlideImage).mockReturnValue(false);
    vi.mocked(canInsertSlideImage).mockReturnValue(false);
    vi.mocked(insertSlideImageFromAsset).mockReset();
    vi.mocked(replaceSlideImageFromAsset).mockReset();
    bridge = new DeckWiseEditorBridge();
    bridge.reset(owner);
  });
  afterEach(() => {
    setReadonlyState(false);
    document.body.replaceChildren();
    vi.restoreAllMocks();
  });

  it.each(['replace', 'insert'] as const)(
    'reviews a PNG against a fresh writable %s target without editing',
    (operation) => {
      const state = mount();
      vi.mocked(canReplaceSlideImage).mockReturnValue(operation === 'replace');
      vi.mocked(canInsertSlideImage).mockReturnValue(operation === 'insert');
      const before = bridge.getContext(owner);
      const context = bridge.reviewImageTarget(owner, reviewedImage, {
        page: before.page,
        editGeneration: before.editGeneration,
        operation,
      });
      expect(context.page).toBe(before.page);
      expect(context.editGeneration).toBeGreaterThan(before.editGeneration);
      expect(context.capabilities.map((item) => item.name)).toContain(`${operation}_slide_image_from_asset`);
      expect(state.history).toEqual({ Index: -1, Points: [] });
      expect(replaceSlideImageFromAsset).not.toHaveBeenCalled();
      expect(insertSlideImageFromAsset).not.toHaveBeenCalled();
    },
  );

  it('rejects stale pages, generations and read-only targets before reviewing metadata', () => {
    const state = mount();
    vi.mocked(canReplaceSlideImage).mockReturnValue(true);
    const before = bridge.getContext(owner);
    const review = vi.spyOn(bridge, 'reviewImage');
    const target = { page: before.page, editGeneration: before.editGeneration, operation: 'replace' };
    expect(() => bridge.reviewImageTarget(owner, reviewedImage, { ...target, page: 2 })).toThrow(/fresh writable/);
    expect(() => bridge.reviewImageTarget(owner, reviewedImage, { ...target, editGeneration: 99 })).toThrow(
      /fresh writable/,
    );
    state.selectedObjects.push({});
    expect(() => bridge.reviewImageTarget(owner, reviewedImage, target)).toThrow(/fresh writable/);
    const current = bridge.getContext(owner);
    setReadonlyState(true);
    expect(() =>
      bridge.reviewImageTarget(owner, reviewedImage, { ...target, editGeneration: current.editGeneration }),
    ).toThrow(/fresh writable/);
    expect(review).not.toHaveBeenCalled();
  });

  it.each(['replace', 'insert'] as const)(
    'keeps an approved %s target current across Chat text reads without accepting a changed selection',
    async (operation) => {
      const state = mount();
      vi.mocked(canReplaceSlideImage).mockReturnValue(operation === 'replace');
      vi.mocked(canInsertSlideImage).mockReturnValue(operation === 'insert');
      const before = bridge.getContext(owner);
      const approved = bridge.reviewImageTarget(owner, reviewedImage, {
        page: before.page,
        editGeneration: before.editGeneration,
        operation,
      });
      for (const requestId of ['chat-read-1', 'chat-read-2']) {
        await bridge.execute(command(approved.editGeneration, requestId), requestId);
        const context = bridge.getContext(owner);
        expect(context.editGeneration).toBe(approved.editGeneration);
        expect(context.capabilities.map((item) => item.name)).toContain(`${operation}_slide_image_from_asset`);
      }
      state.selectedObjects.push({});
      expect(bridge.getContext(owner).editGeneration).toBeGreaterThan(approved.editGeneration);
      await expect(bridge.execute(command(approved.editGeneration, 'old-target'), 'old-target')).rejects.toThrow(
        /changed/,
      );
    },
  );

  it.each(['replace', 'insert'] as const)(
    'revokes reviewed metadata when the requested %s capability is unsupported',
    (operation) => {
      mount();
      vi.mocked(canReplaceSlideImage).mockReturnValue(operation !== 'replace');
      vi.mocked(canInsertSlideImage).mockReturnValue(operation !== 'insert');
      const before = bridge.getContext(owner);
      expect(() =>
        bridge.reviewImageTarget(owner, reviewedImage, {
          page: before.page,
          editGeneration: before.editGeneration,
          operation,
        }),
      ).toThrow(/supported image target/);
      vi.mocked(canReplaceSlideImage).mockReturnValue(true);
      vi.mocked(canInsertSlideImage).mockReturnValue(true);
      expect(bridge.getContext(owner).capabilities.map((item) => item.name)).not.toContain(
        'replace_slide_image_from_asset',
      );
      expect(bridge.getContext(owner).capabilities.map((item) => item.name)).not.toContain(
        'insert_slide_image_from_asset',
      );
    },
  );

  it('refuses a native selection change during the synchronous image review and clears its metadata', () => {
    const state = mount();
    vi.mocked(canReplaceSlideImage).mockReturnValue(true);
    const before = bridge.getContext(owner),
      reviewImage = bridge.reviewImage.bind(bridge);
    vi.spyOn(bridge, 'reviewImage').mockImplementationOnce((...args) => {
      state.selectedObjects.push({});
      return reviewImage(...args);
    });
    expect(() =>
      bridge.reviewImageTarget(owner, reviewedImage, {
        page: before.page,
        editGeneration: before.editGeneration,
        operation: 'replace',
      }),
    ).toThrow(/supported image target/);
    expect(bridge.getContext(owner).capabilities.map((item) => item.name)).not.toContain(
      'replace_slide_image_from_asset',
    );
  });

  it('rejects malformed target fields without changing the legacy reviewed image contract', () => {
    mount();
    vi.mocked(canReplaceSlideImage).mockReturnValue(true);
    const before = bridge.getContext(owner);
    const target = { page: before.page, editGeneration: before.editGeneration, operation: 'replace' };
    for (const invalid of [
      null,
      {},
      { ...target, extra: true },
      { ...target, page: '1' },
      { ...target, page: 0 },
      { ...target, editGeneration: 0 },
      { ...target, operation: 'other' },
      { ...target, operation: ['replace'] },
    ])
      expect(() => bridge.reviewImageTarget(owner, reviewedImage, invalid)).toThrow(/Invalid image target/);
    expect(bridge.reviewImage(owner, reviewedImage).capabilities.map((item) => item.name)).toContain(
      'replace_slide_image_from_asset',
    );
    expect(bridge.reviewImage(owner, null).capabilities.map((item) => item.name)).not.toContain(
      'replace_slide_image_from_asset',
    );
  });

  it('keeps reviewed metadata separate from fresh command bytes and clears both on reset', async () => {
    mount();
    vi.mocked(canReplaceSlideImage).mockReturnValue(true);
    const context = bridge.reviewImage(owner, reviewedImage);
    expect(context.capabilities.some((item) => item.name === 'replace_slide_image_from_asset')).toBe(true);
    const bytes = new Uint8Array(33);
    bytes[0] = 42;
    vi.mocked(replaceSlideImageFromAsset).mockImplementation(async (input) => {
      const loaded = await input.loadAsset(reviewedImage.assetId);
      expect(loaded[0]).toBe(42);
      loaded[0] = 7;
      expect((await input.loadAsset(reviewedImage.assetId))[0]).toBe(42);
      return {
        assetId: reviewedImage.assetId,
        sha256: reviewedImage.sha256,
        imageId: 'native',
        page: 1,
        verified: true,
        saved: false,
      };
    });
    const edit = {
      ...command(context.editGeneration, 'reviewed-edit'),
      tool: 'replace_slide_image_from_asset',
      input: { assetId: reviewedImage.assetId },
    };
    await expect(bridge.executeImage(edit, bytes, edit.requestId)).resolves.toMatchObject({
      result: { assetId: reviewedImage.assetId },
    });
    const next = bridge.getContext(owner);
    vi.mocked(replaceSlideImageFromAsset).mockImplementation(async (input) => {
      await input.loadAsset(reviewedImage.assetId);
      throw new Error('unreachable');
    });
    await expect(
      bridge.execute({ ...edit, requestId: 'no-bytes', editGeneration: next.editGeneration }, 'no-bytes'),
    ).rejects.toThrow(/Fresh authorized/);
    bridge.reset(owner);
    expect(bridge.getContext(owner).capabilities.some((item) => item.name === 'replace_slide_image_from_asset')).toBe(
      false,
    );
    await expect(bridge.executeImage(edit, bytes, edit.requestId)).rejects.toThrow(/Invalid/);
    expect(() => bridge.reviewImage({ ...owner, sessionId: 'other' }, reviewedImage)).toThrow(/expired/);
    expect(() => bridge.reviewImage(owner, { ...reviewedImage, url: 'https://invalid' })).toThrow(/Invalid/);
  });

  it('keeps one command byte lease while busy and revokes it on session reset', async () => {
    mount();
    vi.mocked(canReplaceSlideImage).mockReturnValue(true);
    const context = bridge.reviewImage(owner, reviewedImage);
    let release!: () => void;
    const gate = new Promise<void>((resolve) => {
      release = resolve;
    });
    vi.mocked(replaceSlideImageFromAsset).mockImplementation(async (input) => {
      await gate;
      expect(input.signal?.aborted).toBe(true);
      await expect(input.loadAsset(reviewedImage.assetId)).rejects.toThrow(/Fresh authorized/);
      throw new Error('session-reset-refused');
    });
    const edit = {
      ...command(context.editGeneration, 'pending-reviewed-image'),
      tool: 'replace_slide_image_from_asset',
      input: { assetId: reviewedImage.assetId },
    };
    const pending = bridge.executeImage(edit, new Uint8Array(33), edit.requestId);
    const rejected = expect(pending).rejects.toThrow('session-reset-refused');
    expect(() => bridge.reviewImage(owner, null)).toThrow(/progress/);
    await expect(
      bridge.executeImage({ ...edit, requestId: 'concurrent' }, new Uint8Array(33), 'concurrent'),
    ).rejects.toThrow(/Invalid/);
    bridge.reset({ ...owner, sessionId: 'next-session' });
    release();
    await rejected;
    expect(bridge.isExecuting).toBe(false);
    expect(
      bridge
        .getContext({ ...owner, sessionId: 'next-session' })
        .capabilities.some((item) => item.name === 'replace_slide_image_from_asset'),
    ).toBe(false);
  });

  it('captures native context and executes the existing bounded read tool without a planner', async () => {
    mount();
    const context = bridge.getContext(owner);
    expect(context).toMatchObject({ ...owner, kind: 'slide', page: 1, editGeneration: 1, selectedText: '' });
    expect(context.capabilities.map((tool) => tool.name)).toEqual(['get_presentation_text', 'add_slide_text']);
    expect(bridge.getContext(owner).editGeneration).toBe(1);
    const result = await bridge.execute(command(1), 'request-1');
    expect(result).toMatchObject({
      ...owner,
      requestId: 'request-1',
      editGeneration: 1,
      status: 'sent',
      tool: 'get_presentation_text',
      result: { text: 'Slide 1:\nHello lesson', totalSlides: 1 },
    });
    expect(bridge.getContext(owner).editGeneration).toBe(1);
    await expect(bridge.execute(command(1), 'request-1')).rejects.toThrow(/consumed/);
  });

  it('allows verified native text edits to advance history, then requires fresh context', async () => {
    const state = mount();
    bridge.getContext(owner);
    const execute = vi.spyOn(agentTools.add_slide_text, 'execute').mockImplementation(async () => {
      state.history.Points.push({ Items: [{}] });
      state.history.Index++;
      return { page: 1, verified: true };
    });
    const edit = { ...command(1), tool: 'add_slide_text', input: { text: 'Good!\n你好 😀' } };
    await expect(bridge.execute(edit, 'request-1')).resolves.toMatchObject({
      tool: 'add_slide_text',
      status: 'verified',
      result: { page: 1, verified: true },
    });
    expect(execute).toHaveBeenCalledWith(edit.input, expect.any(AbortSignal));
    await expect(bridge.execute({ ...edit, requestId: 'request-2' }, 'request-2')).rejects.toThrow(/changed/);
    expect(bridge.getContext(owner).editGeneration).toBe(2);
    await expect(bridge.execute({ ...edit, editGeneration: 2 }, 'request-1')).rejects.toThrow(/consumed/);
  });

  it('does not offer or execute writes in a read-only document', async () => {
    mount();
    setReadonlyState(true);
    const context = bridge.getContext(owner);
    expect(context.capabilities.map((tool) => tool.name)).toEqual(['get_presentation_text']);
    const execute = vi.spyOn(agentTools.add_slide_text, 'execute');
    await expect(
      bridge.execute({ ...command(1), tool: 'add_slide_text', input: { text: 'blocked' } }, 'request-1'),
    ).rejects.toThrow();
    expect(execute).not.toHaveBeenCalled();
  });

  it('advertises only current authorized image IDs and resolves bytes in the Office bridge', async () => {
    mount();
    const asset = reviewedImage;
    let assets = [asset];
    const load = vi.fn(async () => new Uint8Array(33));
    bridge.setImageAssetSource({ list: () => assets, load });
    expect(bridge.getContext(owner).capabilities.map((capability) => capability.name)).not.toContain(
      'replace_slide_image_from_asset',
    );

    vi.mocked(canReplaceSlideImage).mockReturnValue(true);
    const context = bridge.getContext(owner);
    const capability = context.capabilities.find((item) => item.name === 'replace_slide_image_from_asset');
    expect(capability?.inputSchema).toMatchObject({ properties: { assetId: { enum: ['reviewed_png_1'] } } });
    assets = [{ ...asset, assetId: 'reviewed_png_2' }];
    const refreshed = bridge.getContext(owner);
    expect(refreshed.editGeneration).toBeGreaterThan(context.editGeneration);
    expect(
      refreshed.capabilities.find((item) => item.name === 'replace_slide_image_from_asset')?.inputSchema,
    ).toMatchObject({ properties: { assetId: { enum: ['reviewed_png_2'] } } });

    vi.mocked(replaceSlideImageFromAsset).mockImplementation(async (input) => {
      expect(input.readCurrent()).toMatchObject({
        ...owner,
        editGeneration: requestGeneration,
        imageReplacementInstalled: true,
      });
      await input.loadAsset('reviewed_png_2', input.signal);
      return {
        assetId: input.asset.assetId,
        sha256: input.asset.sha256,
        imageId: 'office-private-id',
        page: 1,
        verified: true,
        saved: false,
      };
    });
    const request = {
      ...owner,
      requestId: 'image-1',
      editGeneration: refreshed.editGeneration,
      tool: 'replace_slide_image_from_asset',
      input: { assetId: 'reviewed_png_2' },
    };
    await expect(
      bridge.execute({ ...request, requestId: 'image-bad', input: { assetId: 'reviewed_png_1' } }, 'image-bad'),
    ).rejects.toThrow(/not authorized/);
    expect(load).not.toHaveBeenCalled();
    const next = bridge.getContext(owner);
    const requestGeneration = next.editGeneration;
    await expect(bridge.execute({ ...request, editGeneration: next.editGeneration }, 'image-1')).resolves.toMatchObject(
      {
        tool: 'replace_slide_image_from_asset',
        status: 'verified',
        result: { assetId: 'reviewed_png_2', page: 1, verified: true },
      },
    );
    expect(load).toHaveBeenCalledWith('reviewed_png_2', expect.any(AbortSignal));
  });

  it.each(['revoked', 'metadata', 'source', 'session', 'readonly'])(
    'invalidates a pending image command after %s changes',
    async (change) => {
      mount();
      vi.mocked(canReplaceSlideImage).mockReturnValue(true);
      let assets = [reviewedImage];
      let finish!: () => void;
      const load = vi.fn(
        () =>
          new Promise<Uint8Array>((resolve) => {
            finish = () => resolve(new Uint8Array(33));
          }),
      );
      bridge.setImageAssetSource({ list: () => assets, load });
      const context = bridge.getContext(owner);
      vi.mocked(replaceSlideImageFromAsset).mockImplementation(async (input) => {
        await input.loadAsset(input.asset.assetId, input.signal);
        input.signal?.throwIfAborted();
        if (input.readCurrent().imageReplacementInstalled !== true || input.readCurrent().readOnly)
          throw new Error('Image authorization or context changed');
        throw new Error('The invalidated command must not apply');
      });
      const pending = bridge.execute(
        {
          ...owner,
          requestId: 'image-pending',
          editGeneration: context.editGeneration,
          tool: 'replace_slide_image_from_asset',
          input: { assetId: reviewedImage.assetId },
        },
        'image-pending',
      );
      expect(bridge.isExecuting).toBe(true);
      if (change === 'revoked') assets = [];
      if (change === 'metadata') assets = [{ ...reviewedImage, sha256: 'b'.repeat(64) }];
      if (change === 'source') bridge.setImageAssetSource();
      if (change === 'session') bridge.reset({ ...owner, sessionId: 'new' });
      if (change === 'readonly') setReadonlyState(true);
      finish();
      await expect(pending).rejects.toThrow(/changed|expired/);
      expect(bridge.isExecuting).toBe(false);
    },
  );

  it.each(
    [
      [{ ...reviewedImage, assetId: '../bad' }],
      [{ ...reviewedImage, sha256: 'bad' }],
      [{ ...reviewedImage, width: 4097 }],
      [{ ...reviewedImage, byteLength: 0 }],
      [{ ...reviewedImage, url: 'https://example.invalid' }],
      [reviewedImage, reviewedImage],
      Array.from({ length: 101 }, (_, index) => ({ ...reviewedImage, assetId: `asset_${index}` })),
    ].map((assets) => ({ assets })),
  )('rejects malformed authorized asset metadata %#', ({ assets }) => {
    mount();
    vi.mocked(canReplaceSlideImage).mockReturnValue(true);
    bridge.setImageAssetSource({ list: () => assets, load: vi.fn() });
    expect(() => bridge.getContext(owner)).toThrow(/authorized image asset/);
  });

  it('clears image authorization on document reopen and rejects a model-supplied asset payload', async () => {
    mount();
    vi.mocked(canReplaceSlideImage).mockReturnValue(true);
    const load = vi.fn();
    bridge.setImageAssetSource({ list: () => [reviewedImage], load });
    const context = bridge.getContext(owner);
    await expect(
      bridge.execute(
        {
          ...owner,
          requestId: 'image-payload',
          editGeneration: context.editGeneration,
          tool: 'replace_slide_image_from_asset',
          input: { assetId: reviewedImage.assetId },
          hostAsset: reviewedImage,
        },
        'image-payload',
      ),
    ).rejects.toThrow(/Invalid/);
    bridge.reset(owner);
    expect(bridge.getContext(owner).capabilities.map((item) => item.name)).not.toContain(
      'replace_slide_image_from_asset',
    );
    expect(load).not.toHaveBeenCalled();
  });

  it('rejects an unverified write instead of reporting success', async () => {
    mount();
    bridge.getContext(owner);
    vi.spyOn(agentTools.add_slide_text, 'execute').mockResolvedValue({ page: 1 });
    await expect(
      bridge.execute({ ...command(1), tool: 'add_slide_text', input: { text: 'check me' } }, 'request-1'),
    ).rejects.toThrow(/could not be verified/);
  });

  it('aborts a pending native write when its document session is reset', async () => {
    mount();
    bridge.getContext(owner);
    let finish!: () => void;
    let signal!: AbortSignal;
    vi.spyOn(agentTools.add_slide_text, 'execute').mockImplementation(async (_input, abort) => {
      signal = abort!;
      await new Promise<void>((resolve) => {
        finish = resolve;
      });
      return { page: 1, verified: true };
    });
    const pending = bridge.execute(
      { ...command(1), tool: 'add_slide_text', input: { text: 'old document' } },
      'request-1',
    );
    bridge.reset({ ...owner, sessionId: 'new' });
    expect(signal.aborted).toBe(true);
    finish();
    await expect(pending).rejects.toThrow(/expired/);
  });

  it.each(['page', 'selection', 'edit', 'interaction', 'frame'] as const)(
    'rejects a stale %s snapshot before invoking a tool',
    async (change) => {
      const state = mount();
      const context = bridge.getContext(owner);
      const execute = vi.spyOn(agentTools.get_presentation_text, 'execute');
      if (change === 'page') state.navigate();
      if (change === 'selection') state.selectedObjects.push({});
      if (change === 'edit') {
        state.history.Points.push({ Items: [{}] });
        state.history.Index++;
      }
      if (change === 'interaction') state.frame.contentDocument!.dispatchEvent(new KeyboardEvent('keydown'));
      if (change === 'frame') state.frame.remove();
      await expect(bridge.execute(command(context.editGeneration), 'request-1')).rejects.toThrow(/changed/);
      expect(execute).not.toHaveBeenCalled();
      if (change !== 'frame') expect(bridge.getContext(owner).editGeneration).toBe(2);
    },
  );

  it('rejects old sessions, mismatched request IDs, unknown tools and invalid arguments', async () => {
    mount();
    bridge.getContext(owner);
    const execute = vi.spyOn(agentTools.get_presentation_text, 'execute');
    await expect(bridge.execute({ ...command(1), sessionId: 'old' }, 'request-1')).rejects.toThrow(/expired/);
    await expect(bridge.execute(command(1), 'other-id')).rejects.toThrow(/Invalid/);
    await expect(
      bridge.execute({ ...command(1), tool: 'execute_js', input: { script: 'alert(1)' } }, 'request-1'),
    ).rejects.toThrow(/not enabled/);
    await expect(bridge.execute({ ...command(1), input: { maxChars: 8001 } }, 'request-1')).rejects.toThrow();
    await expect(bridge.execute({ ...command(1), input: { script: 'alert(1)' } }, 'request-1')).rejects.toThrow();
    expect(execute).not.toHaveBeenCalled();
    bridge.reset({ ...owner, sessionId: 'new' });
    await expect(bridge.execute(command(1), 'request-1')).rejects.toThrow(/expired/);
  });

  it('serializes commands, rejects changing context during reads and consumes failed requests', async () => {
    const state = mount();
    bridge.getContext(owner);
    let finish!: () => void;
    vi.spyOn(agentTools.get_presentation_text, 'execute').mockImplementation(async () => {
      await new Promise<void>((resolve) => {
        finish = resolve;
      });
      return { text: 'old', totalSlides: 1, slidesRead: 1, truncated: false };
    });
    const pending = bridge.execute(command(1), 'request-1');
    expect(() => bridge.getContext(owner)).toThrow(/in progress/);
    await expect(bridge.execute(command(1, 'request-2'), 'request-2')).rejects.toThrow(/in progress/);
    state.navigate();
    finish();
    await expect(pending).rejects.toThrow(/changed while reading/);
    const context = bridge.getContext(owner);
    await expect(bridge.execute(command(context.editGeneration), 'request-1')).rejects.toThrow(/consumed/);
  });

  it('does not deliver an old command result into a newly opened document', async () => {
    mount();
    bridge.getContext(owner);
    let finish!: () => void;
    vi.spyOn(agentTools.get_presentation_text, 'execute').mockImplementation(async () => {
      await new Promise<void>((resolve) => {
        finish = resolve;
      });
      return {};
    });
    const pending = bridge.execute(command(1), 'request-1');
    bridge.reset({ ...owner, sessionId: 'new' });
    finish();
    await expect(pending).rejects.toThrow(/expired/);
  });
  it('advertises insertion only for current reviewed assets and an empty native selection', async () => {
    mount();
    vi.mocked(canInsertSlideImage).mockReturnValue(true);
    const context = bridge.reviewImage(owner, reviewedImage);
    expect(context.capabilities.map((c) => c.name)).toContain('insert_slide_image_from_asset');
    expect(context.capabilities.map((c) => c.name)).not.toContain('replace_slide_image_from_asset');
    const input = { assetId: reviewedImage.assetId, xMm: 25, yMm: 35, widthMm: 40, heightMm: 30 };
    vi.mocked(insertSlideImageFromAsset).mockImplementation(async (s) => {
      expect(s.readCurrent().imageInsertionInstalled).toBe(true);
      expect(await s.loadAsset(s.asset.assetId)).toEqual(new Uint8Array(33));
      return {
        assetId: s.asset.assetId,
        sha256: s.asset.sha256,
        imageId: 'inserted',
        page: 1,
        verified: true,
        saved: false,
        ...s.box,
      };
    });
    const command = {
      ...owner,
      editGeneration: context.editGeneration,
      requestId: 'insert-1',
      tool: 'insert_slide_image_from_asset',
      input,
    };
    const result = await bridge.executeImage(command, new Uint8Array(33), 'insert-1');
    expect(result).toMatchObject({
      tool: 'insert_slide_image_from_asset',
      result: { ...input, page: 1, verified: true },
    });
    expect(bridge.isExecuting).toBe(false);
    const next = bridge.getContext(owner);
    await expect(
      bridge.execute({ ...command, editGeneration: next.editGeneration, requestId: 'insert-2' }, 'insert-2'),
    ).rejects.toThrow(/Fresh authorized image bytes/);
  });
  it.each(['no-capability', 'extra', 'missing', 'nonfinite', 'negative', 'size', 'asset'])(
    'rejects forged insertion %s before dispatch',
    async (change) => {
      mount();
      vi.mocked(canInsertSlideImage).mockReturnValue(change !== 'no-capability');
      vi.mocked(canReplaceSlideImage).mockReturnValue(true);
      const context = bridge.reviewImage(owner, reviewedImage);
      const input: Record<string, unknown> = {
        assetId: reviewedImage.assetId,
        xMm: 25,
        yMm: 35,
        widthMm: 40,
        heightMm: 30,
      };
      if (change === 'extra') input.url = 'bad';
      if (change === 'missing') delete input.yMm;
      if (change === 'nonfinite') input.widthMm = NaN;
      if (change === 'negative') input.xMm = -1;
      if (change === 'size') input.heightMm = 0;
      if (change === 'asset') input.assetId = 'other';
      await expect(
        bridge.executeImage(
          {
            ...owner,
            editGeneration: context.editGeneration,
            requestId: 'bad-insert',
            tool: 'insert_slide_image_from_asset',
            input,
          },
          new Uint8Array(33),
          'bad-insert',
        ),
      ).rejects.toThrow();
      expect(insertSlideImageFromAsset).not.toHaveBeenCalled();
    },
  );
  it('captures insertion command identity and geometry across asset I/O', async () => {
    mount();
    vi.mocked(canInsertSlideImage).mockReturnValue(true);
    const context = bridge.reviewImage(owner, reviewedImage);
    let finish!: () => void;
    vi.mocked(insertSlideImageFromAsset).mockImplementation(async (s) => {
      await new Promise<void>((r) => {
        finish = r;
      });
      return {
        assetId: s.asset.assetId,
        sha256: s.asset.sha256,
        imageId: 'inserted',
        page: 1,
        verified: true,
        saved: false,
        ...s.box,
      };
    });
    const command = {
      ...owner,
      editGeneration: context.editGeneration,
      requestId: 'pinned-insert',
      tool: 'insert_slide_image_from_asset',
      input: { assetId: reviewedImage.assetId, xMm: 25, yMm: 35, widthMm: 40, heightMm: 30 },
    };
    const pending = bridge.executeImage(command, new Uint8Array(33), 'pinned-insert');
    command.requestId = 'changed';
    command.editGeneration = 100;
    command.input.xMm = 100;
    finish();
    await expect(pending).resolves.toMatchObject({
      requestId: 'pinned-insert',
      editGeneration: context.editGeneration,
      result: { xMm: 25 },
    });
  });
});
