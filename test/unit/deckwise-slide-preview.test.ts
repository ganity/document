import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

const target = { context: { kind: 'slide' }, isCurrent: vi.fn(() => true) };
const current = vi.fn(() => ({
  projectId: 'project',
  documentId: 'document',
  sessionId: 'session',
  sourceSha256: 'a'.repeat(64),
  pageTokens: ['page-token'],
}));
const plainShape = () => ({
  getObjectType: () => 1,
  brush: { transparent: null as number | null },
  txWarpStruct: null as unknown,
});
const api = {
  getCurrentPage: vi.fn(() => 0),
  WordControl: {
    m_oLogicDocument: {
      Slides: [
        {
          draw: vi.fn(),
          cSld: { spTree: [plainShape()] },
          Layout: { cSld: { spTree: [plainShape()] }, Master: { cSld: { spTree: [plainShape()] } } },
        },
      ],
      GetWidthMM: vi.fn(() => 254),
      GetHeightMM: vi.fn(() => 158.75),
    },
  },
};
const initGraphics = vi.fn();
const transformGraphics = vi.fn();
class Graphics {
  init = initGraphics;
  transform = transformGraphics;
  isSupportTextDraw = () => true;
}
const renderer = {
  History: { Index: -1, Points: [] as unknown[] },
  CGraphics: Graphics,
  IsShapeToImageConverter: false,
  g_fontManager: { m_pFont: {} as unknown },
  g_fontManager2: { m_pFont: {} as unknown },
};
const encode = vi.fn();
const latest = vi.fn<() => { editGeneration: number }>(() => ({ editGeneration: 3 }));
const readOnly = vi.fn(() => true);
const history = renderer.History;
vi.mock('../../lib/agent-plugin/document-tool-action', () => ({
  captureDocumentToolTarget: () => target,
}));
vi.mock('../../lib/agent-plugin/editor-bridge', () => ({
  getEditorApi: () => api,
  requireEditorContext: () => ({
    AscCommon: renderer,
    AscDFH: { historyitem_type_Shape: 1, historyitem_type_ImageShape: 2, historyitem_type_GroupShape: 3 },
  }),
}));
vi.mock('../../lib/onlyoffice/readonly', () => ({ getReadonlyMode: () => readOnly() }));
vi.mock('../../lib/unsaved-guard', () => ({ hasUnsavedChanges: () => false }));
vi.mock('../../lib/deckwise-page-lineage', () => ({ deckwisePageLineage: { context: () => current() } }));
vi.mock('../../lib/deckwise-editor-bridge', () => ({ deckwiseEditorBridge: { getContext: latest } }));

const png = Uint8Array.from([
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0, 0, 0, 13, 0x49, 0x48, 0x44, 0x52, 0, 0, 0, 1, 0, 0, 0, 1, 8, 6, 0,
  0, 0, 0x1f, 0x15, 0xc4, 0x89,
]);
const owner = { projectId: 'project', documentId: 'document', sessionId: 'session' };

describe('DeckWise current slide preview', () => {
  it('renders another readonly page without navigating or changing the current page', async () => {
    const { renderDeckWiseCurrentSlidePreview } = await import('../../lib/deckwise-slide-preview');
    const second = { ...api.WordControl.m_oLogicDocument.Slides[0], draw: vi.fn() };
    const original = current();
    api.WordControl.m_oLogicDocument.Slides.push(second);
    current.mockReturnValue({ ...original, pageTokens: ['page-token', 'page-two'] });
    try {
      const result = renderDeckWiseCurrentSlidePreview(owner, 3, false, 2);
      expect(result.page).toBe(2);
      expect(result.pageToken).toBe('page-two');
      expect(second.draw).toHaveBeenCalledOnce();
      expect(api.getCurrentPage()).toBe(0);
    } finally {
      api.WordControl.m_oLogicDocument.Slides.pop();
      current.mockReturnValue(original);
    }
  });
  it('rejects invalid and absent requested pages without drawing', async () => {
    const { renderDeckWiseCurrentSlidePreview } = await import('../../lib/deckwise-slide-preview');
    for (const page of [0, -1, 1.5, Number.NaN, 2]) {
      expect(() => renderDeckWiseCurrentSlidePreview(owner, 3, false, page)).toThrow();
    }
    expect(api.WordControl.m_oLogicDocument.Slides[0].draw).not.toHaveBeenCalled();
  });
  it.each(['token', 'display', 'history'] as const)(
    'discards an offscreen preview if its %s changes while drawing',
    async (change) => {
      const { renderDeckWiseCurrentSlidePreview } = await import('../../lib/deckwise-slide-preview');
      const original = current();
      const second = { ...api.WordControl.m_oLogicDocument.Slides[0], draw: vi.fn() };
      api.WordControl.m_oLogicDocument.Slides.push(second);
      current.mockReturnValue({ ...original, pageTokens: ['page-token', 'page-two'] });
      second.draw.mockImplementationOnce(() => {
        if (change === 'token') current.mockReturnValue({ ...original, pageTokens: ['page-token', 'changed'] });
        if (change === 'display') api.getCurrentPage.mockReturnValue(1);
        if (change === 'history') history.Points.push({});
      });
      try {
        expect(() => renderDeckWiseCurrentSlidePreview(owner, 3, false, 2)).toThrow(/changed while rendering/);
        expect(renderer.IsShapeToImageConverter).toBe(false);
      } finally {
        api.WordControl.m_oLogicDocument.Slides.pop();
        current.mockReturnValue(original);
        api.getCurrentPage.mockReturnValue(0);
        history.Points.length = 0;
      }
    },
  );
  it('checks native text omission against the requested page rather than the displayed page', async () => {
    const { renderDeckWiseCurrentSlidePreview } = await import('../../lib/deckwise-slide-preview');
    const original = current();
    const second = {
      ...api.WordControl.m_oLogicDocument.Slides[0],
      cSld: { spTree: [{ ...plainShape(), txWarpStruct: {} }] },
      draw: vi.fn(),
    };
    api.WordControl.m_oLogicDocument.Slides.push(second);
    current.mockReturnValue({ ...original, pageTokens: ['page-token', 'page-two'] });
    try {
      expect(renderDeckWiseCurrentSlidePreview(owner, 3, true).page).toBe(1);
      expect(() => renderDeckWiseCurrentSlidePreview(owner, 3, true, 2)).toThrow(/ordinary opaque/);
      expect(second.draw).not.toHaveBeenCalled();
      expect(api.getCurrentPage()).toBe(0);
    } finally {
      api.WordControl.m_oLogicDocument.Slides.pop();
      current.mockReturnValue(original);
    }
  });
  beforeEach(() => {
    vi.spyOn(HTMLCanvasElement.prototype, 'getContext').mockReturnValue({} as CanvasRenderingContext2D);
    vi.spyOn(HTMLCanvasElement.prototype, 'toDataURL').mockImplementation(() => encode());
    encode.mockReturnValue(`data:image/png;base64,${Buffer.from(png).toString('base64')}`);
    api.WordControl.m_oLogicDocument.Slides[0].draw.mockReset();
    renderer.IsShapeToImageConverter = false;
  });
  afterEach(() => vi.restoreAllMocks());

  it('draws to a 1600px canvas at the exact page aspect ratio and restores renderer state', async () => {
    const { renderDeckWiseCurrentSlidePreview } = await import('../../lib/deckwise-slide-preview');
    api.WordControl.m_oLogicDocument.Slides[0].draw.mockImplementationOnce(() => {
      expect(renderer.IsShapeToImageConverter).toBe(true);
    });
    renderDeckWiseCurrentSlidePreview(owner, 3);
    expect(initGraphics).toHaveBeenLastCalledWith(expect.anything(), 1600, 1000, 254, 158.75);
    expect(transformGraphics).toHaveBeenLastCalledWith(1, 0, 0, 1, 0, 0);
    expect(renderer.IsShapeToImageConverter).toBe(false);
    expect(renderer.g_fontManager.m_pFont).toBe(null);
    expect(renderer.g_fontManager2.m_pFont).toBe(null);
  });

  it('omits native text only on the temporary graphics instance and labels the result', async () => {
    const { renderDeckWiseCurrentSlidePreview } = await import('../../lib/deckwise-slide-preview');
    api.WordControl.m_oLogicDocument.Slides[0].draw.mockImplementationOnce((graphics: Graphics) => {
      expect(graphics.isSupportTextDraw()).toBe(false);
    });
    expect(renderDeckWiseCurrentSlidePreview(owner, 3, true).nativeTextOmitted).toBe(true);
    api.WordControl.m_oLogicDocument.Slides[0].draw.mockImplementationOnce((graphics: Graphics) => {
      expect(graphics.isSupportTextDraw()).toBe(true);
    });
    expect(renderDeckWiseCurrentSlidePreview(owner, 3)).not.toHaveProperty('nativeTextOmitted');
    expect(new Graphics().isSupportTextDraw()).toBe(true);
  });

  it.each(['slide', 'layout', 'master'] as const)(
    'refuses unsupported %s content before drawing a textless preview',
    async (scope) => {
      const { renderDeckWiseCurrentSlidePreview } = await import('../../lib/deckwise-slide-preview');
      const slide = api.WordControl.m_oLogicDocument.Slides[0];
      const tree =
        scope === 'slide'
          ? slide.cSld.spTree
          : scope === 'layout'
            ? slide.Layout.cSld.spTree
            : slide.Layout.Master.cSld.spTree;
      for (const shape of [
        { ...plainShape(), getObjectType: () => 3 },
        { ...plainShape(), txWarpStruct: {} },
        { ...plainShape(), brush: { transparent: 128 } },
        { ...plainShape(), getOuterShdw: () => ({}) },
      ]) {
        tree.push(shape);
        try {
          expect(() => renderDeckWiseCurrentSlidePreview(owner, 3, true)).toThrow(/ordinary opaque/);
        } finally {
          tree.pop();
        }
      }
      expect(slide.draw).not.toHaveBeenCalled();
      expect(HTMLCanvasElement.prototype.getContext).not.toHaveBeenCalled();
    },
  );

  it.each(['slide', 'layout', 'master'] as const)(
    'omits text inside ordinary nested %s groups without changing their objects',
    async (scope) => {
      const { renderDeckWiseCurrentSlidePreview } = await import('../../lib/deckwise-slide-preview');
      const slide = api.WordControl.m_oLogicDocument.Slides[0];
      const tree =
        scope === 'slide'
          ? slide.cSld.spTree
          : scope === 'layout'
            ? slide.Layout.cSld.spTree
            : slide.Layout.Master.cSld.spTree;
      const leaf = plainShape();
      const group = {
        ...plainShape(),
        getObjectType: () => 3,
        spTree: [{ ...plainShape(), getObjectType: () => 3, spTree: [leaf] }],
      };
      tree.push(group);
      slide.draw.mockImplementationOnce((graphics: Graphics) => {
        expect(graphics.isSupportTextDraw()).toBe(false);
      });
      try {
        expect(renderDeckWiseCurrentSlidePreview(owner, 3, true).nativeTextOmitted).toBe(true);
        expect(group.spTree[0].spTree[0]).toBe(leaf);
        expect(new Graphics().isSupportTextDraw()).toBe(true);
      } finally {
        tree.pop();
      }
    },
  );

  it.each(['unknown', 'shadow', 'transparent', 'warp', 'cycle', 'deep', 'wide'] as const)(
    'rejects unsupported %s group content before drawing',
    async (kind) => {
      const { renderDeckWiseCurrentSlidePreview } = await import('../../lib/deckwise-slide-preview');
      type Shape = ReturnType<typeof plainShape> & { spTree?: Shape[]; getOuterShdw?(): unknown };
      const group: Shape = { ...plainShape(), getObjectType: () => 3, spTree: [plainShape()] };
      const child = group.spTree![0];
      if (kind === 'unknown') child.getObjectType = () => 9;
      if (kind === 'shadow') child.getOuterShdw = () => ({});
      if (kind === 'transparent') child.brush.transparent = 128;
      if (kind === 'warp') child.txWarpStruct = {};
      if (kind === 'cycle') group.spTree!.push(group);
      if (kind === 'wide') group.spTree = Array.from({ length: 1001 }, plainShape);
      if (kind === 'deep') {
        let next = group;
        for (let index = 0; index < 20; index++) {
          const nested: Shape = { ...plainShape(), getObjectType: () => 3, spTree: [plainShape()] };
          next.spTree = [nested];
          next = nested;
        }
      }
      const slide = api.WordControl.m_oLogicDocument.Slides[0];
      slide.cSld.spTree.push(group);
      try {
        expect(() => renderDeckWiseCurrentSlidePreview(owner, 3, true)).toThrow(/ordinary opaque/);
        expect(slide.draw).not.toHaveBeenCalled();
        expect(HTMLCanvasElement.prototype.getContext).not.toHaveBeenCalled();
      } finally {
        slide.cSld.spTree.pop();
      }
    },
  );

  it('rejects malformed native-text flags', async () => {
    const { renderDeckWiseCurrentSlidePreview } = await import('../../lib/deckwise-slide-preview');
    expect(() => renderDeckWiseCurrentSlidePreview(owner, 3, 'true' as never)).toThrow(/bound PPTX/);
  });

  it('keeps source freshness and read-only fences when native text is omitted', async () => {
    const { renderDeckWiseCurrentSlidePreview } = await import('../../lib/deckwise-slide-preview');
    latest.mockReturnValueOnce({ editGeneration: 4 });
    expect(() => renderDeckWiseCurrentSlidePreview(owner, 3, true)).toThrow(/bound PPTX/);
    readOnly.mockReturnValueOnce(false);
    expect(() => renderDeckWiseCurrentSlidePreview(owner, 3, true)).toThrow(/bound PPTX/);
    history.Points.push({});
    try {
      expect(() => renderDeckWiseCurrentSlidePreview(owner, 3, true)).toThrow(/bound PPTX/);
    } finally {
      history.Points.length = 0;
    }
    expect(HTMLCanvasElement.prototype.getContext).not.toHaveBeenCalled();
  });

  it('restores renderer state and clears the canvas if native drawing fails', async () => {
    const { renderDeckWiseCurrentSlidePreview } = await import('../../lib/deckwise-slide-preview');
    renderer.IsShapeToImageConverter = true;
    api.WordControl.m_oLogicDocument.Slides[0].draw.mockImplementationOnce(() => {
      throw new Error('draw failed');
    });
    expect(() => renderDeckWiseCurrentSlidePreview(owner, 3)).toThrow('draw failed');
    expect(renderer.IsShapeToImageConverter).toBe(true);
    expect(renderer.g_fontManager.m_pFont).toBe(null);
    const canvas = vi.mocked(HTMLCanvasElement.prototype.getContext).mock.contexts[0] as HTMLCanvasElement;
    expect(canvas.width).toBe(0);
    expect(canvas.height).toBe(0);
  });

  it('refuses invalid or extreme page dimensions before allocating a canvas', async () => {
    const { renderDeckWiseCurrentSlidePreview } = await import('../../lib/deckwise-slide-preview');
    api.WordControl.m_oLogicDocument.GetWidthMM.mockReturnValueOnce(Number.NaN);
    expect(() => renderDeckWiseCurrentSlidePreview(owner, 3)).toThrow(/unavailable/);
    api.WordControl.m_oLogicDocument.GetWidthMM.mockReturnValueOnce(0.0001);
    expect(() => renderDeckWiseCurrentSlidePreview(owner, 3)).toThrow(/aspect ratio/);
  });

  it('discards a render if its current page changes during native drawing', async () => {
    const { renderDeckWiseCurrentSlidePreview } = await import('../../lib/deckwise-slide-preview');
    api.getCurrentPage.mockReturnValueOnce(0).mockReturnValueOnce(1);
    expect(() => renderDeckWiseCurrentSlidePreview(owner, 3)).toThrow(/changed while rendering/);
    expect(renderer.IsShapeToImageConverter).toBe(false);
  });

  it('cleans up the renderer if PNG encoding fails', async () => {
    const { renderDeckWiseCurrentSlidePreview } = await import('../../lib/deckwise-slide-preview');
    encode.mockImplementationOnce(() => {
      throw new Error('PNG encoding failed');
    });
    expect(() => renderDeckWiseCurrentSlidePreview(owner, 3)).toThrow('PNG encoding failed');
    expect(renderer.IsShapeToImageConverter).toBe(false);
    const canvas = vi.mocked(HTMLCanvasElement.prototype.getContext).mock.contexts[0] as HTMLCanvasElement;
    expect(canvas.width).toBe(0);
    expect(canvas.height).toBe(0);
  });

  it('requires a fresh read-only copy, including after undoing all edits', async () => {
    const { renderDeckWiseCurrentSlidePreview } = await import('../../lib/deckwise-slide-preview');
    readOnly.mockReturnValueOnce(false);
    expect(() => renderDeckWiseCurrentSlidePreview(owner, 3)).toThrow(/bound PPTX/);
    history.Points.push({});
    expect(() => renderDeckWiseCurrentSlidePreview(owner, 3)).toThrow(/bound PPTX/);
    history.Points.length = 0;
  });
  it('requires the exact latest generation', async () => {
    const { renderDeckWiseCurrentSlidePreview } = await import('../../lib/deckwise-slide-preview');
    latest.mockReturnValueOnce({ editGeneration: 4 });
    expect(() => renderDeckWiseCurrentSlidePreview(owner, 3)).toThrow(/bound PPTX/);
  });

  it('returns bounded native PNG tied to the saved source and current page token', async () => {
    const { renderDeckWiseCurrentSlidePreview } = await import('../../lib/deckwise-slide-preview');
    encode.mockReturnValue(`data:image/png;base64,${Buffer.from(png).toString('base64')}`);
    expect(renderDeckWiseCurrentSlidePreview(owner, 3)).toEqual({
      sourceSha256: 'a'.repeat(64),
      pageToken: 'page-token',
      page: 1,
      mime: 'image/png',
      png,
    });
  });

  it('refuses stale targets, missing source lineage and non-PNG output', async () => {
    const { renderDeckWiseCurrentSlidePreview } = await import('../../lib/deckwise-slide-preview');
    target.isCurrent.mockReturnValueOnce(false);
    expect(() => renderDeckWiseCurrentSlidePreview(owner, 3)).toThrow(/bound PPTX/);
    current.mockImplementationOnce(() => undefined as never);
    expect(() => renderDeckWiseCurrentSlidePreview(owner, 3)).toThrow(/bound PPTX/);
    encode.mockReturnValueOnce('data:image/jpeg;base64,AA==');
    expect(() => renderDeckWiseCurrentSlidePreview(owner, 3)).toThrow(/PNG/);
  });
});
