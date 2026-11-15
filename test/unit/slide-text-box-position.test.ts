import { afterEach, beforeEach, expect, it, vi } from 'vitest';
import {
  canSetSlideTextBoxPosition,
  setSlideTextBoxPositionTool,
} from '../../lib/agent-plugin/slide-text-box-position';
import { canSetSlideTextBoxSize, setSlideTextBoxSizeTool } from '../../lib/agent-plugin/slide-text-box-size';
import { DeckWiseEditorBridge } from '../../lib/deckwise-editor-bridge';
import { setReadonlyState } from '../../lib/onlyoffice/readonly';
const owner = { projectId: 'project', documentId: 'doc', sessionId: 'session' };
const imageType = 72548352;
function imageNamespace(frame: HTMLIFrameElement) {
  return frame.contentWindow as Window & { AscDFH: Record<string, number> };
}
function mount() {
  const history = { Index: -1, Points: [] as Array<{ Items: unknown[] }> };
  const run = {
    Type: 39,
    Selection: { Use: true, StartPos: 0, EndPos: 5 },
    Content: [...'Hello'].map((c) => ({ Type: 1, GetCodePoint: () => c.codePointAt(0)! })),
  };
  const content = {
    Content: [{ Content: [run] }],
    GetText: () => 'Hello',
    GetSelectedText: () => 'Hello',
    GetSelectionState: () => [run.Selection],
  };
  const shape = {
    getDocContent: () => content,
    x: 20,
    y: 10,
    extX: 40,
    extY: 20,
    rot: 0,
    flipH: false,
    flipV: false,
    group: null,
    spPr: { geometry: { preset: 'rect' } },
    isPlaceholder: () => false,
    getObjectType: () => 72351744,
  };
  const slide = { cSld: { spTree: [shape] } };
  const controller = { selectedObjects: [shape], getTargetDocContent: () => content };
  const logic = {
    Slides: [slide],
    GetWidthMM: () => 320,
    GetHeightMM: () => 200,
    GetSelectedSlides: () => [0],
    GetSelectionState: () => [run.Selection],
    GetCurrentController: () => controller,
    IsMasterMode: () => false,
    CanEdit: () => true,
  };
  const api = {
    isDocumentLoadComplete: true,
    isLoadFullApi: true,
    isViewMode: false,
    canEdit: () => true,
    isLongAction: () => false,
    isSlideShow: () => false,
    getCurrentPage: () => 0,
    WordControl: { m_oLogicDocument: logic },
    ImgApply: vi.fn((value: { Position?: { X: number; Y: number }; Width?: number; Height?: number }) => {
      if (value.Position) {
        shape.x = value.Position.X;
        shape.y = value.Position.Y;
      }
      if (value.Width !== undefined) shape.extX = value.Width;
      if (value.Height !== undefined) shape.extY = value.Height;
      history.Index++;
      history.Points.push({ Items: [{}] });
    }),
  };
  const frame = document.createElement('iframe');
  document.body.append(frame);
  Object.assign(frame.contentWindow!, {
    editor: api,
    Asc: {},
    AscCommon: { History: history },
    AscDFH: { historyitem_type_Shape: 72351744 },
  });
  return { shape, api, logic, controller, content, run, history, frame };
}

beforeEach(() => setReadonlyState(false));
afterEach(() => {
  setReadonlyState(false);
  document.body.replaceChildren();
  vi.restoreAllMocks();
});
it('moves only the selected text box, preserves text and dimensions, and consumes its generation', async () => {
  const state = mount();
  const bridge = new DeckWiseEditorBridge();
  bridge.reset(owner);
  const ctx = bridge.getContext(owner);
  expect(ctx.capabilities.map((c) => c.name)).toContain('set_slide_text_box_position');
  const command = {
    ...owner,
    requestId: 'position-1',
    editGeneration: ctx.editGeneration,
    tool: 'set_slide_text_box_position',
    input: { xMm: 25, yMm: 35 },
  };
  await expect(bridge.execute(command, command.requestId)).resolves.toMatchObject({
    status: 'verified',
    result: { xMm: 25, yMm: 35, page: 1, verified: true },
  });
  expect(state.shape).toMatchObject({ x: 25, y: 35, extX: 40, extY: 20 });
  expect(state.content.GetText()).toBe('Hello');
  await expect(bridge.execute({ ...command, requestId: 'position-2' }, 'position-2')).rejects.toThrow(/changed/);
});
it.each([
  'readonly',
  'view',
  'master',
  'slideshow',
  'busy',
  'unloaded',
  'no-api',
  'no-type',
  'placeholder',
  'group',
  'rotation',
  'flip',
  'geometry',
  'multi',
  'foreign',
  'connector',
  'no-text',
  'no-size',
  'nonfinite',
  'offpage',
])('withholds and rejects %s before mutation', async (change) => {
  const state = mount();
  if (change === 'readonly') setReadonlyState(true);
  if (change === 'view') state.api.isViewMode = true;
  if (change === 'master') state.logic.IsMasterMode = () => true;
  if (change === 'slideshow') state.api.isSlideShow = () => true;
  if (change === 'busy') state.api.isLongAction = () => true;
  if (change === 'unloaded') state.api.isDocumentLoadComplete = false;
  if (change === 'no-api') Reflect.deleteProperty(state.api, 'ImgApply');
  if (change === 'no-type') state.shape.getObjectType = () => 0;
  if (change === 'placeholder') state.shape.isPlaceholder = () => true;
  if (change === 'group') Reflect.set(state.shape, 'group', {});
  if (change === 'rotation') state.shape.rot = 0.1;
  if (change === 'flip') state.shape.flipH = true;
  if (change === 'geometry') state.shape.spPr.geometry.preset = 'ellipse';
  if (change === 'multi') state.controller.selectedObjects.push(state.shape);
  if (change === 'connector') state.logic.Slides[0].cSld.spTree.push({ ...state.shape, getObjectType: () => 999 });
  if (change === 'foreign') state.logic.Slides[0].cSld.spTree = [];
  if (change === 'no-text') state.run.Selection.EndPos = 0;
  if (change === 'no-size') state.shape.extX = 0;
  if (change === 'nonfinite') state.shape.y = NaN;
  if (change === 'offpage') state.shape.x = 300;
  expect(canSetSlideTextBoxSize()).toBe(false);
  await expect(setSlideTextBoxSizeTool.execute({ widthMm: 80, heightMm: 30 })).rejects.toThrow();
  expect(canSetSlideTextBoxPosition()).toBe(false);
  await expect(setSlideTextBoxPositionTool.execute({ xMm: 25, yMm: 35 })).rejects.toThrow();
  if (change !== 'no-api') expect(state.api.ImgApply).not.toHaveBeenCalled();
});
it.each([
  { xMm: -1, yMm: 0 },
  { xMm: Infinity, yMm: 0 },
  { xMm: 0, yMm: NaN },
  { xMm: 1001, yMm: 0 },
  { xMm: 300, yMm: 0 },
  { xMm: 0, yMm: 190 },
])('rejects invalid or outside-page coordinates %s', async (input) => {
  const state = mount();
  await expect(setSlideTextBoxPositionTool.execute(input)).rejects.toThrow();
  expect(state.api.ImgApply).not.toHaveBeenCalled();
});
it.each(['wrong-position', 'resize', 'text', 'selection', 'page-size', 'cancel'])(
  'rejects uncertain %s after dispatch without retry',
  async (change) => {
    const state = mount();
    const abort = new AbortController();
    const apply = state.api.ImgApply.getMockImplementation()!;
    state.api.ImgApply.mockImplementation((value) => {
      apply(value);
      if (change === 'wrong-position') state.shape.x = 24;
      if (change === 'resize') state.shape.extX = 41;
      if (change === 'text') state.run.Content[0].GetCodePoint = () => 88;
      if (change === 'selection') state.run.Selection.EndPos = 2;
      if (change === 'page-size') state.logic.GetWidthMM = () => 321;
      if (change === 'cancel') abort.abort();
    });
    await expect(setSlideTextBoxPositionTool.execute({ xMm: 25, yMm: 35 }, abort.signal)).rejects.toThrow();
    expect(state.api.ImgApply).toHaveBeenCalledTimes(1);
  },
);
it('rejects foreign input at the bound bridge', async () => {
  const state = mount();
  const bridge = new DeckWiseEditorBridge();
  bridge.reset(owner);
  const ctx = bridge.getContext(owner);
  await expect(
    bridge.execute(
      {
        ...owner,
        requestId: 'bad',
        editGeneration: ctx.editGeneration,
        tool: 'set_slide_text_box_position',
        input: { xMm: 25, yMm: 35, script: 'bad' },
      },
      'bad',
    ),
  ).rejects.toThrow();
  expect(state.api.ImgApply).not.toHaveBeenCalled();
});

it.each(['ordinary', 'rotated-flipped', 'cropped'])(
  'moves the selected text box on a page with an untouched %s image',
  async (kind) => {
    const state = mount();
    Object.assign(imageNamespace(state.frame).AscDFH, { historyitem_type_ImageShape: imageType });
    const peer = { ...state.shape, getObjectType: () => imageType };
    if (kind === 'rotated-flipped') Object.assign(peer, { rot: Math.PI / 4, flipH: true, flipV: true });
    Reflect.set(peer, 'blipFill', {
      RasterImageId: 'local-peer.png',
      ...(kind === 'cropped' ? { srcRect: { l: 10, t: 20, r: 90, b: 80 } } : {}),
    });
    state.logic.Slides[0].cSld.spTree.push(peer);
    const geometry = {
      x: peer.x,
      y: peer.y,
      extX: peer.extX,
      extY: peer.extY,
      rot: peer.rot,
      flipH: peer.flipH,
      flipV: peer.flipV,
    };
    const fill = structuredClone(Reflect.get(peer, 'blipFill'));
    const bridge = new DeckWiseEditorBridge();
    bridge.reset(owner);
    const context = bridge.getContext(owner);
    expect(context.capabilities.map((c) => c.name)).toContain('set_slide_text_box_position');
    await expect(
      bridge.execute(
        {
          ...owner,
          requestId: 'mixed-position',
          editGeneration: context.editGeneration,
          tool: 'set_slide_text_box_position',
          input: { xMm: 25, yMm: 35 },
        },
        'mixed-position',
      ),
    ).resolves.toMatchObject({
      status: 'verified',
      result: { xMm: 25, yMm: 35, page: 1, verified: true },
    });
    expect(state.shape).toMatchObject({ x: 25, y: 35, extX: 40, extY: 20 });
    expect(peer).toMatchObject(geometry);
    expect(Reflect.get(peer, 'blipFill')).toEqual(fill);
    expect(state.api.ImgApply).toHaveBeenCalledTimes(1);
  },
);

it.each(['unknown-image-type', 'missing-image-type', 'grouped-image', 'placeholder-image'])(
  'withholds positioning with an unsupported %s peer before mutation',
  async (kind) => {
    const state = mount();
    if (kind !== 'missing-image-type')
      Object.assign(imageNamespace(state.frame).AscDFH, { historyitem_type_ImageShape: imageType });
    const peer = { ...state.shape, getObjectType: () => (kind === 'unknown-image-type' ? 999 : imageType) };
    if (kind === 'grouped-image') Reflect.set(peer, 'group', {});
    if (kind === 'placeholder-image') peer.isPlaceholder = () => true;
    state.logic.Slides[0].cSld.spTree.push(peer);
    expect(canSetSlideTextBoxSize()).toBe(false);
    await expect(setSlideTextBoxSizeTool.execute({ widthMm: 80, heightMm: 30 })).rejects.toThrow();
    expect(canSetSlideTextBoxPosition()).toBe(false);
    await expect(setSlideTextBoxPositionTool.execute({ xMm: 25, yMm: 35 })).rejects.toThrow();
    expect(state.api.ImgApply).not.toHaveBeenCalled();
  },
);

it('resizes only the selected text box at its original position through the bound bridge', async () => {
  const state = mount();
  const peer = { ...state.shape, x: 150, getObjectType: () => imageType };
  Object.assign(imageNamespace(state.frame).AscDFH, { historyitem_type_ImageShape: imageType });
  state.logic.Slides[0].cSld.spTree.push(peer);
  const bridge = new DeckWiseEditorBridge();
  bridge.reset(owner);
  const ctx = bridge.getContext(owner);
  expect(ctx.capabilities.map((c) => c.name)).toContain('set_slide_text_box_size');
  const command = {
    ...owner,
    requestId: 'size-1',
    editGeneration: ctx.editGeneration,
    tool: 'set_slide_text_box_size',
    input: { widthMm: 80.5, heightMm: 30 },
  };
  await expect(bridge.execute(command, command.requestId)).resolves.toMatchObject({
    status: 'verified',
    result: { widthMm: 80.5, heightMm: 30, page: 1, verified: true },
  });
  expect(state.shape).toMatchObject({ x: 20, y: 10, extX: 80.5, extY: 30 });
  expect(peer).toMatchObject({ x: 150, y: 10, extX: 40, extY: 20 });
  expect(state.content.GetText()).toBe('Hello');
  await expect(bridge.execute({ ...command, requestId: 'size-2' }, 'size-2')).rejects.toThrow(/changed/);
  expect(state.api.ImgApply).toHaveBeenCalledTimes(1);
});
it.each([
  { widthMm: 0, heightMm: 30 },
  { widthMm: -1, heightMm: 30 },
  { widthMm: 80, heightMm: NaN },
  { widthMm: Infinity, heightMm: 30 },
  { widthMm: 1001, heightMm: 30 },
  { widthMm: 301, heightMm: 30 },
  { widthMm: 80, heightMm: 191 },
])('rejects invalid or outside-page dimensions %s before dispatch', async (input) => {
  const state = mount();
  await expect(setSlideTextBoxSizeTool.execute(input)).rejects.toThrow();
  expect(state.api.ImgApply).not.toHaveBeenCalled();
});
it.each(['wrong-width', 'wrong-height', 'move', 'text', 'selection', 'page-size', 'cancel'])(
  'rejects uncertain resize %s after one dispatch',
  async (change) => {
    const state = mount();
    const abort = new AbortController();
    const apply = state.api.ImgApply.getMockImplementation()!;
    state.api.ImgApply.mockImplementation((value) => {
      apply(value);
      if (change === 'wrong-width') state.shape.extX = 79;
      if (change === 'wrong-height') state.shape.extY = 29;
      if (change === 'move') state.shape.x = 21;
      if (change === 'text') state.run.Content[0].GetCodePoint = () => 88;
      if (change === 'selection') state.run.Selection.EndPos = 2;
      if (change === 'page-size') state.logic.GetWidthMM = () => 321;
      if (change === 'cancel') abort.abort();
    });
    await expect(setSlideTextBoxSizeTool.execute({ widthMm: 80, heightMm: 30 }, abort.signal)).rejects.toThrow();
    expect(state.api.ImgApply).toHaveBeenCalledTimes(1);
  },
);
it('rejects foreign resize keys and cancellation before mutation', async () => {
  const state = mount();
  const bridge = new DeckWiseEditorBridge();
  bridge.reset(owner);
  const ctx = bridge.getContext(owner);
  await expect(
    bridge.execute(
      {
        ...owner,
        requestId: 'bad-size',
        editGeneration: ctx.editGeneration,
        tool: 'set_slide_text_box_size',
        input: { widthMm: 80, heightMm: 30, script: 'bad' },
      },
      'bad-size',
    ),
  ).rejects.toThrow();
  const abort = new AbortController();
  abort.abort();
  await expect(setSlideTextBoxSizeTool.execute({ widthMm: 80, heightMm: 30 }, abort.signal)).rejects.toThrow();
  expect(state.api.ImgApply).not.toHaveBeenCalled();
});
