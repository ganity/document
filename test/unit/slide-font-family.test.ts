import { afterEach, beforeEach, expect, it, vi } from 'vitest';
import { getSlideFontFamilies, setSlideFontFamilyTool } from '../../lib/agent-plugin/slide-font-family';
import { DeckWiseEditorBridge } from '../../lib/deckwise-editor-bridge';
import { setReadonlyState } from '../../lib/onlyoffice/readonly';

const owner = { projectId: 'project', documentId: 'doc', sessionId: 'session' };
function mount() {
  let family = 'Arial';
  const loader = {
    currentInfoLoaded: null as unknown,
    fontInfos: [
      { Name: 'Arial', indexR: 0, indexI: 0, indexB: 0, indexBI: 0 },
      { Name: 'Courier New', indexR: 1, indexI: 1, indexB: 1, indexBI: 1 },
    ],
    fontFiles: [{ CheckLoaded: () => true }, { CheckLoaded: () => true }],
  };
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
  const shape = { getDocContent: () => content };
  const slide = { cSld: { spTree: [shape] } };
  const controller = { selectedObjects: [shape], getTargetDocContent: () => content };
  const logic = {
    Slides: [slide],
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
    get_TextProps: () => ({ get_TextPr: () => ({ get_FontFamily: () => ({ get_Name: () => family }) }) }),
    put_TextPrFontName: vi.fn((value: string) => {
      family = value;
      history.Index++;
      history.Points.push({ Items: [{}] });
    }),
  };
  const frame = document.createElement('iframe');
  document.body.append(frame);
  Object.assign(frame.contentWindow!, { editor: api, Asc: {}, AscCommon: { History: history, g_font_loader: loader } });
  return { api, logic, controller, content, run, history, loader, family: () => family };
}
beforeEach(() => setReadonlyState(false));
afterEach(() => {
  setReadonlyState(false);
  document.body.replaceChildren();
  vi.restoreAllMocks();
});

it('sets one selected text range natively, verifies family/text and advances the bound generation', async () => {
  const state = mount();
  const bridge = new DeckWiseEditorBridge();
  bridge.reset(owner);
  const ctx = bridge.getContext(owner);
  expect(ctx.capabilities.map((c) => c.name)).toContain('set_slide_font_family');
  const command = {
    ...owner,
    requestId: 'font-1',
    editGeneration: ctx.editGeneration,
    tool: 'set_slide_font_family',
    input: { fontFamily: 'Courier New' },
  };
  await expect(bridge.execute(command, command.requestId)).resolves.toMatchObject({
    status: 'verified',
    result: { fontFamily: 'Courier New', page: 1, verified: true },
  });
  expect(state.family()).toBe('Courier New');
  expect(state.content.GetText()).toBe('Hello');
  await expect(bridge.execute({ ...command, requestId: 'font-2' }, 'font-2')).rejects.toThrow(/changed/);
});

it.each([
  'readonly',
  'view',
  'master',
  'loading',
  'busy',
  'slideshow',
  'no-selection',
  'multiple',
  'foreign-shape',
  'missing-setter',
  'missing-reader',
])('withholds the capability and rejects %s without calling a native setter', async (change) => {
  const state = mount();
  const setter = state.api.put_TextPrFontName;
  if (change === 'readonly') setReadonlyState(true);
  if (change === 'view') state.api.isViewMode = true;
  if (change === 'master') state.logic.IsMasterMode = () => true;
  if (change === 'loading') state.api.isLoadFullApi = false;
  if (change === 'busy') state.api.isLongAction = () => true;
  if (change === 'slideshow') state.api.isSlideShow = () => true;
  if (change === 'no-selection') state.run.Selection.Use = false;
  if (change === 'multiple') state.controller.selectedObjects.push(state.controller.selectedObjects[0]);
  if (change === 'foreign-shape') state.logic.Slides[0].cSld.spTree = [];
  if (change === 'missing-setter') Reflect.deleteProperty(state.api, 'put_TextPrFontName');
  if (change === 'missing-reader') Reflect.deleteProperty(state.api, 'get_TextProps');
  expect(getSlideFontFamilies()).toEqual([]);
  await expect(setSlideFontFamilyTool.execute({ fontFamily: 'Courier New' })).rejects.toThrow();
  expect(setter).not.toHaveBeenCalled();
});

it('rejects unloaded, unknown or malformed fonts and cancellation before writing', async () => {
  const state = mount();
  for (const fontFamily of ['', 'Unknown', ' Courier New', 'Courier New\n', 'x'.repeat(129)])
    await expect(setSlideFontFamilyTool.execute({ fontFamily })).rejects.toThrow();
  state.loader.fontFiles[1].CheckLoaded = () => false;
  expect(getSlideFontFamilies()).toEqual(['Arial']);
  await expect(setSlideFontFamilyTool.execute({ fontFamily: 'Courier New' })).rejects.toThrow(/loaded-font/);
  state.loader.currentInfoLoaded = {};
  expect(getSlideFontFamilies()).toEqual([]);
  await expect(setSlideFontFamilyTool.execute({ fontFamily: 'Arial' })).rejects.toThrow(/idle/);
  await expect(setSlideFontFamilyTool.execute({ fontFamily: 'Arial' }, AbortSignal.abort())).rejects.toThrow();
  expect(state.api.put_TextPrFontName).not.toHaveBeenCalled();
});

it.each(['wrong-family', 'changed-selection', 'changed-text', 'cancel'])(
  'rejects %s after dispatch without retrying or reporting success',
  async (change) => {
    const state = mount();
    const abort = new AbortController();
    state.api.put_TextPrFontName.mockImplementation(() => {
      if (change === 'changed-selection') state.run.Selection.EndPos = 2;
      if (change === 'changed-text') state.run.Content[0].GetCodePoint = () => 88;
      if (change === 'cancel') abort.abort();
    });
    await expect(setSlideFontFamilyTool.execute({ fontFamily: 'Courier New' }, abort.signal)).rejects.toThrow();
    expect(state.api.put_TextPrFontName).toHaveBeenCalledTimes(1);
  },
);

it('rejects removed capability and foreign command fields at the bridge', async () => {
  const state = mount();
  const bridge = new DeckWiseEditorBridge();
  bridge.reset(owner);
  const context = bridge.getContext(owner);
  const command = {
    ...owner,
    requestId: 'font-1',
    editGeneration: context.editGeneration,
    tool: 'set_slide_font_family',
    input: { fontFamily: 'Courier New', script: 'bad' },
  };
  await expect(bridge.execute(command, command.requestId)).rejects.toThrow();
  expect(state.api.put_TextPrFontName).not.toHaveBeenCalled();
});

it.each(['indexR', 'indexI', 'indexB', 'indexBI'] as const)(
  'withholds a family with an unloaded %s face, without scheduling a font load',
  async (face) => {
    const state = mount();
    state.loader.fontInfos[1][face] = 2;
    state.loader.fontFiles.push({ CheckLoaded: () => false });
    expect(getSlideFontFamilies()).toEqual(['Arial']);
    await expect(setSlideFontFamilyTool.execute({ fontFamily: 'Courier New' })).rejects.toThrow(/loaded-font/);
    expect(state.api.put_TextPrFontName).not.toHaveBeenCalled();
  },
);

it('rejects a family removed after capability discovery without touching the selection', async () => {
  const state = mount();
  const bridge = new DeckWiseEditorBridge();
  bridge.reset(owner);
  const context = bridge.getContext(owner);
  expect(context.capabilities.find((c) => c.name === 'set_slide_font_family')?.inputSchema).toMatchObject({
    properties: { fontFamily: { enum: ['Arial', 'Courier New'] } },
  });
  state.loader.fontFiles[1].CheckLoaded = () => false;
  const command = {
    ...owner,
    requestId: 'removed-font',
    editGeneration: context.editGeneration,
    tool: 'set_slide_font_family',
    input: { fontFamily: 'Courier New' },
  };
  await expect(bridge.execute(command, command.requestId)).rejects.toThrow();
  expect(state.api.put_TextPrFontName).not.toHaveBeenCalled();
  expect(state.content.GetSelectedText()).toBe('Hello');
});
