import { afterEach, beforeEach, expect, it, vi } from 'vitest';
import { canSetSlideTextBold, setSlideTextBoldTool } from '../../lib/agent-plugin/slide-text-bold';
import { DeckWiseEditorBridge } from '../../lib/deckwise-editor-bridge';
import { setReadonlyState } from '../../lib/onlyoffice/readonly';

const owner = { projectId: 'project', documentId: 'doc', sessionId: 'session' };
function mount() {
  let size = false;
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
    get_TextProps: () => ({ get_TextPr: () => ({ get_Bold: () => size }) }),
    put_TextPrBold: vi.fn((value: boolean) => {
      size = value;
      history.Index++;
      history.Points.push({ Items: [{}] });
    }),
  };
  const frame = document.createElement('iframe');
  document.body.append(frame);
  Object.assign(frame.contentWindow!, { editor: api, AscCommon: { History: history } });
  return { api, logic, controller, content, run, history, size: () => size };
}
beforeEach(() => setReadonlyState(false));
afterEach(() => {
  setReadonlyState(false);
  document.body.replaceChildren();
  vi.restoreAllMocks();
});

it('sets one selected text range natively, verifies bold/text and advances the bound generation', async () => {
  const state = mount();
  const bridge = new DeckWiseEditorBridge();
  bridge.reset(owner);
  const ctx = bridge.getContext(owner);
  expect(ctx.capabilities.map((c) => c.name)).toContain('set_slide_text_bold');
  const command = {
    ...owner,
    requestId: 'font-1',
    editGeneration: ctx.editGeneration,
    tool: 'set_slide_text_bold',
    input: { enabled: true },
  };
  await expect(bridge.execute(command, command.requestId)).resolves.toMatchObject({
    status: 'verified',
    result: { enabled: true, page: 1, verified: true },
  });
  expect(state.size()).toBe(true);
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
  const setter = state.api.put_TextPrBold;
  if (change === 'readonly') setReadonlyState(true);
  if (change === 'view') state.api.isViewMode = true;
  if (change === 'master') state.logic.IsMasterMode = () => true;
  if (change === 'loading') state.api.isLoadFullApi = false;
  if (change === 'busy') state.api.isLongAction = () => true;
  if (change === 'slideshow') state.api.isSlideShow = () => true;
  if (change === 'no-selection') state.run.Selection.Use = false;
  if (change === 'multiple') state.controller.selectedObjects.push(state.controller.selectedObjects[0]);
  if (change === 'foreign-shape') state.logic.Slides[0].cSld.spTree = [];
  if (change === 'missing-setter') Reflect.deleteProperty(state.api, 'put_TextPrBold');
  if (change === 'missing-reader') Reflect.deleteProperty(state.api, 'get_TextProps');
  expect(canSetSlideTextBold()).toBe(false);
  await expect(setSlideTextBoldTool.execute({ enabled: true })).rejects.toThrow();
  expect(setter).not.toHaveBeenCalled();
});

it('rejects invalid booleans and cancellation before writing', async () => {
  const state = mount();
  const bridge = new DeckWiseEditorBridge();
  bridge.reset(owner);
  for (const enabled of [null, 0, 1, 'true', undefined]) {
    const ctx = bridge.getContext(owner);
    const requestId = 'invalid-' + String(enabled);
    await expect(
      bridge.execute(
        { ...owner, requestId, editGeneration: ctx.editGeneration, tool: 'set_slide_text_bold', input: { enabled } },
        requestId,
      ),
    ).rejects.toThrow();
  }
  await expect(setSlideTextBoldTool.execute({ enabled: true }, AbortSignal.abort())).rejects.toThrow();
  expect(state.api.put_TextPrBold).not.toHaveBeenCalled();
});

it('clears existing bold with an explicit false, including a reverse selection', async () => {
  const state = mount();
  await setSlideTextBoldTool.execute({ enabled: true });
  state.run.Selection.StartPos = 5;
  state.run.Selection.EndPos = 0;
  await expect(setSlideTextBoldTool.execute({ enabled: false })).resolves.toEqual({
    enabled: false,
    page: 1,
    verified: true,
  });
  expect(state.size()).toBe(false);
});

it.each(['wrong-bold', 'changed-selection', 'changed-text', 'cancel'])(
  'rejects %s after dispatch without retrying or reporting success',
  async (change) => {
    const state = mount();
    const abort = new AbortController();
    state.api.put_TextPrBold.mockImplementation(() => {
      if (change === 'changed-selection') state.run.Selection.EndPos = 2;
      if (change === 'changed-text') state.run.Content[0].GetCodePoint = () => 88;
      if (change === 'cancel') abort.abort();
    });
    await expect(setSlideTextBoldTool.execute({ enabled: true }, abort.signal)).rejects.toThrow();
    expect(state.api.put_TextPrBold).toHaveBeenCalledTimes(1);
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
    tool: 'set_slide_text_bold',
    input: { enabled: true, script: 'bad' },
  };
  await expect(bridge.execute(command, command.requestId)).rejects.toThrow();
  expect(state.api.put_TextPrBold).not.toHaveBeenCalled();
});
