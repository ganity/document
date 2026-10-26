import { afterEach, beforeEach, expect, it, vi } from 'vitest';
import { canSetSlideTextColor, setSlideTextColorTool } from '../../lib/agent-plugin/slide-text-color';
import { DeckWiseEditorBridge } from '../../lib/deckwise-editor-bridge';
import { setReadonlyState } from '../../lib/onlyoffice/readonly';

const owner = { projectId: 'project', documentId: 'doc', sessionId: 'session' };
function mount() {
  let rgb = { R: 0, G: 0, B: 0, A: 255 };
  const calculated = {
    Unifill: {
      fill: {
        type: 3,
        color: {
          color: { type: 1 },
          get RGBA() {
            return rgb;
          },
        },
      },
    },
  };
  class Color {
    r = 0;
    g = 0;
    b = 0;
    type = 0;
    put_type(v: number) {
      this.type = v;
    }
    put_r(v: number) {
      this.r = v;
    }
    put_g(v: number) {
      this.g = v;
    }
    put_b(v: number) {
      this.b = v;
    }
  }
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
    GetCalculatedTextPr: () => calculated,
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
    put_TextColor: vi.fn((value: Color) => {
      expect(value.type).toBe(1);
      rgb = { R: value.r, G: value.g, B: value.b, A: 255 };
      history.Index++;
      history.Points.push({ Items: [{}] });
    }),
  };
  const frame = document.createElement('iframe');
  document.body.append(frame);
  Object.assign(frame.contentWindow!, {
    editor: api,
    Asc: { asc_CColor: Color, c_oAscColor: { COLOR_TYPE_SRGB: 1 }, c_oAscFill: { FILL_TYPE_SOLID: 3 } },
    AscCommon: { History: history },
  });
  return { api, logic, controller, content, run, history, rgb: () => rgb, frame, calculated };
}
beforeEach(() => setReadonlyState(false));
afterEach(() => {
  setReadonlyState(false);
  document.body.replaceChildren();
  vi.restoreAllMocks();
});

it('sets one selected text range natively, verifies color/text and advances the bound generation', async () => {
  const state = mount();
  const bridge = new DeckWiseEditorBridge();
  bridge.reset(owner);
  const ctx = bridge.getContext(owner);
  expect(ctx.capabilities.map((c) => c.name)).toContain('set_slide_text_color');
  const command = {
    ...owner,
    requestId: 'font-1',
    editGeneration: ctx.editGeneration,
    tool: 'set_slide_text_color',
    input: { colorHex: '336699' },
  };
  await expect(bridge.execute(command, command.requestId)).resolves.toMatchObject({
    status: 'verified',
    result: { colorHex: '336699', page: 1, verified: true },
  });
  expect(state.rgb()).toEqual({ R: 51, G: 102, B: 153, A: 255 });
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
  'missing-constructor',
  'missing-color-method',
  'unsupported-constants',
])('withholds the capability and rejects %s without calling a native setter', async (change) => {
  const state = mount();
  const setter = state.api.put_TextColor;
  if (change === 'readonly') setReadonlyState(true);
  if (change === 'view') state.api.isViewMode = true;
  if (change === 'master') state.logic.IsMasterMode = () => true;
  if (change === 'loading') state.api.isLoadFullApi = false;
  if (change === 'busy') state.api.isLongAction = () => true;
  if (change === 'slideshow') state.api.isSlideShow = () => true;
  if (change === 'no-selection') state.run.Selection.Use = false;
  if (change === 'multiple') state.controller.selectedObjects.push(state.controller.selectedObjects[0]);
  if (change === 'foreign-shape') state.logic.Slides[0].cSld.spTree = [];
  if (change === 'missing-setter') Reflect.deleteProperty(state.api, 'put_TextColor');
  if (change === 'missing-reader') Reflect.deleteProperty(state.logic, 'GetCalculatedTextPr');
  const asc = Reflect.get(state.frame.contentWindow!, 'Asc');
  if (change === 'missing-constructor') Reflect.deleteProperty(asc, 'asc_CColor');
  if (change === 'missing-color-method') Reflect.deleteProperty(asc.asc_CColor.prototype, 'put_b');
  if (change === 'unsupported-constants') asc.c_oAscFill.FILL_TYPE_SOLID = 99;
  expect(canSetSlideTextColor()).toBe(false);
  await expect(setSlideTextColorTool.execute({ colorHex: 'AABBCC' })).rejects.toThrow();
  expect(setter).not.toHaveBeenCalled();
});

it('rejects invalid colors and cancellation before writing', async () => {
  const state = mount();
  for (const color of ['', '#AABBCC', 'aabbcc', 'GG0011', '12345', 'AABBCCDD'])
    await expect(setSlideTextColorTool.execute({ colorHex: color })).rejects.toThrow();
  await expect(setSlideTextColorTool.execute({ colorHex: 'AABBCC' }, AbortSignal.abort())).rejects.toThrow();
  expect(state.api.put_TextColor).not.toHaveBeenCalled();
});

it.each([
  'wrong-color',
  'changed-selection',
  'changed-text',
  'cancel',
  'theme-color',
  'transparent',
  'mixed',
  'gradient',
  'fill-alpha',
  'modifier',
  'invalid-channel',
])('rejects %s after dispatch without retrying or reporting success', async (change) => {
  const state = mount();
  const abort = new AbortController();
  const apply = state.api.put_TextColor.getMockImplementation()!;
  state.api.put_TextColor.mockImplementation((color) => {
    apply(color);
    if (change === 'wrong-color') state.rgb().R = 0;
    if (change === 'theme-color') state.calculated.Unifill.fill.color.color.type = 3;
    if (change === 'transparent') state.rgb().A = 128;
    if (change === 'gradient') state.calculated.Unifill.fill.type = 4;
    if (change === 'fill-alpha') Reflect.set(state.calculated.Unifill, 'transparent', 128);
    if (change === 'modifier') Reflect.set(state.calculated.Unifill.fill.color, 'Mods', { Mods: [{}] });
    if (change === 'invalid-channel') state.rgb().R = 256;
    if (change === 'mixed') Reflect.deleteProperty(state.calculated, 'Unifill');
    if (change === 'changed-selection') state.run.Selection.EndPos = 2;
    if (change === 'changed-text') state.run.Content[0].GetCodePoint = () => 88;
    if (change === 'cancel') abort.abort();
  });
  await expect(setSlideTextColorTool.execute({ colorHex: 'AABBCC' }, abort.signal)).rejects.toThrow();
  expect(state.api.put_TextColor).toHaveBeenCalledTimes(1);
});

it('rejects removed capability and foreign command fields at the bridge', async () => {
  const state = mount();
  const bridge = new DeckWiseEditorBridge();
  bridge.reset(owner);
  const context = bridge.getContext(owner);
  const command = {
    ...owner,
    requestId: 'font-1',
    editGeneration: context.editGeneration,
    tool: 'set_slide_text_color',
    input: { colorHex: 'AABBCC', script: 'bad' },
  };
  await expect(bridge.execute(command, command.requestId)).rejects.toThrow();
  expect(state.api.put_TextColor).not.toHaveBeenCalled();
});
