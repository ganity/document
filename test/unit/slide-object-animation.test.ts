import { afterEach, beforeEach, expect, it, vi } from 'vitest';
import { addSlideObjectAnimation, canAddSlideObjectAnimation } from '../../lib/agent-plugin/slide-object-animation';
import { setReadonlyState } from '../../lib/onlyoffice/readonly';
import { DeckWiseEditorBridge } from '../../lib/deckwise-editor-bridge';

function mount() {
  const shape = {
    Id: 'ordinary-1',
    group: null,
    isPlaceholder: () => false,
    getObjectType: () => 72548352,
    blipFill: { RasterImageId: 'original.png', srcRect: null },
    spPr: {
      geometry: { preset: 'rect' },
      xfrm: { offX: 10, offY: 10, extX: 20, extY: 20, rot: 0, flipH: false, flipV: false },
    },
  };
  const effects: Array<ReturnType<typeof effect>> = [];
  function effect(presetID: number, target = shape.Id) {
    return {
      cTn: { presetClass: 1, presetID, presetSubtype: 0, nodeType: 2 },
      getObjectId: () => target,
      asc_getDuration: () => (presetID === 1 ? 1 : 500),
      asc_getDelay: () => 0,
    };
  }
  const timing = {
    bldLst: null,
    getAllAnimEffects: () => effects,
    getSelectedEffects: () => [],
    getEffectsSequences: () => (effects.length ? [[null, ...effects]] : []),
  };
  const slide = { cSld: { spTree: [shape] }, timing };
  const controller = { selectedObjects: [shape], getTargetDocContent: () => undefined };
  const logic = {
    Slides: [slide],
    CanEdit: () => true,
    IsMasterMode: () => false,
    IsSelectionLocked: () => false,
    GetWidthMM: () => 320,
    GetHeightMM: () => 200,
    GetCurrentController: () => controller,
    GetSelectedSlides: () => [0],
  };
  const history = { Index: -1, Points: [] as Array<{ Items: unknown[] }> };
  const api = {
    isDocumentLoadComplete: true,
    isLoadFullApi: true,
    isViewMode: false,
    canEdit: () => true,
    isSlideShow: () => false,
    isLongAction: () => false,
    getCurrentPage: () => 0,
    asc_IsStartedAnimationPreview: () => false,
    WordControl: { m_oLogicDocument: logic },
    ImgApply: vi.fn(),
    ImageLoader: { LoadImage: vi.fn(), LoadImagesWithCallback: vi.fn() },
    asc_AddAnimation: vi.fn((_class: number, preset: number) => effects.push(effect(preset))),
  };
  const frame = document.createElement('iframe');
  document.body.append(frame);
  Object.assign(frame.contentWindow!, {
    editor: api,
    Asc: {},
    AscCommon: {
      History: history,
      changestype_Timing: 7,
      CollaborativeEditing: { Get_GlobalLock: () => false },
      sendImgUrls: vi.fn(),
      getFullImageSrc2: vi.fn(),
      g_oDocumentUrls: { getImageLocal: vi.fn(), imagePath2Local: vi.fn() },
    },
    AscDFH: { historyitem_type_ImageShape: 72548352, historyitem_type_Shape: 555 },
    AscFormat: { PRESET_CLASS_ENTR: 1, ENTRANCE_APPEAR: 1, ENTRANCE_FADE: 10, NODE_TYPE_CLICKEFFECT: 2 },
  });
  return { shape, slide, effects, effect, controller, logic, api, frame };
}
beforeEach(() => setReadonlyState(false));
afterEach(() => {
  setReadonlyState(false);
  document.body.replaceChildren();
  vi.restoreAllMocks();
});

it.each(['appear', 'fade'] as const)(
  'appends native %s once and preserves prior effects, image and geometry',
  (name) => {
    const s = mount(),
      prior = s.effect(10);
    s.effects.push(prior);
    const before = JSON.stringify(s.shape);
    expect(canAddSlideObjectAnimation()).toBe(true);
    expect(addSlideObjectAnimation({ effect: name })).toEqual({ effect: name, page: 1, verified: true });
    expect(s.api.asc_AddAnimation).toHaveBeenCalledExactlyOnceWith(
      1,
      name === 'appear' ? 1 : 10,
      0,
      null,
      false,
      false,
    );
    expect(s.effects[0]).toBe(prior);
    expect(JSON.stringify(s.shape)).toBe(before);
  },
);
it('supports an ordinary rectangular shape without text', () => {
  const s = mount();
  s.shape.getObjectType = () => 555;
  Reflect.deleteProperty(s.shape, 'blipFill');
  expect(canAddSlideObjectAnimation()).toBe(true);
  expect(addSlideObjectAnimation({ effect: 'appear' }).verified).toBe(true);
});
it('supports a whole selected plain text box without a nonempty text selection', () => {
  const s = mount();
  s.shape.getObjectType = () => 555;
  Reflect.deleteProperty(s.shape, 'blipFill');
  const content = { GetText: () => 'Lesson', GetSelectionState: () => [] };
  Reflect.set(s.shape, 'getDocContent', () => content);
  Reflect.set(s.controller, 'getTargetDocContent', () => content);
  expect(canAddSlideObjectAnimation()).toBe(true);
  expect(addSlideObjectAnimation({ effect: 'fade' }).verified).toBe(true);
});
it('binds the capability and one-shot command to the current native generation', async () => {
  const s = mount(),
    bridge = new DeckWiseEditorBridge();
  const owner = { projectId: 'project', documentId: 'document', sessionId: 'session' };
  bridge.reset(owner);
  const context = bridge.getContext(owner);
  expect(context.selectedText).toBe('');
  expect(context.capabilities.some((c) => c.name === 'add_slide_object_animation')).toBe(true);
  const command = {
    ...owner,
    editGeneration: context.editGeneration,
    requestId: 'animation-once',
    tool: 'add_slide_object_animation',
    input: { effect: 'appear' },
  };
  expect(await bridge.execute(command, command.requestId)).toMatchObject({
    status: 'verified',
    result: { effect: 'appear', page: 1, verified: true },
  });
  await expect(bridge.execute(command, command.requestId)).rejects.toThrow(/consumed/);
  expect(s.api.asc_AddAnimation).toHaveBeenCalledTimes(1);
});
it.each(['selection', 'readonly', 'generation'])(
  'rejects late %s changes at the bridge before mutation',
  async (change) => {
    const s = mount(),
      bridge = new DeckWiseEditorBridge();
    const owner = { projectId: 'project', documentId: 'document', sessionId: 'session' };
    bridge.reset(owner);
    const context = bridge.getContext(owner);
    if (change === 'selection') s.controller.selectedObjects = [];
    if (change === 'readonly') setReadonlyState(true);
    const command = {
      ...owner,
      editGeneration: context.editGeneration + (change === 'generation' ? 1 : 0),
      requestId: 'late-command',
      tool: 'add_slide_object_animation',
      input: { effect: 'fade' },
    };
    await expect(bridge.execute(command, command.requestId)).rejects.toThrow(/context.*changed|fresh context/);
    expect(s.api.asc_AddAnimation).not.toHaveBeenCalled();
  },
);
it.each([
  'readonly',
  'locked',
  'preview',
  'group',
  'placeholder',
  'cropped',
  'rotation',
  'offpage',
  'multiple',
  'complex',
  'builds',
  'foreign-effect',
  'unknown-sdk',
])('rejects %s before native dispatch', (change) => {
  const s = mount();
  if (change === 'readonly') setReadonlyState(true);
  if (change === 'locked') s.logic.IsSelectionLocked = () => true;
  if (change === 'preview') s.api.asc_IsStartedAnimationPreview = () => true;
  if (change === 'group') Reflect.set(s.shape, 'group', {});
  if (change === 'placeholder') s.shape.isPlaceholder = () => true;
  if (change === 'cropped') Reflect.set(s.shape.blipFill, 'srcRect', { l: 1 });
  if (change === 'rotation') s.shape.spPr.xfrm.rot = 1;
  if (change === 'offpage') s.shape.spPr.xfrm.offX = 319;
  if (change === 'multiple') s.controller.selectedObjects.push(s.shape);
  if (change === 'complex') s.effects.push(s.effect(99));
  if (change === 'builds') Reflect.set(s.slide.timing, 'bldLst', {});
  if (change === 'foreign-effect') {
    const foreign = s.effect(10, 'other');
    Reflect.set(s.slide.timing, 'getSelectedEffects', () => [foreign]);
  }
  if (change === 'unknown-sdk') Reflect.set(s.frame.contentWindow!, 'AscFormat', {});
  expect(canAddSlideObjectAnimation()).toBe(false);
  expect(() => addSlideObjectAnimation({ effect: 'fade' })).toThrow();
  expect(s.api.asc_AddAnimation).not.toHaveBeenCalled();
});
it.each([{}, { effect: 'zoom' }, { effect: 'appear', duration: 1000 }, { effect: 'fade', replace: true }])(
  'rejects unsupported parameters %j',
  (input) => {
    const s = mount();
    expect(() => addSlideObjectAnimation(input)).toThrow();
    expect(s.api.asc_AddAnimation).not.toHaveBeenCalled();
  },
);
it.each(['wrong-target', 'changed-old-effect', 'changed-media', 'changed-page', 'aborted'] as const)(
  'reports unverified %s without a second dispatch',
  (change) => {
    const s = mount(),
      abort = new AbortController();
    s.effects.push(s.effect(10));
    s.api.asc_AddAnimation.mockImplementation((_class, preset) => {
      s.effects.push(s.effect(preset, change === 'wrong-target' ? 'other' : s.shape.Id));
      if (change === 'changed-old-effect') s.effects[0].cTn.presetSubtype = 1;
      if (change === 'changed-media') s.shape.blipFill.RasterImageId = 'changed.png';
      if (change === 'changed-page') s.api.getCurrentPage = () => 1;
      if (change === 'aborted') abort.abort();
      return s.effects.length;
    });
    expect(() => addSlideObjectAnimation({ effect: 'fade' }, abort.signal)).toThrow(/unverified.*Undo/);
    expect(s.api.asc_AddAnimation).toHaveBeenCalledTimes(1);
  },
);
