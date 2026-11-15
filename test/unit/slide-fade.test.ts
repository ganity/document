import { afterEach, expect, it, vi } from 'vitest';
import { canSetSlideFade, prepareSlideFade } from '../../lib/agent-plugin/slide-fade';
import { setReadonlyState } from '../../lib/onlyoffice/readonly';

function mount() {
  const history = { Index: -1, Points: [] as Array<{ Items: unknown[] }> };
  const transition = () => ({
    TransitionType: 0,
    TransitionOption: -1,
    TransitionDuration: 500,
    SlideAdvanceOnMouseClick: true,
    SlideAdvanceAfter: false,
    SlideAdvanceDuration: 10000,
    ShowLoop: true,
  });
  const slides = [0, 1].map((i) => ({
    transition: transition(),
    timing: { original: i },
    cSld: { spTree: [{ Id: String(i), getText: (): string => 'original', spPr: { xfrm: { offX: 20 } } }] },
    deleteLock: { Lock: { Type: 1 } },
    transitionLock: { Lock: { Type: 1 } },
  }));
  const controller = { selectedObjects: [] as unknown[], getTargetDocContent: (): undefined => undefined };
  const logic = {
    Slides: slides,
    FocusOnNotes: false,
    CanEdit: () => true,
    IsMasterMode: () => false,
    GetSelectedSlides: () => [0],
    GetCurrentController: () => controller,
    isLoopShowMode: () => false,
  };
  class NativeTransition {
    TransitionType?: number;
    TransitionOption?: number;
    TransitionDuration?: number;
    ShowLoop?: boolean;
    get_ShowLoop() {
      return this.ShowLoop;
    }
  }
  const api = {
    isDocumentLoadComplete: true,
    isLoadFullApi: true,
    isViewMode: false,
    canEdit: () => true,
    isSlideShow: () => false,
    isLongAction: () => false,
    getCurrentPage: () => 0,
    WordControl: { m_oLogicDocument: logic },
    ApplySlideTransition: vi.fn((value: NativeTransition) => {
      Object.assign(slides[0].transition, Object.fromEntries(Object.entries(value).filter(([, v]) => v !== undefined)));
      history.Index++;
      history.Points.push({ Items: [{}] });
    }),
  };
  const common = {
    History: history,
    c_oAscLockTypes: { kLockTypeNone: 1, kLockTypeMine: 2 },
    CollaborativeEditing: { Get_GlobalLock: () => false },
  };
  const asc = {
    CAscSlideTransition: NativeTransition,
    c_oAscSlideTransitionTypes: { None: 0, Fade: 1 },
    c_oAscSlideTransitionParams: { Fade_Smoothly: 0 },
  };
  const frame = document.createElement('iframe');
  document.body.append(frame);
  Object.assign(frame.contentWindow!, { editor: api, Asc: asc, AscCommon: common });
  return { api, logic, slides, controller, common, asc, frame, history };
}
afterEach(() => {
  setReadonlyState(false);
  document.body.replaceChildren();
  vi.restoreAllMocks();
});

it('prepares without mutation, applies just the current transition, preserves global loop and consumes once', () => {
  const s = mount(),
    original = JSON.stringify(s.slides.map((slide) => slide.transition));
  expect(canSetSlideFade()).toBe(true);
  const prepared = prepareSlideFade();
  expect(JSON.stringify(s.slides.map((slide) => slide.transition))).toBe(original);
  expect(prepared.apply()).toEqual({ page: 1, transition: 'fade', durationMs: 750, verified: true, saved: false });
  expect(s.api.ApplySlideTransition).toHaveBeenCalledTimes(1);
  expect(s.api.ApplySlideTransition.mock.calls[0][0].get_ShowLoop()).toBeUndefined();
  expect(s.slides[0].transition).toEqual({
    ...JSON.parse(original)[0],
    TransitionType: 1,
    TransitionOption: 0,
    TransitionDuration: 750,
  });
  expect(s.slides[1].transition).toEqual(JSON.parse(original)[1]);
  expect(s.logic.isLoopShowMode()).toBe(false);
  expect(() => prepared.apply()).toThrow(/consumed/);
  expect(() => prepareSlideFade()).toThrow(/already/);
});
it.each([
  'readonly',
  'view',
  'cannot-edit',
  'logic-readonly',
  'loading',
  'busy',
  'show',
  'master',
  'notes',
  'global-lock',
  'delete-lock',
  'transition-lock',
  'multi-slide',
  'other-slide',
  'no-slides',
  'unsupported-transition',
  'black-fade',
  'auto-advance',
  'no-click',
  'extension',
  'missing-api',
  'missing-constant',
  'no-history',
  'invalid-duration',
])('refuses %s before any edit', (change) => {
  const s = mount();
  if (change === 'readonly') setReadonlyState(true);
  if (change === 'view') s.api.isViewMode = true;
  if (change === 'cannot-edit') s.api.canEdit = () => false;
  if (change === 'logic-readonly') s.logic.CanEdit = () => false;
  if (change === 'loading') s.api.isLoadFullApi = false;
  if (change === 'busy') s.api.isLongAction = () => true;
  if (change === 'show') s.api.isSlideShow = () => true;
  if (change === 'master') s.logic.IsMasterMode = () => true;
  if (change === 'notes') s.logic.FocusOnNotes = true;
  if (change === 'global-lock') s.common.CollaborativeEditing.Get_GlobalLock = () => true;
  if (change === 'delete-lock') s.slides[0].deleteLock.Lock.Type = 3;
  if (change === 'transition-lock') s.slides[0].transitionLock.Lock.Type = 3;
  if (change === 'multi-slide') s.logic.GetSelectedSlides = () => [0, 1];
  if (change === 'other-slide') s.logic.GetSelectedSlides = () => [1];
  if (change === 'no-slides') s.logic.GetSelectedSlides = () => [];
  if (change === 'unsupported-transition') s.slides[0].transition.TransitionType = 9;
  if (change === 'black-fade') Object.assign(s.slides[0].transition, { TransitionType: 1, TransitionOption: 1 });
  if (change === 'auto-advance') s.slides[0].transition.SlideAdvanceAfter = true;
  if (change === 'no-click') s.slides[0].transition.SlideAdvanceOnMouseClick = false;
  if (change === 'extension') Object.assign(s.slides[0].transition, { Sound: 'secret' });
  if (change === 'missing-api') Reflect.deleteProperty(s.api, 'ApplySlideTransition');
  if (change === 'missing-constant') Reflect.deleteProperty(s.asc.c_oAscSlideTransitionParams, 'Fade_Smoothly');
  if (change === 'no-history') Reflect.deleteProperty(s.common, 'History');
  if (change === 'invalid-duration') s.slides[0].transition.TransitionDuration = NaN;
  expect(canSetSlideFade()).toBe(false);
  expect(() => prepareSlideFade()).toThrow();
  if (change !== 'missing-api') expect(s.api.ApplySlideTransition).not.toHaveBeenCalled();
});
it.each([
  'page',
  'selection',
  'history',
  'peer-transition',
  'peer-text',
  'timing',
  'readonly',
  'abort',
  'document',
  'input',
])('refuses stale %s after preparation', (change) => {
  const s = mount(),
    abort = new AbortController(),
    prepared = prepareSlideFade(abort.signal);
  if (change === 'page') s.api.getCurrentPage = () => 1;
  if (change === 'selection') s.controller.selectedObjects.push({});
  if (change === 'history') {
    s.history.Index++;
    s.history.Points.push({ Items: [{}] });
  }
  if (change === 'peer-transition') s.slides[1].transition.TransitionDuration = 600;
  if (change === 'peer-text') s.slides[1].cSld.spTree[0].getText = () => 'changed';
  if (change === 'timing') s.slides[0].timing = { original: 123 };
  if (change === 'readonly') setReadonlyState(true);
  if (change === 'abort') abort.abort();
  if (change === 'document') s.frame.remove();
  if (change === 'input') s.frame.contentDocument!.dispatchEvent(new Event('pointerdown', { bubbles: true }));
  expect(() => prepared.apply()).toThrow();
  expect(s.api.ApplySlideTransition).not.toHaveBeenCalled();
});
it.each(['no-op', 'wrong-value', 'peer-transition', 'peer-text', 'loop', 'throw', 'abort'])(
  'does not report success or retry after native %s',
  (change) => {
    const s = mount(),
      abort = new AbortController(),
      original = s.api.ApplySlideTransition.getMockImplementation()!;
    s.api.ApplySlideTransition.mockImplementation((value) => {
      if (change === 'no-op') return;
      original(value);
      if (change === 'wrong-value') s.slides[0].transition.TransitionDuration = 100;
      if (change === 'peer-transition') s.slides[1].transition.TransitionDuration = 100;
      if (change === 'peer-text') s.slides[1].cSld.spTree[0].getText = () => 'changed';
      if (change === 'loop') s.logic.isLoopShowMode = () => true;
      if (change === 'throw') throw new Error('native error');
      if (change === 'abort') abort.abort();
    });
    const prepared = prepareSlideFade(abort.signal);
    expect(() => prepared.apply()).toThrow(/could not be verified/);
    expect(() => prepared.apply()).toThrow(/consumed/);
    expect(s.api.ApplySlideTransition).toHaveBeenCalledTimes(1);
  },
);
it.each(['loop', 'advance', 'unknown', 'missing-getter'])(
  'rejects incompatible constructor %s before invoking mutation',
  (change) => {
    const s = mount();
    class Incompatible extends s.asc.CAscSlideTransition {
      constructor() {
        super();
        if (change === 'loop') this.ShowLoop = true;
        if (change === 'advance') Object.assign(this, { SlideAdvanceAfter: true });
        if (change === 'unknown') Object.assign(this, { Sound: 'unexpected' });
        if (change === 'missing-getter') Object.assign(this, { get_ShowLoop: undefined });
      }
    }
    s.asc.CAscSlideTransition = Incompatible;
    expect(() => prepareSlideFade().apply()).toThrow(/constructor/);
    expect(s.api.ApplySlideTransition).not.toHaveBeenCalled();
  },
);
