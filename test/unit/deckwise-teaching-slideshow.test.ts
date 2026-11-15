import { afterEach, beforeEach, expect, it, vi } from 'vitest';
const state = vi.hoisted(() => ({ api: null as any, dirty: false }));
vi.mock('../../lib/agent-plugin/editor-bridge', () => ({ getEditorApi: () => state.api }));
vi.mock('../../lib/unsaved-guard', () => ({ hasUnsavedChanges: () => state.dirty }));
import { createTeachingSlideshow, inspectTeachingAnimations } from '../../lib/deckwise-teaching-slideshow';
import * as nativeTeaching from '../../lib/deckwise-native-teaching';
let manager: any, callbacks: Map<string, () => void>, complete: boolean;
let show: ReturnType<typeof createTeachingSlideshow> | undefined;
beforeEach(() => {
  vi.useFakeTimers();
  state.dirty = false;
  complete = false;
  callbacks = new Map();
  manager = {
    Mode: false,
    SlideNum: -1,
    DemonstrationDiv: null,
    Transition: { IsPlaying: () => false },
    IsMainSeqFinished: () => complete,
    StartAnimation: vi.fn(),
    OnNextSlideAnimPlayer: vi.fn(() => true),
  };
  state.api = {
    WordControl: { DemonstrationManager: manager, m_oLogicDocument: { Slides: [{ cSld: { spTree: [] } }] } },
    isSlideShow: () => manager.Mode,
    asc_registerCallback: (event: string, fn: () => void) => callbacks.set(event, fn),
    asc_unregisterCallback: (event: string) => callbacks.delete(event),
    StartDemonstration: (id: string, page: number) => {
      manager.Mode = true;
      manager.SlideNum = page;
      manager.DemonstrationDiv = document.getElementById(id);
    },
    EndDemonstration: vi.fn(() => {
      manager.Mode = false;
      callbacks.get('asc_onEndDemonstration')?.();
    }),
    DemonstrationPause: vi.fn(),
    DemonstrationPlay: vi.fn(),
  };
});
afterEach(() => {
  show?.dispose();
  show = undefined;
  document.body.replaceChildren();
  vi.useRealTimers();
  vi.unstubAllGlobals();
});
it('defers control layout writes outside ResizeObserver delivery and cancels them on exit', () => {
  let resized: () => void = () => {};
  const disconnect = vi.fn();
  vi.stubGlobal(
    'ResizeObserver',
    class {
      constructor(callback: () => void) {
        resized = callback;
      }
      observe() {}
      disconnect = disconnect;
    },
  );
  show = createTeachingSlideshow({ doc: document, page: 1, onReady: vi.fn(), onExit: vi.fn(), onError: vi.fn() });
  const height = vi.spyOn(show.controls, 'getBoundingClientRect').mockReturnValue({ height: 128 } as DOMRect);
  const style = vi.spyOn(show.surface.style, 'setProperty');
  resized();
  resized();
  expect(style).not.toHaveBeenCalled();
  vi.advanceTimersByTime(32);
  expect(style).toHaveBeenCalledExactlyOnceWith('--teaching-controls-height', '128px');
  resized();
  vi.advanceTimersByTime(32);
  expect(style).toHaveBeenCalledTimes(1);
  height.mockReturnValue({ height: 160 } as DOMRect);
  resized();
  show.dispose();
  vi.advanceTimersByTime(32);
  expect(style).toHaveBeenCalledTimes(1);
  expect(disconnect).toHaveBeenCalledOnce();
});
it('starts teaching exactly once when the native canvas is ready', () => {
  const ready = vi.fn();
  show = createTeachingSlideshow({ doc: document, page: 1, onReady: ready, onExit: vi.fn(), onError: vi.fn() });
  vi.advanceTimersByTime(64);
  expect(ready).toHaveBeenCalledOnce();
  complete = true;
  vi.advanceTimersByTime(64);
  expect(ready).toHaveBeenCalledTimes(1);
  show.pause();
  show.resume();
  expect(state.api.DemonstrationPlay).not.toHaveBeenCalled();
  show.dispose();
  expect(callbacks.size).toBe(0);
  expect(manager.Mode).toBe(false);
  expect(document.querySelector('[data-teaching-slideshow]')).toBeNull();
});
it('freezes native preparation and exits on manual page navigation', async () => {
  const ready = vi.fn(),
    exit = vi.fn(() => show?.dispose());
  show = createTeachingSlideshow({ doc: document, page: 1, onReady: ready, onExit: exit, onError: vi.fn() });
  show.pause();
  complete = true;
  vi.advanceTimersByTime(64);
  expect(ready).not.toHaveBeenCalled();
  show.resume();
  vi.advanceTimersByTime(32);
  expect(ready).toHaveBeenCalledTimes(1);
  manager.SlideNum = 1;
  callbacks.get('asc_onDemonstrationSlideChanged')!();
  expect(exit).not.toHaveBeenCalled();
  await Promise.resolve();
  expect(exit).toHaveBeenCalledTimes(1);
  expect(callbacks.size).toBe(0);
});
it('rejects unsaved edits, automatic page advance and native embedded media before starting', () => {
  const input = { doc: document, page: 1, onReady: vi.fn(), onExit: vi.fn(), onError: vi.fn() };
  state.dirty = true;
  expect(() => createTeachingSlideshow(input)).toThrow('保存');
  state.dirty = false;
  const slide = state.api.WordControl.m_oLogicDocument.Slides[0];
  slide.isAdvanceAfterTransition = () => true;
  expect(() => createTeachingSlideshow(input)).toThrow('自动换页');
  slide.isAdvanceAfterTransition = () => false;
  slide.cSld.spTree = [{ getMediaFileName: () => 'embedded.mp4' }];
  expect(() => createTeachingSlideshow(input)).toThrow('嵌入媒体');
  expect(manager.Mode).toBe(false);
});
it('reports stale source and native errors without starting teaching', () => {
  const error = vi.fn((_error: Error) => show?.dispose()),
    exit = vi.fn(() => show?.dispose()),
    ready = vi.fn();
  show = createTeachingSlideshow({ doc: document, page: 1, onReady: ready, onExit: exit, onError: error });
  manager.Transition.IsPlaying = () => {
    throw new Error('native failure');
  };
  vi.advanceTimersByTime(32);
  expect(error.mock.calls[0][0].message).toBe('native failure');
  expect(ready).not.toHaveBeenCalled();
  expect(callbacks.size).toBe(0);
  show = createTeachingSlideshow({ doc: document, page: 1, onReady: ready, onExit: exit, onError: error });
  state.dirty = true;
  vi.advanceTimersByTime(32);
  expect(exit).toHaveBeenCalledTimes(1);
});

it('allows verified appear/fade but rejects animated geometry and nested native media', () => {
  const input = { doc: document, page: 1, onReady: vi.fn(), onExit: vi.fn(), onError: vi.fn() };
  const slide = state.api.WordControl.m_oLogicDocument.Slides[0];
  slide.cSld.spTree = [{ Id: 'target' }];
  slide.timing = {
    getAllAnimEffects: () => [
      {
        cTn: { presetClass: 1, presetID: 10, presetSubtype: 0, nodeType: 2 },
        getObjectId: () => 'target',
        asc_getDuration: () => 500,
        asc_getDelay: () => 0,
        isIdle: () => true,
        isAtEnd: () => false,
      },
    ],
  };
  slide.timing.getEffectsSequences = () => [[null, ...slide.timing.getAllAnimEffects()]];
  const effect = slide.timing.getAllAnimEffects()[0];
  slide.timing.getAllAnimEffects = () => [effect];
  show = createTeachingSlideshow(input);
  show.dispose();
  slide.timing.getAllAnimEffects = () => [{ cTn: { presetClass: 4, presetID: 1, presetSubtype: 0 } }];
  expect(() => createTeachingSlideshow(input)).toThrow('animation state');
  slide.timing = undefined;
  slide.cSld.spTree = [{ spTree: [{ isAudio: () => true }] }];
  expect(() => createTeachingSlideshow(input)).toThrow('嵌入媒体');
});

it('rejects triggered native sequences before starting any playback', () => {
  const effect = { cTn: { presetClass: 1, presetID: 10, presetSubtype: 0 } };
  state.api.WordControl.m_oLogicDocument.Slides[0].timing = {
    getAllAnimEffects: () => [effect],
    getEffectsSequences: () => [['trigger', effect]],
  };
  expect(() =>
    createTeachingSlideshow({ doc: document, page: 1, onReady: vi.fn(), onExit: vi.fn(), onError: vi.fn() }),
  ).toThrow('animation state');
  expect(manager.Mode).toBe(false);
});

function steps() {
  const state = ['idle', 'idle'];
  const effects = state.map((_, index) => ({
    cTn: { presetClass: 1, presetID: 10, presetSubtype: 0, nodeType: 2 },
    getObjectId: () => 'target',
    asc_getDuration: () => 500,
    asc_getDelay: (): number => 0,
    isIdle: () => state[index] === 'idle',
    isAtEnd: () => state[index] === 'end',
  }));
  const slide = stateApiSlide();
  slide.cSld.spTree = [{ Id: 'target' }];
  slide.timing = { getAllAnimEffects: () => effects, getEffectsSequences: () => [[null, ...effects]] };
  manager.OnNextSlideAnimPlayer.mockImplementation(() => {
    const index = state.findIndex((value) => value === 'idle');
    if (index < 0) return false;
    state[index] = 'active';
    return true;
  });
  manager.IsMainSeqFinished = () => state.every((value) => value === 'end');
  return { state, effects, slide };
}
function stateApiSlide() {
  return state.api.WordControl.m_oLogicDocument.Slides[0];
}
it('keeps finite automatic group animations beside native teaching without blocking narration', () => {
  const native = vi.spyOn(nativeTeaching, 'hasNativeTeaching').mockReturnValue(true);
  try {
    const s = steps();
    const group = { Id: 'group', spTree: [] as any[] };
    group.spTree.push({ Id: 'target', group });
    s.slide.cSld.spTree = [group];
    s.effects[0].cTn.nodeType = 7;
    Object.assign(s.effects[1].cTn, { nodeType: 0, presetClass: 0, presetID: 6 });
    s.effects[1].asc_getDelay = () => 500;
    const ready = vi.fn();
    show = createTeachingSlideshow({ doc: document, page: 1, onReady: ready, onExit: vi.fn(), onError: vi.fn() });
    vi.advanceTimersByTime(64);
    expect(ready).toHaveBeenCalledOnce();
    s.state.fill('end');
    vi.advanceTimersByTime(64);
    expect(ready).toHaveBeenCalledOnce();
  } finally {
    native.mockRestore();
  }
});
it('plays delayed automatic PPT-master effects beside narration without changing saved timing', () => {
  const s = steps(),
    ready = vi.fn(),
    error = vi.fn();
  s.effects[0].cTn.nodeType = 0;
  s.effects[1].cTn.nodeType = 0;
  s.effects[1].cTn.presetID = 37;
  s.effects[1].asc_getDelay = () => 500;
  const saved = s.slide.timing;
  expect(() => inspectTeachingAnimations(1)).toThrow('尚未支持同步教学提示');
  show = createTeachingSlideshow({ doc: document, page: 1, onReady: ready, onExit: vi.fn(), onError: error });
  vi.advanceTimersByTime(32);
  expect(ready).toHaveBeenCalledOnce();
  expect(manager.OnNextSlideAnimPlayer).not.toHaveBeenCalled();
  expect(state.api.DemonstrationPause).not.toHaveBeenCalled();
  expect(s.slide.timing).toBe(saved);
  expect(error).not.toHaveBeenCalled();
  show.pause();
  show.resume();
  expect(manager.StartAnimation).toHaveBeenCalledWith(0);
});
it('automatically finishes unlinked native steps alongside narration without clicks or advancing the page', () => {
  const s = steps(),
    ready = vi.fn(),
    error = vi.fn();
  show = createTeachingSlideshow({ doc: document, page: 1, onReady: ready, onExit: vi.fn(), onError: error });
  vi.advanceTimersByTime(32);
  expect(s.state).toEqual(['active', 'idle']);
  vi.advanceTimersByTime(1000);
  expect(manager.OnNextSlideAnimPlayer).toHaveBeenCalledTimes(1);
  expect(ready).toHaveBeenCalledOnce();
  show.pause();
  s.state[0] = 'end';
  vi.advanceTimersByTime(64);
  expect(s.state[1]).toBe('idle');
  show.resume();
  vi.advanceTimersByTime(32);
  expect(s.state[1]).toBe('active');
  s.state[1] = 'end';
  vi.advanceTimersByTime(64);
  expect(ready).toHaveBeenCalledOnce();
  vi.advanceTimersByTime(1000);
  expect(manager.OnNextSlideAnimPlayer).toHaveBeenCalledTimes(2);
  expect(manager.SlideNum).toBe(0);
  expect(error).not.toHaveBeenCalled();
});

it('autoplays only unlinked effects when speech-linked animations share the page', async () => {
  const s = steps(),
    ready = vi.fn(),
    error = vi.fn();
  const linked = s.effects[1];
  const player = {
    onSpClick: vi.fn(() => {
      s.state[1] = 'active';
      return true;
    }),
  };
  s.slide.getAnimationPlayer = () => player;
  manager.IsMainSeqFinished = () => s.state[0] === 'end';
  const schedule = {
    slide: s.slide,
    timing: s.slide.timing,
    hasUnbound: true,
    scheduled: [{ effect: linked, token: 'linked' }],
    install() {},
    restore: vi.fn(),
  };
  show = createTeachingSlideshow({
    doc: document,
    page: 1,
    timelineStepCount: 1,
    schedule: schedule as any,
    onReady: ready,
    onExit: vi.fn(),
    onError: error,
  });
  vi.advanceTimersByTime(32);
  expect(s.state).toEqual(['active', 'idle']);
  vi.advanceTimersByTime(1000);
  expect(manager.OnNextSlideAnimPlayer).toHaveBeenCalledTimes(1);
  s.state[0] = 'end';
  vi.advanceTimersByTime(64);
  expect(ready).toHaveBeenCalledOnce();
  expect(s.state[1]).toBe('idle');
  const pending = show.next(0);
  expect(player.onSpClick).toHaveBeenCalledOnce();
  s.state[1] = 'end';
  vi.advanceTimersByTime(32);
  await pending;
  expect(manager.OnNextSlideAnimPlayer).toHaveBeenCalledTimes(1);
  expect(error).not.toHaveBeenCalled();
  show.dispose();
  expect(schedule.restore).toHaveBeenCalledOnce();
});

it.each([
  [1, 23, 16], // PPT-master basic zoom, distinct from the SDK's zoom preset 53.
  [0, 1, 2], // Fill color.
  [0, 30, 0], // Lighten.
])('plays a saved PPT-master preset %i/%i/%i to native completion', async (presetClass, presetID, presetSubtype) => {
  const s = steps();
  Object.assign(s.effects[0].cTn, { presetClass, presetID, presetSubtype });
  const ready = vi.fn();
  show = createTeachingSlideshow({ doc: document, page: 1, onReady: ready, onExit: vi.fn(), onError: vi.fn() });
  const next = show.next();
  expect(s.state[0]).toBe('active');
  expect(ready).not.toHaveBeenCalled();
  s.state[0] = 'end';
  vi.advanceTimersByTime(32);
  await next;
  expect(ready).toHaveBeenCalledOnce();
  expect(s.state[1]).toBe('active');
  s.state[1] = 'end';
  vi.advanceTimersByTime(32);
  expect(ready).toHaveBeenCalledTimes(1);
});
it('completes each native step from actual state, blocking duplicates and preserving pause', async () => {
  const s = steps(),
    ready = vi.fn(),
    error = vi.fn();
  show = createTeachingSlideshow({ doc: document, page: 1, onReady: ready, onExit: vi.fn(), onError: error });
  const first = show.next(),
    finished = vi.fn();
  void first.then(finished);
  await expect(show.next()).rejects.toThrow();
  vi.advanceTimersByTime(5000);
  expect(finished).not.toHaveBeenCalled();
  show.pause();
  s.state[0] = 'end';
  vi.advanceTimersByTime(64);
  expect(finished).not.toHaveBeenCalled();
  show.resume();
  expect(manager.StartAnimation).toHaveBeenCalledWith(0);
  expect(state.api.DemonstrationPlay).not.toHaveBeenCalled();
  vi.advanceTimersByTime(64);
  await first;
  expect(ready).toHaveBeenCalledOnce();
  expect(s.state[1]).toBe('active');
  s.state[1] = 'end';
  vi.advanceTimersByTime(64);
  expect(ready).toHaveBeenCalledOnce();
  expect(error).not.toHaveBeenCalled();
  await expect(show.next()).rejects.toThrow();
  expect(document.querySelector('[data-native-animation-controls]')).toBeNull();
  expect(manager.SlideNum).toBe(0);
});

it('keeps native completion checks for an effect on a mapped group child', async () => {
  const s = steps();
  const child = { Id: 'target', group: undefined as unknown };
  const group = { Id: 'group', spTree: [child] };
  child.group = group;
  s.slide.cSld.spTree = [group];
  show = createTeachingSlideshow({ doc: document, page: 1, onReady: vi.fn(), onExit: vi.fn(), onError: vi.fn() });
  const next = show.next();
  expect(s.state[0]).toBe('active');
  s.state[0] = 'end';
  vi.advanceTimersByTime(32);
  await next;
});
it('rejects pending steps on disposal, failure or manual advancement without reporting success', async () => {
  for (const kind of ['dispose', 'advance', 'refuse'] as const) {
    const s = steps(),
      error = vi.fn(() => show?.dispose());
    show = createTeachingSlideshow({ doc: document, page: 1, onReady: vi.fn(), onExit: vi.fn(), onError: error });
    if (kind === 'refuse') manager.OnNextSlideAnimPlayer.mockReturnValueOnce(false);
    const pending = show.next(),
      rejection = expect(pending).rejects.toThrow();
    if (kind === 'dispose') show.dispose();
    if (kind === 'advance') {
      s.state[1] = 'active';
      vi.advanceTimersByTime(64);
    }
    await rejection;
    if (kind === 'advance') expect(error).toHaveBeenCalledOnce();
    show.dispose();
  }
});
it('rejects unsupported step state, duration, delay, target and changed page before starting a step', async () => {
  const input = { doc: document, page: 1, onReady: vi.fn(), onExit: vi.fn(), onError: vi.fn() };
  for (const change of [
    { asc_getDuration: () => 0 },
    { asc_getDuration: () => 5001 },
    { asc_getDelay: () => 10 },
    { isIdle: undefined },
    { getObjectId: () => 'missing' },
    { cTn: { presetClass: 1, presetID: 10, presetSubtype: 0, nodeType: 3 } },
  ]) {
    const s = steps();
    Object.assign(s.effects[0], change);
    expect(() => createTeachingSlideshow({ ...input, timelineStepCount: 2 })).toThrow();
    expect(manager.Mode).toBe(false);
  }
  steps();
  show = createTeachingSlideshow(input);
  state.dirty = true;
  await expect(show.next()).rejects.toThrow('页面已变化');
  expect(manager.OnNextSlideAnimPlayer).not.toHaveBeenCalled();
});
it('fences changed effect identity immediately and rejects a pending step on native sequence replacement', async () => {
  const s = steps(),
    error = vi.fn(() => show?.dispose());
  show = createTeachingSlideshow({ doc: document, page: 1, onReady: vi.fn(), onExit: vi.fn(), onError: error });
  s.slide.timing.getAllAnimEffects = () => [{ ...s.effects[0] }, s.effects[1]];
  await expect(show.next()).rejects.toThrow('序列已变化');
  expect(manager.OnNextSlideAnimPlayer).not.toHaveBeenCalled();
  s.slide.timing.getAllAnimEffects = () => s.effects;
  const pending = show.next(),
    rejected = expect(pending).rejects.toThrow('序列已变化');
  s.slide.timing.getAllAnimEffects = () => [];
  vi.advanceTimersByTime(64);
  await rejected;
  expect(error).toHaveBeenCalledOnce();
});

it('starts bound teaching before native clicks and advances only the exact owned ordinal', async () => {
  const s = steps(),
    ready = vi.fn(),
    error = vi.fn(() => show?.dispose());
  show = createTeachingSlideshow({
    doc: document,
    page: 1,
    timelineStepCount: 2,
    onReady: ready,
    onExit: vi.fn(),
    onError: error,
  });
  vi.advanceTimersByTime(32);
  expect(ready).toHaveBeenCalledOnce();
  expect(manager.OnNextSlideAnimPlayer).not.toHaveBeenCalled();
  expect(document.querySelector('[data-native-animation-controls]')).toBeNull();
  await expect(show.next(1)).rejects.toThrow('不一致');
  const first = show.next(0);
  s.state[0] = 'end';
  vi.advanceTimersByTime(32);
  await first;
  vi.advanceTimersByTime(1000);
  expect(s.state[1]).toBe('idle');
  await expect(show.next(0)).rejects.toThrow('不一致');
  const second = show.next(1);
  show.pause();
  s.state[1] = 'end';
  const completed = vi.fn();
  void second.then(completed);
  vi.advanceTimersByTime(32);
  expect(completed).not.toHaveBeenCalled();
  show.resume();
  vi.advanceTimersByTime(32);
  await second;
  expect(ready).toHaveBeenCalledOnce();
  expect(error).not.toHaveBeenCalled();
});
it('rejects mismatched bindings and outside advancement, and blocks canvas clicks while retaining controls', () => {
  const s = steps(),
    error = vi.fn(() => show?.dispose());
  const input = { doc: document, page: 1, onReady: vi.fn(), onExit: vi.fn(), onError: error };
  expect(() => createTeachingSlideshow({ ...input, timelineStepCount: 1 })).toThrow('数量');
  show = createTeachingSlideshow({ ...input, timelineStepCount: 2 });
  const clicked = vi.fn();
  show.canvas.addEventListener('click', clicked);
  const event = new MouseEvent('click', { bubbles: true, cancelable: true });
  show.canvas.dispatchEvent(event);
  expect(clicked).not.toHaveBeenCalled();
  expect(event.defaultPrevented).toBe(true);
  const controls = vi.fn();
  show.controls.addEventListener('click', controls);
  show.controls.click();
  expect(controls).toHaveBeenCalledOnce();
  s.state[0] = 'end';
  vi.advanceTimersByTime(32);
  expect(error).toHaveBeenCalledOnce();
  expect(manager.Mode).toBe(false);
  const after = new MouseEvent('click', { bubbles: true, cancelable: true });
  document.body.dispatchEvent(after);
  expect(after.defaultPrevented).toBe(false);
});

it('blocks native canvas advancement even without bound timeline animations and restores handlers on exit', () => {
  complete = true;
  show = createTeachingSlideshow({ doc: document, page: 1, onReady: vi.fn(), onExit: vi.fn(), onError: vi.fn() });
  const nativeClick = vi.fn();
  show.canvas.addEventListener('mouseup', nativeClick);
  const event = new MouseEvent('mouseup', { bubbles: true, cancelable: true });
  show.canvas.dispatchEvent(event);
  expect(nativeClick).not.toHaveBeenCalled();
  expect(event.defaultPrevented).toBe(true);
  const controlClick = vi.fn();
  show.controls.addEventListener('click', controlClick);
  show.controls.click();
  expect(controlClick).toHaveBeenCalledOnce();
  const escape = new KeyboardEvent('keydown', { key: 'Escape', bubbles: true, cancelable: true });
  show.canvas.dispatchEvent(escape);
  expect(escape.defaultPrevented).toBe(false);
  show.dispose();
  const after = new MouseEvent('mouseup', { bubbles: true, cancelable: true });
  document.body.dispatchEvent(after);
  expect(after.defaultPrevented).toBe(false);
});
