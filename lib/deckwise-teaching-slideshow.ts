import { hasNativeTeaching } from './deckwise-native-teaching';
import type { SpeechNativeSchedule } from './deckwise-teaching-native-schedule';
import { getEditorApi } from './agent-plugin/editor-bridge';
import { hasUnsavedChanges } from './unsaved-guard';
import { teachingShapes } from './deckwise-teaching-preview';

export class TeachingAnimationSyncError extends Error {}

/** Read the supported saved main sequence without starting a presentation. */
export function inspectTeachingAnimations(page: number, grouped = false, mode: 'bound' | 'native' = 'bound') {
  const native = hasNativeTeaching();
  const api = getEditorApi() as any;
  const manager = api?.WordControl?.DemonstrationManager;
  const slide = api?.WordControl?.m_oLogicDocument?.Slides?.[page - 1];
  if (
    (hasUnsavedChanges() && !hasNativeTeaching()) ||
    !slide ||
    api.isSlideShow?.() ||
    !manager?.IsMainSeqFinished ||
    !api.asc_registerCallback ||
    !api.asc_unregisterCallback ||
    !api.StartDemonstration ||
    !api.EndDemonstration
  )
    throw new Error('请先保存并重新打开 PPT，再启动教学放映');
  let effects: any[] = [];
  if (slide.timing) {
    effects = slide.timing.getAllAnimEffects?.();
    const sequences = slide.timing.getEffectsSequences?.();
    if (
      !Array.isArray(effects) ||
      effects.length > 200 ||
      effects.some(
        (effect) => !effect.cTn || typeof effect.isIdle !== 'function' || typeof effect.isAtEnd !== 'function',
      )
    )
      throw new Error('Native animation state is unavailable');
    if (
      mode === 'bound' &&
      (!sequences ||
        slide.timing.bldLst != null ||
        (effects.length
          ? sequences.length !== 1 ||
            sequences[0][0] !== null ||
            sequences[0].length !== effects.length + 1 ||
            effects.some((effect: any, i: number) => sequences[0][i + 1] !== effect)
          : sequences.length !== 0) ||
        effects.some(
          (effect: any) =>
            ![
              [1, 1, 0],
              [1, 10, 0],
              [2, 10, 0],
              [0, 26, 0],
              [1, 53, 16],
              [1, 23, 16],
              [0, 1, 2],
              [0, 30, 0],
              [1, 2, 8],
              [1, 2, 2],
              [0, 32, 0],
              ...(native
                ? [
                    [0, 6, 0],
                    [1, 37, 0],
                    [1, 22, 4],
                    [1, 23, 16],
                  ]
                : []),
            ].some(
              ([kind, id, subtype]) =>
                effect.cTn?.presetClass === kind && effect.cTn.presetID === id && effect.cTn.presetSubtype === subtype,
            ) ||
            (effect.cTn.nodeType !== 2 &&
              !(grouped && effect.cTn.nodeType === 7) &&
              !(native && [0, 7].includes(effect.cTn.nodeType))) ||
            typeof effect.isIdle !== 'function' ||
            typeof effect.isAtEnd !== 'function' ||
            typeof effect.getObjectId !== 'function' ||
            teachingShapes<any>(slide.cSld.spTree).filter((shape) => shape.Id === effect.getObjectId()).length !== 1 ||
            !Number.isInteger(effect.asc_getDuration?.()) ||
            effect.asc_getDuration() < 1 ||
            effect.asc_getDuration() > 5000 ||
            (native
              ? !Number.isInteger(effect.asc_getDelay?.()) || effect.asc_getDelay() < 0 || effect.asc_getDelay() > 5000
              : effect.asc_getDelay?.() !== 0) ||
            (effect.asc_getRepeatCount?.() ?? 1000) !== 1000 ||
            effect.asc_getRewind?.(),
        ))
    )
      throw new TeachingAnimationSyncError('本页包含尚未支持同步教学提示的动画，请使用原生放映');
  }
  if (slide.isAdvanceAfterTransition?.()) throw new Error('当前页设置了自动换页，请先关闭自动换页再预览教学');
  const hasMedia = (shapes: any[]): boolean =>
    shapes.some(
      (shape) =>
        shape.getMediaFileName?.() ||
        shape.isVideo?.() ||
        shape.isAudio?.() ||
        (shape.spTree && hasMedia(shape.spTree)),
    );
  // Embedded native media needs its own verified owner/completion adapter.
  if (hasMedia(slide.cSld.spTree)) throw new Error('当前页含原生嵌入媒体，请使用原生放映或教学播放器，避免重复播放');
  if (
    effects.length &&
    (typeof manager.OnNextSlideAnimPlayer !== 'function' || typeof manager.StartAnimation !== 'function')
  )
    throw new Error('当前编辑器不支持逐步播放原生动画');
  const steps: (typeof effects)[] = [];
  for (const effect of effects) {
    if (effect.cTn.nodeType === 2 || ((native || mode === 'native') && !steps.length)) steps.push([effect]);
    else if (!steps.length) throw new Error('Native sequence must start with an on-click effect');
    else steps[steps.length - 1].push(effect);
  }
  return { api, manager, slide, effects, steps };
}

/** One native page; only native state, never elapsed wall time, completes a click step. */
export function createTeachingSlideshow(input: {
  doc: Document;
  page: number;
  timelineStepCount?: number;
  generatedMotions?: boolean;
  schedule?: SpeechNativeSchedule;
  onReady(): void;
  onExit(): void;
  onError(error: Error): void;
}) {
  const bound = input.timelineStepCount !== undefined;
  const schedule = input.schedule;
  const { api, manager, slide, effects, steps } = inspectTeachingAnimations(
    input.page,
    input.generatedMotions,
    bound && !schedule ? 'bound' : 'native',
  );
  if (schedule && schedule.slide !== slide) throw new Error('Scheduled animation slide changed');
  const unboundEffects = schedule
    ? schedule.timing
        .getAllAnimEffects()
        .filter((effect) => !schedule.scheduled.some((entry) => entry.effect === effect))
    : bound
      ? []
      : effects;
  if (bound && !schedule && (!effects.length || input.timelineStepCount !== steps.length))
    throw new Error('教学时间轴与已保存 PPT 动画数量不一致，请重新绑定');
  const surface = input.doc.createElement('section');
  surface.dataset.teachingSlideshow = '';
  surface.dataset.teachingPage = String(input.page);
  surface.style.cssText = 'position:fixed;inset:0;z-index:10000;background:#000;color:#fff;';
  surface.setAttribute('aria-label', '本页教学放映');
  const canvas = input.doc.createElement('div');
  canvas.id = `deckwise-teaching-show-${crypto.randomUUID()}`;
  canvas.dataset.deckwiseTeachingCanvas = '';
  // The native controller keeps this explicit canvas instead of opening its default preview.
  canvas.style.cssText = 'position:absolute;inset:0 0 var(--teaching-controls-height,112px);';
  const controls = input.doc.createElement('nav');
  controls.setAttribute('aria-label', '教学放映控制');
  controls.style.cssText =
    'z-index:50;position:absolute;bottom:0;left:0;right:0;min-height:112px;box-sizing:border-box;display:flex;flex-wrap:wrap;align-content:center;align-items:center;gap:8px;padding:10px 16px;background:#171c25;color:#f8fafc;font:14px/1.4 sans-serif;';
  const style = input.doc.createElement('style');
  style.textContent = `[data-teaching-slideshow] nav button { appearance:none;box-sizing:border-box;min-height:36px;padding:7px 12px;border:1px solid #64748b;border-radius:6px;background:#283548;color:#f8fafc;font:600 14px/20px sans-serif;cursor:pointer;white-space:nowrap; }
    [data-teaching-slideshow] nav button:hover:not(:disabled) { background:#3c4f68; }
    [data-teaching-slideshow] nav button:disabled { background:#202734;color:#94a3b8;border-color:#3f4b5e;cursor:default; }
    [data-teaching-slideshow] nav button:focus-visible,[data-teaching-slideshow] nav input:focus-visible { outline:2px solid #7dd3fc;outline-offset:2px; }
    [data-teaching-slideshow] nav [hidden] { display:none!important; }
    [data-teaching-slideshow] [data-playback-seek] { display:flex;align-items:center;gap:10px;flex:1 1 100%;margin:0;color:#f8fafc;font:14px/20px sans-serif; }
    [data-teaching-slideshow] input[type=range] { flex:1;min-width:60px;accent-color:#7dd3fc; }
    [data-teaching-slideshow] output { color:#f8fafc;white-space:nowrap;font-variant-numeric:tabular-nums; }
    [data-teaching-slideshow] [data-show-status] { overflow:hidden;text-overflow:ellipsis;white-space:nowrap;flex:1;min-width:80px; }`;
  surface.append(style);
  for (const type of ['keydown', 'keyup', 'keypress', 'pointerdown', 'mousedown', 'wheel'])
    controls.addEventListener(type, (event) => event.stopPropagation());
  surface.append(canvas, controls);
  input.doc.body.append(surface);
  let layoutFrame = 0;
  const Observer = input.doc.defaultView?.ResizeObserver;
  const layout = Observer
    ? new Observer(() => {
        // Writing an ancestor during observer delivery makes the SDK treat layout feedback as an editing error.
        layoutFrame ||= requestAnimationFrame(() => {
          layoutFrame = 0;
          const height = `${controls.getBoundingClientRect().height}px`;
          if (surface.style.getPropertyValue('--teaching-controls-height') !== height)
            surface.style.setProperty('--teaching-controls-height', height);
        });
      })
    : undefined;
  layout?.observe(controls);
  let disposed = false,
    ready = false,
    paused = false,
    frame = 0;
  let ownedCompleted = 0;
  const scheduledPending = new Map<number, { resolve(): void; reject(error: Error): void }>();
  const scheduledFired = new Set<number>();
  let pending: { index: number; resolve(): void; reject(error: Error): void } | null = null;
  const nativeControls = input.doc.createElement('span');
  nativeControls.style.cssText = 'display:flex;align-items:center;gap:8px;';
  nativeControls.dataset.nativeAnimationControls = '';
  const progress = input.doc.createElement('span');
  progress.setAttribute('role', 'status');
  const button = (label: string, action: () => void) => {
    const node = input.doc.createElement('button');
    node.type = 'button';
    node.textContent = label;
    node.addEventListener('click', action);
    return node;
  };
  const failPending = (error: Error) => {
    const current = pending;
    pending = null;
    current?.reject(error);
    for (const pending of scheduledPending.values()) pending.reject(error);
    scheduledPending.clear();
  };
  const next = async (expectedIndex?: number): Promise<void> => {
    if (
      (hasUnsavedChanges() && !hasNativeTeaching()) ||
      getEditorApi() !== api ||
      !manager.Mode ||
      manager.DemonstrationDiv !== canvas ||
      manager.SlideNum !== input.page - 1 ||
      api.WordControl.m_oLogicDocument.Slides[input.page - 1] !== slide
    )
      return Promise.reject(new Error('原生动画页面已变化，请重新打开教学放映'));
    if (schedule) {
      if (expectedIndex === undefined && !ready && !paused && schedule.hasUnbound) {
        if (manager.OnNextSlideAnimPlayer() !== true) throw new Error('Native unbound animation did not start');
        return;
      }
      const entry = expectedIndex === undefined ? undefined : schedule.scheduled[expectedIndex];
      if (
        disposed ||
        paused ||
        !ready ||
        !entry ||
        scheduledFired.has(expectedIndex!) ||
        slide.timing !== schedule.timing
      )
        throw new Error('Scheduled native animation is unavailable');
      return new Promise<void>((resolve, reject) => {
        scheduledPending.set(expectedIndex!, { resolve, reject });
        scheduledFired.add(expectedIndex!);
        const player = slide.getAnimationPlayer();
        if (player.onSpClick({ Get_Id: () => entry.token, getMediaFileName: () => null }) !== true) {
          scheduledPending.delete(expectedIndex!);
          reject(new Error('Native animation trigger failed'));
        }
      });
    }
    const current = slide.timing?.getAllAnimEffects() ?? [];
    if (current.length !== effects.length || current.some((effect: any, i: number) => effect !== effects[i]))
      return Promise.reject(new Error('原生动画序列已变化，请保存并重新打开 PPT'));
    if (
      disposed ||
      (!bound && effects.every((effect) => effect.isAtEnd())) ||
      paused ||
      pending ||
      manager.Transition.IsPlaying()
    )
      return Promise.reject(new Error('请等待当前动画完成，或先继续播放'));
    const index = steps.findIndex((group) => group.some((effect) => !effect.isAtEnd()));
    if (bound && (expectedIndex !== index || index !== ownedCompleted))
      return Promise.reject(new Error('原生动画步骤与教学时间轴不一致'));
    if (
      index < 0 ||
      steps[index].some((effect) => !effect.isIdle()) ||
      steps
        .slice(index + 1)
        .flat()
        .some((effect) => !effect.isIdle())
    )
      return Promise.reject(new Error('原生动画已在播放或已完成'));
    return new Promise<void>((resolve, reject) => {
      pending = { index, resolve, reject };
      try {
        // Unlike NextSlide, this finite SDK entry never falls through to advancing the page.
        if (manager.OnNextSlideAnimPlayer() !== true) throw new Error('原生动画步骤未能启动');
      } catch (error) {
        failPending(error instanceof Error ? error : new Error(String(error)));
      }
    });
  };
  const nextButton = button('下一步 PPT 动画', () => {
    void next().catch((error) => {
      if (!disposed) input.onError(error);
    });
  });
  const pause = () => {
    if (!disposed) {
      paused = true;
      api.DemonstrationPause();
    }
  };
  const resume = () => {
    if (!disposed) {
      paused = false;
      // Starting the paused animation player cannot advance to the next slide.
      if (schedule || (!bound && effects.length) || pending) manager.StartAnimation(input.page - 1);
    }
  };
  const pauseButton = button('暂停 PPT 动画', pause);
  const resumeButton = button('继续 PPT 动画', resume);
  nextButton.disabled = pauseButton.disabled = resumeButton.disabled = true;
  nativeControls.append(progress, nextButton, pauseButton, resumeButton);
  if (effects.length && (!bound || schedule?.hasUnbound)) controls.append(nativeControls);
  // The teaching timeline owns all advancement; prevent native canvas/key clicks racing it.
  const blockAdvance = (event: Event) => {
    if (event.type.startsWith('key') && (event as KeyboardEvent).key === 'Escape') return;
    if (event.type.startsWith('key') && controls.contains(event.target as Node)) return;
    if (event.type.startsWith('key') || canvas.contains(event.target as Node)) {
      event.preventDefault();
      event.stopImmediatePropagation();
    }
  };
  const blockedEvents = [
    'keydown',
    'keyup',
    'keypress',
    'pointerdown',
    'pointerup',
    'mousedown',
    'mouseup',
    'click',
    'dblclick',
    'wheel',
    'touchstart',
    'touchend',
  ];
  for (const type of blockedEvents) input.doc.addEventListener(type, blockAdvance, true);
  const changed = () => {
    if (disposed) return;
    if (!manager.Mode) input.onExit();
    // Native NextSlide still paints after notifying observers; never destroy its canvas in that stack.
    else if (manager.SlideNum !== input.page - 1)
      queueMicrotask(() => {
        if (!disposed) input.onExit();
      });
  };
  const events = ['asc_onDemonstrationSlideChanged', 'asc_onEndDemonstration'];
  for (const event of events) api.asc_registerCallback(event, changed);
  const dispose = () => {
    if (disposed) return;
    disposed = true;
    layout?.disconnect();
    cancelAnimationFrame(layoutFrame);
    failPending(new Error('原生动画放映已停止'));
    cancelAnimationFrame(frame);
    for (const event of events) api.asc_unregisterCallback(event, changed);
    for (const type of blockedEvents) input.doc.removeEventListener(type, blockAdvance, true);
    try {
      if (manager.Mode && manager.DemonstrationDiv === canvas) api.EndDemonstration();
    } finally {
      schedule?.restore();
      surface.remove();
    }
  };
  const check = () => {
    if (disposed) return;
    try {
      if (
        !manager.Mode ||
        manager.DemonstrationDiv !== canvas ||
        manager.SlideNum !== input.page - 1 ||
        api.WordControl.m_oLogicDocument.Slides[input.page - 1] !== slide ||
        getEditorApi() !== api ||
        (hasUnsavedChanges() && !hasNativeTeaching())
      ) {
        input.onExit();
        return;
      }
      if (schedule) {
        if (slide.timing !== schedule.timing) throw new Error('Scheduled native animation changed');
        const active = unboundEffects.some((effect) => !effect.isIdle() && !effect.isAtEnd());
        progress.textContent = `PPT 动画 ${unboundEffects.filter((effect) => effect.isAtEnd()).length}/${unboundEffects.length}`;
        nextButton.disabled = paused || ready || active || manager.Transition.IsPlaying();
        pauseButton.disabled = paused;
        resumeButton.disabled = !paused;
        for (const [index, pending] of scheduledPending) {
          if (!paused && schedule.scheduled[index].effect.isAtEnd()) {
            scheduledPending.delete(index);
            pending.resolve();
          }
        }
      }
      if (effects.length && !schedule) {
        const current = slide.timing.getAllAnimEffects();
        if (current.length !== effects.length || current.some((effect: any, i: number) => effect !== effects[i]))
          throw new Error('原生动画序列已变化，请保存并重新打开 PPT');
        const completed = effects.filter((effect) => effect.isAtEnd()).length;
        const active = effects.some((effect) => !effect.isIdle() && !effect.isAtEnd());
        progress.textContent = `PPT 动画 ${completed}/${effects.length}${paused ? ' · 已暂停' : active ? ' · 播放中' : ''}`;
        nextButton.disabled =
          paused || active || !!pending || completed === effects.length || manager.Transition.IsPlaying();
        pauseButton.disabled = paused || manager.Transition.IsPlaying();
        resumeButton.disabled = !paused;
        if (bound) {
          if (
            steps.some((group, index) =>
              group.some((effect) =>
                index < ownedCompleted ? !effect.isAtEnd() : index !== pending?.index && !effect.isIdle(),
              ),
            )
          )
            throw new Error('原生动画被教学时间轴之外的操作推进，请重新开始放映');
        }
        if (pending) {
          if (
            steps
              .slice(pending.index + 1)
              .flat()
              .some((effect) => !effect.isIdle())
          )
            throw new Error('动画步骤被其他操作推进，请重新开始教学放映');
          if (!paused && steps[pending.index].every((effect) => effect.isAtEnd())) {
            const finished = pending;
            pending = null;
            if (bound) ownedCompleted++;
            finished.resolve();
          }
        }
      }
      // Automatic effects keep their native delays; only idle click effects need a synthetic click.
      if (
        !paused &&
        !pending &&
        !manager.Transition.IsPlaying() &&
        unboundEffects.some(
          (effect, index) =>
            effect.cTn.nodeType === 2 &&
            effect.isIdle() &&
            !effect.isAtEnd() &&
            unboundEffects.slice(0, index).every((previous) => previous.isAtEnd()),
        ) &&
        unboundEffects.every((effect) => effect.isIdle() || effect.isAtEnd())
      ) {
        if (manager.OnNextSlideAnimPlayer() !== true) throw new Error('Native page-entry animation did not start');
      }
      if (!ready && !paused && !manager.Transition.IsPlaying()) {
        ready = true;
        if (bound && !schedule?.hasUnbound) nativeControls.remove();
        input.onReady();
      }
      if (unboundEffects.every((effect) => effect.isAtEnd())) nativeControls.remove();
      if (!disposed) frame = requestAnimationFrame(check);
    } catch (error) {
      failPending(error instanceof Error ? error : new Error(String(error)));
      input.onError(error instanceof Error ? error : new Error(String(error)));
    }
  };
  try {
    schedule?.install();
    api.StartDemonstration(canvas.id, input.page - 1);
    if (!manager.Mode || manager.DemonstrationDiv !== canvas) throw new Error('原生放映未能启动');
    frame = requestAnimationFrame(check);
  } catch (error) {
    dispose();
    throw error;
  }
  return {
    surface,
    canvas,
    controls,
    next,
    pause,
    resume,
    dispose,
  };
}
