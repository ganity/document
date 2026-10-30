import { getEditorApi } from './agent-plugin/editor-bridge';
import { hasUnsavedChanges } from './unsaved-guard';

/** Read the supported saved main sequence without starting a presentation. */
export function inspectTeachingAnimations(page: number) {
  const api = getEditorApi() as any;
  const manager = api?.WordControl?.DemonstrationManager;
  const slide = api?.WordControl?.m_oLogicDocument?.Slides?.[page - 1];
  if (
    hasUnsavedChanges() ||
    !slide ||
    api.isSlideShow?.() ||
    !manager?.IsMainSeqFinished ||
    !api.asc_registerCallback ||
    !api.asc_unregisterCallback ||
    !api.StartDemonstration ||
    !api.EndDemonstration
  )
    throw new Error('请先保存并重新打开 PPT，再启动教学放映');
  // shortcut: only static-geometry main sequences; add motion/exit projection before widening.
  let effects: any[] = [];
  if (slide.timing) {
    effects = slide.timing.getAllAnimEffects?.();
    const sequences = slide.timing.getEffectsSequences?.();
    if (
      !Array.isArray(effects) ||
      effects.length > 200 ||
      !sequences ||
      slide.timing.bldLst != null ||
      (effects.length
        ? sequences.length !== 1 ||
          sequences[0][0] !== null ||
          sequences[0].length !== effects.length + 1 ||
          effects.some((effect: any, i: number) => sequences[0][i + 1] !== effect)
        : sequences.length !== 0) ||
      effects.some(
        (effect: any) =>
          effect.cTn?.presetClass !== 1 ||
          ![1, 10].includes(effect.cTn.presetID) ||
          effect.cTn.presetSubtype !== 0 ||
          effect.cTn.nodeType !== 2 ||
          typeof effect.isIdle !== 'function' ||
          typeof effect.isAtEnd !== 'function' ||
          typeof effect.getObjectId !== 'function' ||
          slide.cSld.spTree.filter((shape: any) => !shape.group && shape.Id === effect.getObjectId()).length !== 1 ||
          !Number.isInteger(effect.asc_getDuration?.()) ||
          effect.asc_getDuration() < 1 ||
          effect.asc_getDuration() > 5000 ||
          effect.asc_getDelay?.() !== 0,
      )
    )
      throw new Error('本页包含尚未支持同步教学提示的动画，请使用原生放映');
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
  return { api, manager, slide, effects };
}

/** One native page; only native state, never elapsed wall time, completes a click step. */
export function createTeachingSlideshow(input: {
  doc: Document;
  page: number;
  timelineStepCount?: number;
  onReady(): void;
  onExit(): void;
  onError(error: Error): void;
}) {
  const { api, manager, slide, effects } = inspectTeachingAnimations(input.page);
  const bound = input.timelineStepCount !== undefined;
  if (bound && (!effects.length || input.timelineStepCount !== effects.length))
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
  canvas.style.cssText = 'position:absolute;inset:0 0 60px;';
  const controls = input.doc.createElement('nav');
  controls.setAttribute('aria-label', '教学放映控制');
  controls.style.cssText =
    'z-index:50;position:absolute;bottom:0;left:0;right:0;height:60px;display:flex;align-items:center;gap:8px;padding:0 12px;';
  for (const type of ['keydown', 'keyup', 'keypress', 'pointerdown', 'mousedown', 'wheel'])
    controls.addEventListener(type, (event) => event.stopPropagation());
  surface.append(canvas, controls);
  input.doc.body.append(surface);
  let disposed = false,
    ready = false,
    paused = false,
    frame = 0;
  let ownedCompleted = 0;
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
  };
  const next = async (expectedIndex?: number): Promise<void> => {
    if (
      hasUnsavedChanges() ||
      getEditorApi() !== api ||
      !manager.Mode ||
      manager.DemonstrationDiv !== canvas ||
      manager.SlideNum !== input.page - 1 ||
      api.WordControl.m_oLogicDocument.Slides[input.page - 1] !== slide
    )
      return Promise.reject(new Error('原生动画页面已变化，请重新打开教学放映'));
    const current = slide.timing?.getAllAnimEffects() ?? [];
    if (current.length !== effects.length || current.some((effect: any, i: number) => effect !== effects[i]))
      return Promise.reject(new Error('原生动画序列已变化，请保存并重新打开 PPT'));
    if (disposed || (!bound && ready) || paused || pending || manager.Transition.IsPlaying())
      return Promise.reject(new Error('请等待当前动画完成，或先继续播放'));
    const index = effects.findIndex((effect) => !effect.isAtEnd());
    if (bound && (expectedIndex !== index || index !== ownedCompleted))
      return Promise.reject(new Error('原生动画步骤与教学时间轴不一致'));
    if (index < 0 || !effects[index].isIdle() || effects.slice(index + 1).some((effect) => !effect.isIdle()))
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
      if (bound ? !!pending : !ready) manager.StartAnimation(input.page - 1);
    }
  };
  const pauseButton = button('暂停 PPT 动画', pause);
  const resumeButton = button('继续 PPT 动画', resume);
  nextButton.disabled = pauseButton.disabled = resumeButton.disabled = true;
  nativeControls.append(progress, nextButton, pauseButton, resumeButton);
  if (effects.length && !bound) controls.append(nativeControls);
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
    failPending(new Error('原生动画放映已停止'));
    cancelAnimationFrame(frame);
    for (const event of events) api.asc_unregisterCallback(event, changed);
    for (const type of blockedEvents) input.doc.removeEventListener(type, blockAdvance, true);
    if (manager.Mode && manager.DemonstrationDiv === canvas) api.EndDemonstration();
    surface.remove();
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
        hasUnsavedChanges()
      ) {
        input.onExit();
        return;
      }
      if (effects.length) {
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
            effects.some((effect, index) =>
              index < ownedCompleted ? !effect.isAtEnd() : index !== pending?.index && !effect.isIdle(),
            )
          )
            throw new Error('原生动画被教学时间轴之外的操作推进，请重新开始放映');
        }
        if (pending) {
          if (effects.slice(pending.index + 1).some((effect) => !effect.isIdle()))
            throw new Error('动画步骤被其他操作推进，请重新开始教学放映');
          if (!paused && effects[pending.index].isAtEnd()) {
            const finished = pending;
            pending = null;
            if (bound) ownedCompleted++;
            finished.resolve();
          }
        }
      }
      if (
        !ready &&
        !paused &&
        !manager.Transition.IsPlaying() &&
        (bound || manager.IsMainSeqFinished(input.page - 1))
      ) {
        ready = true;
        nativeControls.remove();
        if (!bound) api.DemonstrationPause();
        input.onReady();
      }
      if (!disposed) frame = requestAnimationFrame(check);
    } catch (error) {
      failPending(error instanceof Error ? error : new Error(String(error)));
      input.onError(error instanceof Error ? error : new Error(String(error)));
    }
  };
  try {
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
