import { assertNativeTeachingAnimations } from './deckwise-native-teaching';
import { hasNativeTeaching, nativeTeachingObjects, readNativeTeaching, replaceNativeTeachingPage } from './deckwise-native-teaching';
import { routeNativeTeaching, nativeTeachingHistory } from './deckwise-native-teaching-router';
import { teachingObjectChoices, appendTeachingObjectChoice } from './deckwise-teaching-object-choice';
import { createTeachingAnimationForm } from './deckwise-teaching-animation-form';
import {
  listTeachingNativeAnimations,
  relocateAnimationTrigger,
  type AnimationCue,
} from './deckwise-teaching-animation-bindings';
import { prepareSpeechNativeSchedule } from './deckwise-teaching-native-schedule';
import { t } from '@ranuts/shared/i18n';
import {
  speechPauseMode,
  speechPauseLabel,
  teachingStepTitle,
  teachingStepSummary,
  teachingFlowStyles,
  decorateTeachingButton,
  createTeachingPanelLayout,
} from './deckwise-teaching-flow';
import { hasUnsavedChanges } from './unsaved-guard';
import { createSidebarEntry } from './agent-plugin/ui/sidebar-entry';
import { mountPanelContext } from './agent-plugin/ui/panel-context';
import { deckwiseEditorBridge } from './deckwise-editor-bridge';
import { getEditorApi } from './agent-plugin/editor-bridge';
import { deckwisePageLineage } from './deckwise-page-lineage';
import { integrateGeneratedNativeMotions, nativeSpeechMotionSteps } from './deckwise-teaching-native-motion';
import { createTeachingSlideshow, inspectTeachingAnimations } from './deckwise-teaching-slideshow';
import {
  createTeachingAttentionPreview,
  attentionTarget,
  selectedAttentionTargets,
  loadTeachingPreviewRuntime,
  type TimelinePreview,
  type AttentionCue,
} from './deckwise-teaching-preview';

type Action = { id: string; type: string; [key: string]: any };
type View = {
  teaching?: { pageType: string; phase: string; objective: string; durationSeconds: number };
  saveWarnings?: string[];
  teachingObjects?: Array<{
    objectId: string;
    name: string;
    role?: string;
    aliases?: string[];
    status: string;
    effectTargetEligible?: boolean;
    sourceShapeId?: string;
    sourceObjectName?: string;
  }>;
  viewId: string;
  pageId: string;
  pageNumber: number;
  project: {
    taskType?: string;
    teachingRevision: string;
    source?: { sha256: string };
    pages: Array<{ pageId: string; locked: boolean; reviewState?: string; actions: Action[] }>;
  };
};
type Context = ReturnType<typeof deckwiseEditorBridge.getContext>;
type Pending = { resolve(value: any): void; reject(error: Error): void; timer: ReturnType<typeof setTimeout> };
type Media = { kind: 'music' | 'video'; intentId: string; title: string; model: string; inPage: boolean };
type Model = { model: string; voices: string[]; providerDefaultVoice: boolean };
type LessonRun = {
  cover: HTMLElement;
  page: number;
  total: number;
  navigating: boolean;
  targetPage?: number;
  finishPage?: (error?: Error) => void;
};

/** UI only: authoritative teaching, drafts and paid media operations stay with Host. */
export function mountDeckwiseTeachingSidebar(binding: any, send: (type: string, payload: any, id?: string) => void) {
  let lifecycle: ReturnType<typeof mountPanelContext> | undefined;
  let disposed = false,
    open = true,
    busy = false,
    running = false;
  let nativeSlideshow = false;
  let projectingMotions = false;
  const projectedSources = new Set<string>();
  let context: Context | null = null,
    view: View | null = null;
  let actions: Action[] = [],
    baseline = '',
    draftRevision: string | null = null;
  let draftQueue = Promise.resolve(),
    draftError = '',
    draftInstruction = '',
    epoch = 0;
  const undoActions: string[] = [],
    redoActions: string[] = [];
  let readQueue = Promise.resolve();
  let renderedEditable = false;
  let selectedActionId: string | null = null;
  let draggedActionId: string | null = null;
  let statusText = '',
    statusError = false;
  let models: Model[] = [],
    tasks: any[] = [];
  let panel: HTMLElement | null = null,
    nativeDocument: Document | null = null;
  let panelLayout: ReturnType<typeof createTeachingPanelLayout> | undefined;
  let previewAudio: HTMLMediaElement | null = null,
    previewUrl: string | null = null;
  let mediaLibrary: Media[] = [];
  let mediaAdoption: { key: string; actionId: string } | null = null;
  let saveTimer: ReturnType<typeof setTimeout> | undefined;
  let statusTimer: ReturnType<typeof setTimeout> | undefined;
  let attention: Awaited<ReturnType<typeof createTeachingAttentionPreview>> | null = null;
  let scheduler: InstanceType<
    Awaited<ReturnType<typeof createTeachingAttentionPreview>>['shared']['SpeechCueScheduler']
  > | null = null;
  let timeline: TimelinePreview | null = null,
    previewMode = 'idle';
  let previewVideo: HTMLVideoElement | null = null;
  const timelineUrls: string[] = [];
  let timelineRun = 0;
  let previewRequest = 0;
  let restoreFocus: (() => void) | undefined;
  let lesson: LessonRun | null = null;
  let teachingSlideshow: ReturnType<typeof createTeachingSlideshow> | null = null;
  let startingSlideshow = false;
  const stopAudio = (preserveLesson?: LessonRun) => {
    if (lesson && lesson !== preserveLesson) {
      const cancelled = lesson;
      lesson = null;
      cancelled.cover.remove();
      cancelled.finishPage?.(new Error('整课教学放映已停止'));
    }
    timelineRun++;
    timeline?.dispose();
    timeline = null;
    const show = teachingSlideshow;
    teachingSlideshow = null;
    const wasStarting = startingSlideshow;
    startingSlideshow = true;
    try {
      show?.dispose();
    } finally {
      startingSlideshow = wasStarting;
    }
    previewVideo?.remove();
    previewVideo = null;
    previewMode = 'idle';
    timelineUrls.splice(0).forEach((url) => URL.revokeObjectURL(url));
    panel?.querySelector('[data-teaching-question]')?.remove();
    panel?.querySelectorAll('[data-playing]').forEach((node) => node.removeAttribute('data-playing'));
    renderControls();
    scheduler?.stop();
    scheduler = null;
    attention?.dispose();
    attention = null;
    if (previewAudio) {
      previewAudio.pause();
      previewAudio.removeAttribute('src');
      previewAudio.load();
      previewAudio.remove();
    }
    previewAudio = null;
    if (previewUrl) URL.revokeObjectURL(previewUrl);
    previewUrl = null;
  };
  const pending = new Map<string, Pending>();
  const localDrafts = new Map<string, { actions: Action[]; baseline: string }>();
  let stopNativeClicks: (() => void) | undefined;
  const page = () => view?.project.pages.find((entry) => entry.pageId === view?.pageId);
  const changed = () => JSON.stringify(actions) !== baseline;
  const editable = () => !!view && !page()?.locked && !busy && !draftError && !lesson;
  const request = async (operation: string, args: any = {}, target = context): Promise<any> => {
    if (!target || disposed) throw new Error('Teaching context is unavailable');
    if (hasNativeTeaching()) {
      const native = await routeNativeTeaching(operation, args, target.page ?? 1);
      if (native && !(native.value as { nativeService?: boolean })?.nativeService) return native.value;
    }
    return new Promise<any>((resolve, reject) => {
      if (!target || disposed) return reject(new Error('Teaching context is unavailable'));
      const id = crypto.randomUUID();
      const timer = setTimeout(() => {
        pending.delete(id);
        reject(new Error('Teaching request timed out; check its status before retrying generation'));
      }, 60_000);
      pending.set(id, { resolve, reject, timer });
      send('document:teaching-request', { operation, args, context: target }, id);
    });
  };
  const scope = () => ({ viewId: view!.viewId, expectedTeachingRevision: view!.project.teachingRevision });
  const element = <K extends keyof HTMLElementTagNameMap>(tag: K, text?: string) => {
    const node = nativeDocument!.createElement(tag);
    if (text) node.textContent = text;
    return node;
  };
  const notice = (text: string, error = false) => {
    statusText = text;
    statusError = error;
    const showStatus = teachingSlideshow?.controls.querySelector('[data-show-status]');
    if (showStatus) showStatus.textContent = lesson ? `第 ${lesson.page}/${lesson.total} 页 · ${text}` : text;
    if (!panel) return;
    const node = panel.querySelector<HTMLElement>('[data-status]');
    if (node) {
      node.textContent = text;
      node.title = text;
      node.setAttribute('role', error ? 'alert' : 'status');
    }
  };
  const button = (text: string, action: () => void, disabled = false) => {
    const node = element('button', text);
    node.type = 'button';
    node.disabled = disabled;
    node.addEventListener('click', action);
    return node;
  };
  const run = async (operation: () => Promise<void>, lockPanel = true) => {
    if (busy || running) return;
    const token = epoch;
    running = true;
    if (panel && lockPanel) panel.inert = true;
    try {
      await operation();
    } catch (error) {
      if (token === epoch) notice(String(error), true);
    } finally {
      running = false;
      if (panel && lockPanel) panel.inert = false;
      restoreFocus?.();
      restoreFocus = undefined;
      if (token === epoch) renderControls();
    }
  };
  function renderPreviewControls() {
    const controls = [
      ...(panel?.querySelectorAll<HTMLButtonElement>('[data-preview-command]') ?? []),
      ...(teachingSlideshow?.controls.querySelectorAll<HTMLButtonElement>('[data-preview-command]') ?? []),
    ];
    controls.forEach((node) => {
      node.hidden = ['pause', 'resume'].includes(node.dataset.previewCommand ?? '')
        ? !timeline
        : node.dataset.previewCommand === 'teacher' && previewMode !== 'waiting_interaction';
      switch (node.dataset.previewCommand) {
        case 'lesson':
        case 'show':
        case 'play':
          node.disabled =
            !view || busy || running || !!draftError || !!timeline || !actions.length || nativeSlideshow || !!lesson;
          break;
        case 'toggle':
          node.textContent = previewMode === 'paused' ? '继续播放' : '暂停播放';
          node.title = node.textContent;
          node.disabled = !['playing', 'paused'].includes(previewMode);
          break;
        case 'pause':
          node.disabled = previewMode !== 'playing';
          break;
        case 'resume':
          node.disabled = previewMode !== 'paused';
          break;
        case 'teacher':
          node.disabled = previewMode !== 'waiting_interaction';
          break;
      }
    });
  }
  function renderControls() {
    panelLayout?.sync(open);
    renderPreviewControls();
    panel?.querySelectorAll<HTMLButtonElement>('[data-teaching-history]').forEach((node) => {
      const entries = node.dataset.teachingHistory === 'undo' ? undoActions : redoActions;
      node.disabled = !editable() || running || (!hasNativeTeaching() && !entries.length);
    });
    panel?.querySelectorAll<HTMLButtonElement>('[data-adopt-media]').forEach((node) => {
      const select = panel?.querySelector<HTMLSelectElement>('[data-media-library]');
      node.disabled = !editable() || changed() || running || !!timeline || !select?.value;
    });
    panel
      ?.querySelectorAll<HTMLButtonElement>('[data-save]')
      .forEach((node) => (node.disabled = !editable() || !changed()));
  }
  const replaceNativeTeachingPageForSidebar = (current: NonNullable<ReturnType<typeof readNativeTeaching>>) =>
    replaceNativeTeachingPage({ documentId: current.documentId, expectedRevision: view!.project.teachingRevision, pageId: view!.pageId, actions });
  const persistDraft = () => {
    if (!view || !context) return;
    if (hasNativeTeaching()) {
      const next = readNativeTeaching()!;
      const updated = replaceNativeTeachingPageForSidebar(next);
      view.project.pages = updated.pages;
      view.project.teachingRevision = updated.revision;
      actions = structuredClone(updated.pages.find(p => p.pageId === view!.pageId)!.actions);
      baseline = JSON.stringify(actions);
      notice('教学修改已进入原生保存流程');
      renderControls();
      return;
    }
    localDrafts.set(view.pageId, { actions: structuredClone(actions), baseline: JSON.stringify(view.project) });
    const args = { viewId: view.viewId },
      target = context,
      content = structuredClone(actions);
    draftQueue = draftQueue
      .then(async () => {
        const result = await request(
          'draft:set',
          { ...args, expectedRevision: draftRevision, draft: { actions: content, instruction: draftInstruction } },
          target,
        );
        draftRevision = result.draft?.revision ?? null;
      })
      .catch((error) => {
        draftError = String(error);
        notice(`草稿未保存：${draftError}。请保留当前编辑器。`, true);
      });
    notice('教学已修改，正在自动保存…');
    scheduleSave();
    renderControls();
  };
  const remember = (entries: string[], content: string) => {
    entries.push(content);
    // shortcut: draft undo is bounded to 20 edits and 4 Mi UTF-16 units; durable history belongs to Host.
    while (entries.length > 20 || entries.reduce((size, entry) => size + entry.length, 0) > 4 * 1024 * 1024)
      entries.shift();
  };
  const edit = (next: Action[], redraw = true) => {
    if (!editable() || JSON.stringify(next) === JSON.stringify(actions)) return;
    const native = next.filter((action) => action.type === 'native_animation');
    if (native.some((action, index) => action.stepIndex !== index)) {
      notice('请保持 PPT 动画的原始顺序；可将讲解、音乐、停顿移到动画之间。删除绑定请使用「移除动画绑定」。', true);
      return;
    }
    stopAudio();
    if (!hasNativeTeaching()) remember(undoActions, JSON.stringify(actions));
    redoActions.length = 0;
    actions = next;
    persistDraft();
    if (redraw) render();
  };
  const changeHistory = (undo: boolean) => {
    if (hasNativeTeaching()) {
      if (!view || page()?.locked || running || lesson || draftError) return;
      stopAudio();
      nativeTeachingHistory(undo);
      void load(deckwiseEditorBridge.getContext(binding));
      return;
    }
    if (!editable() || running) return;
    const from = undo ? undoActions : redoActions,
      to = undo ? redoActions : undoActions;
    if (!from.length) return;
    stopAudio();
    remember(to, JSON.stringify(actions));
    actions = JSON.parse(from.pop()!);
    persistDraft();
    render();
    notice(undo ? '已撤销未保存的教学修改' : '已重做未保存的教学修改');
  };
  const applyView = async (next: View, preserveLesson?: LessonRun, preserveHistory = false) => {
    if (!preserveHistory) undoActions.length = redoActions.length = 0;
    if (view?.pageId !== next.pageId) selectedActionId = null;
    view = next;
    actions = structuredClone(page()!.actions);
    baseline = JSON.stringify(actions);
    const recovered = await request('draft:get', { viewId: view.viewId });
    if (view !== next) throw new Error('Teaching page changed while loading its draft');
    draftRevision = recovered.draft?.revision ?? null;
    draftInstruction = recovered.draft?.instruction ?? '';
    draftError = '';
    if (recovered.draft) {
      if (JSON.stringify(recovered.draft.baseline) !== JSON.stringify(next.project)) {
        draftError = '未保存修改与当前版本冲突。可删除草稿，继续使用已保存内容。';
      } else {
        actions = recovered.draft.actions;
        draftError = '';
        localDrafts.delete(next.pageId);
      }
    }
    const local = localDrafts.get(next.pageId);
    if (local) {
      if (local.baseline === JSON.stringify(next.project)) actions = local.actions;
      else draftError = '未保存修改与当前版本冲突。可删除草稿，继续使用已保存内容。';
    }
    render(preserveLesson);
  };
  // Selection and save callbacks can renew the binding while the Host parses saved bytes.
  const openTeaching = async (target: Context, isCurrent = () => true) => {
    for (let attempt = 0; ; attempt++) {
      let data: View | undefined, failure: unknown;
      try {
        data = await request('open', {}, target);
      } catch (error) {
        failure = error;
      }
      if (disposed) throw new Error('Teaching editor closed');
      // A queued read already owns the newer context; do not retry it from the obsolete read too.
      if (!isCurrent()) throw new Error('Teaching page superseded');
      const current = deckwiseEditorBridge.getContext(binding);
      if (
        attempt < 2 &&
        current.page === target.page &&
        !current.readOnly &&
        current.editGeneration !== target.editGeneration
      ) {
        target = current;
        continue;
      }
      if (failure) throw failure;
      return { data: data!, target };
    }
  };
  const load = async (next: Context, preserveLesson?: LessonRun) => {
    const samePage = !!view && view.pageNumber === next.page;
    const selectionRefresh = samePage && context?.editGeneration !== next.editGeneration;
    clearTimeout(saveTimer);
    const token = ++epoch;
    if (!samePage) {
      undoActions.length = redoActions.length = 0;
      tasks = [];
      models = [];
      mediaLibrary = [];
      mediaAdoption = null;
      notice('正在读取本页讲解…');
    }
    stopAudio(preserveLesson);
    context = next;
    busy = true;
    if (!view) notice('正在同步 PPT 并读取本页教学…');
    await draftQueue;
    try {
      const opening = readQueue.then(() => {
        if (disposed || token !== epoch) throw new Error('Teaching page superseded');
        return openTeaching(next, () => token === epoch);
      });
      readQueue = opening.then(
        () => {},
        () => {},
      );
      const { data, target } = await opening;
      if (disposed || token !== epoch) return;
      // Do not redirect an in-flight read to a newly selected slide.
      const current = deckwiseEditorBridge.getContext(binding);
      if (current.page !== target.page || current.editGeneration !== target.editGeneration) {
        busy = false;
        queueMicrotask(() => lifecycle?.update());
        return;
      }
      context = target;
      // Selection changes renew the Host binding, not the authored timeline or its form.
      if (
        selectionRefresh &&
        renderedEditable &&
        !lesson &&
        view?.project.source?.sha256 &&
        data.pageId === view.pageId &&
        JSON.stringify({ ...data, viewId: undefined }) === JSON.stringify({ ...view, viewId: undefined })
      ) {
        view = data;
        busy = false;
        renderControls();
        return;
      }
      if (hasNativeTeaching() && samePage && view?.project.teachingRevision === data.project.teachingRevision) {
        view = data;
        busy = false;
        renderControls();
        return;
      }
      await applyView(data, preserveLesson);
      busy = false;
      statusText = [changed() ? '已恢复本页未保存草稿' : '', ...(data.saveWarnings ?? [])].filter(Boolean).join('；');
      statusError = false;
      render(preserveLesson);
      if (changed() && !draftError) scheduleSave();
      if (
        !preserveLesson &&
        !changed() &&
        data.project.taskType === 'scene' &&
        actions.some(
          (a) =>
            a.type === 'speech' &&
            a.visualCues?.some(
              (c: AttentionCue) => c.effect?.type === 'element_animation' && !c.trigger && !c.nativeAnimation,
            ),
        )
      ) {
        const key = `${data.project.source?.sha256}/${data.pageId}/${data.project.teachingRevision}`;
        if (!projectedSources.has(key)) {
          projectingMotions = true;
          try {
            const shared = await loadTeachingPreviewRuntime();
            if (disposed || token !== epoch || changed() || view?.viewId !== data.viewId) return;
            projectedSources.add(key);
            if (integrateGeneratedNativeMotions(data.pageNumber, actions, data.teachingObjects ?? [], shared)) {
              const currentContext = deckwiseEditorBridge.getContext(binding);
              const { data: saved, target } = await openTeaching(currentContext);
              if (disposed || token !== epoch) return;
              if (deckwiseEditorBridge.getContext(binding).page !== data.pageNumber)
                throw new Error('Teaching page changed during native animation save');
              context = target;
              await applyView(saved);
              notice('教学动画已写入 PPT，可在「动画」菜单编辑。');
            }
          } catch (error) {
            notice(`PPT 动画同步未完成：${String(error)}`, true);
          } finally {
            projectingMotions = false;
            if (!disposed) lifecycle?.update();
          }
        }
      }
      const owner = view!.viewId;
      const results = await Promise.allSettled([
        request('tts:models', scope()),
        request('tts:list', scope()),
        request('media:list', scope()),
      ]);
      if (disposed || token !== epoch || view?.viewId !== owner) return;
      if (results[0].status === 'fulfilled') models = results[0].value;
      if (results[1].status === 'fulfilled') tasks = results[1].value;
      if (results[2].status === 'fulfilled' && Array.isArray(results[2].value)) mediaLibrary = results[2].value;
      if (preserveLesson && lesson === preserveLesson) render(preserveLesson);
      else if (
        !lesson &&
        !running &&
        !timeline &&
        !previewAudio &&
        !teachingSlideshow &&
        !changed() &&
        !panel?.contains(nativeDocument?.activeElement ?? null)
      )
        render();
      if (results.slice(0, 2).some((result) => result.status === 'rejected'))
        notice('本页讲解已读取，语音服务暂不可用；可稍后重试。', true);
      else if (results[2].status === 'rejected') notice('本页讲解已读取，已有素材列表暂不可用，可点击刷新重试。', true);
    } catch (error) {
      if (token === epoch) {
        busy = false;
        view = null;
        render(preserveLesson);
        notice(String(error), true);
      }
    }
  };
  const saveCurrent = async () => {
    clearTimeout(saveTimer);
    if (!view || !context) throw new Error('教学页面尚未就绪');
    if (draftError) throw new Error(draftError);
    if (!changed()) return;
    if (page()?.locked) throw new Error('本页已锁定，无法保存教学修改');
    const target = context,
      currentView = view,
      content = structuredClone(actions),
      token = epoch;
    const active = nativeDocument?.activeElement;
    const narration = active?.tagName === 'TEXTAREA' ? (active as HTMLTextAreaElement) : null;
    const actionId = narration?.closest<HTMLElement>('[data-action-id]')?.dataset.actionId;
    const selection = narration
      ? ([narration.selectionStart, narration.selectionEnd, narration.selectionDirection] as const)
      : null;
    const scrollTop = panel?.scrollTop ?? 0;
    // Preview preparation normally leaves Stop available; committing must also block edits.
    const savingPanel = panel;
    if (savingPanel) savingPanel.inert = true;
    try {
      notice('正在保存本页教学…');
      await draftQueue;
      if (draftError) throw new Error(draftError);
      if (token !== epoch) throw new Error('教学页面已切换，请在当前页重试');
      const next = await request('write', { ...scope(), actions: content }, target);
      if (token !== epoch) throw new Error('教学页面已切换，请在当前页重试');
      // A pending field blur may arrive while the previous snapshot is being saved.
      const retainNewerEdits = () => {
        if (JSON.stringify(actions) === JSON.stringify(content)) return false;
        view = next;
        baseline = JSON.stringify(page()!.actions);
        persistDraft();
        return true;
      };
      await draftQueue;
      if (token !== epoch) throw new Error('教学页面已切换，请在当前页重试');
      if (retainNewerEdits()) return;
      const clearing = draftQueue.then(async () => {
        const result = await request(
          'draft:set',
          { viewId: currentView.viewId, expectedRevision: draftRevision, draft: null },
          target,
        );
        draftRevision = result.draft?.revision ?? null;
      });
      // Draft edits arriving during clear must use the revision returned by that clear.
      draftQueue = clearing.then(
        () => {},
        () => {},
      );
      await clearing;
      await draftQueue;
      if (token !== epoch) throw new Error('教学页面已切换，请在当前页重试');
      if (retainNewerEdits()) return;
      localDrafts.delete(currentView.pageId);
      await applyView(next, undefined, true);
      if (panel) panel.scrollTop = scrollTop;
      if (actionId && selection)
        restoreFocus = () => {
          if (
            token !== epoch ||
            disposed ||
            (nativeDocument?.activeElement && panel?.contains(nativeDocument.activeElement))
          )
            return;
          const input = [...(panel?.querySelectorAll<HTMLElement>('[data-teaching-inspector]') ?? [])]
            .find((card) => card.dataset.actionId === actionId)
            ?.querySelector('textarea');
          input?.focus({ preventScroll: true });
          input?.setSelectionRange(...selection);
        };
      notice('本页教学已保存');
    } finally {
      if (savingPanel) savingPanel.inert = false;
    }
  };
  const save = () => void run(saveCurrent, false);
  const discardChanges = async () => {
    if (!view || !context) return;
    clearTimeout(saveTimer);
    const current = view,
      target = context,
      token = epoch;
    await draftQueue;
    if (disposed || token !== epoch) return;
    const result = await request(
      'draft:set',
      { viewId: current.viewId, expectedRevision: draftRevision, draft: null },
      target,
    );
    if (disposed || token !== epoch) return;
    stopAudio();
    localDrafts.delete(current.pageId);
    draftRevision = result.draft?.revision ?? null;
    draftError = draftInstruction = '';
    undoActions.length = redoActions.length = 0;
    actions = structuredClone(page()!.actions);
    baseline = JSON.stringify(actions);
    notice('已删除草稿，当前显示已保存内容');
    render();
  };
  function scheduleSave() {
    clearTimeout(saveTimer);
    saveTimer = setTimeout(() => {
      if (disposed || !changed() || draftError || page()?.locked) return;
      if (
        busy ||
        running ||
        timeline ||
        previewAudio ||
        attention ||
        teachingSlideshow ||
        panel?.querySelector('[data-cue-form]:not([hidden])')
      ) {
        scheduleSave();
        return;
      }
      save();
    }, 1200);
  }
  const pollTasks = () => {
    clearTimeout(statusTimer);
    statusTimer = setTimeout(async () => {
      if (disposed || !open) return;
      if (!view || busy || running || changed() || previewAudio || timeline) {
        pollTasks();
        return;
      }
      const token = epoch;
      const waiting = tasks.filter(
        (task) =>
          task.phase === 'submission-unknown' ||
          (task.phase === 'submitted' && !['ready', 'failed', 'cancelled'].includes(task.task?.state)),
      );
      try {
        for (const task of waiting) {
          const next = await request('tts:status', { ...scope(), intentId: task.intentId });
          if (token !== epoch || disposed) return;
          tasks = tasks.map((entry) => (entry.intentId === next.intentId ? next : entry));
        }
        if (!lesson && waiting.length && !panel?.contains(nativeDocument?.activeElement ?? null)) render();
      } catch (error) {
        if (token === epoch) notice(`语音状态检查失败：${String(error)}`, true);
      } finally {
        if (!disposed) pollTasks();
      }
    }, 10_000);
  };
  const refresh = () => {
    try {
      void load(deckwiseEditorBridge.getContext(binding));
    } catch (error) {
      notice(String(error), true);
    }
  };

  const previewAttention = async (isCurrent = () => true, slideshow?: HTMLElement) => {
    if (nativeSlideshow && !slideshow) throw new Error('请先退出原生幻灯片放映，再预览教学提示');
    const runId = timelineRun;
    if (!view?.project.source?.sha256 || !nativeDocument) throw new Error('请重新读取本页教学以获取当前来源');
    const token = epoch;
    const next = await createTeachingAttentionPreview({
      doc: nativeDocument,
      page: view.pageNumber,
      sourceSha256: view.project.source.sha256,
      objects: view.teachingObjects ?? [],
      slideshow,
      onError: (error) => {
        stopAudio();
        notice(error.message, true);
      },
    });
    if (disposed || token !== epoch || runId !== timelineRun || (nativeSlideshow && !slideshow) || !isCurrent()) {
      next.dispose();
      throw new Error('教学预览页面已变化');
    }
    attention = next;
    return next;
  };

  const playTimeline = (inSlideshow = false, preserveLesson?: LessonRun) =>
    run(async () => {
      if (nativeSlideshow) throw new Error('请先退出原生幻灯片放映，再预览教学时间轴');
      const requestId = previewRequest;
      if (changed()) await saveCurrent();
      if (requestId !== previewRequest || nativeSlideshow || disposed) return;
      if (!view || draftError) throw new Error('教学页面尚未就绪');
      const content = structuredClone(actions),
        token = epoch;
      if (
        content.some(
          (node) =>
            !['speech', 'audio_clip', 'video_clip', 'delay', 'native_animation'].includes(node.type) &&
            !(
              node.type === 'interaction_checkpoint' &&
              ['teacher_continue', 'chapter_summary'].includes(node.checkpointKind)
            ),
        )
      )
        throw new Error('本页包含尚未接入的互动或未生成的视频，请先完成准备或使用高级教学工具中的教学播放器');
      for (const entry of content)
        for (const cue of entry.visualCues ?? []) {
          if (
            cue.trigger?.kind === 'text' &&
            (cue.trigger.stale ||
              cue.trigger.audioAssetId !== entry.audioAssetId ||
              entry.text.slice(cue.trigger.start, cue.trigger.end) !== cue.trigger.text)
          )
            throw new Error(t('teachingTimingMarkRequired'));
        }
      const nativeNodes = content.filter((node) => node.type === 'native_animation');
      if (!inSlideshow && hasNativeTeaching() && (nativeNodes.length || content.some(node =>
        node.visualCues?.some((cue: AttentionCue) => cue.effect?.type === 'element_animation')))) inSlideshow = true;
      if (nativeNodes.length && !inSlideshow)
        throw new Error('本页已绑定 PPT 动画，请使用「幻灯片放映」预览完整时间轴');
      if (
        nativeNodes.some(
          (node, index) =>
            node.version !== 1 || node.stepIndex !== index || node.sourceSha256 !== view!.project.source?.sha256,
        )
      )
        throw new Error('PPT 动画绑定已过期，请重新绑定已保存的当前页');
      if (nativeNodes.length) {
        assertNativeTeachingAnimations(view!.pageNumber);
        if (deckwisePageLineage.context()?.sourceSha256 !== view!.project.source?.sha256)
          throw new Error('请保存并重新打开当前 PPT 后放映');
        const { effects, slide } = inspectTeachingAnimations(view!.pageNumber);
        if (effects.length !== nativeNodes.length) throw new Error('PPT 动画步骤数量已变化，请重新绑定');
        const nativeTargets = new Set(effects.map((effect) => effect.getObjectId()));
        const revealed = new Set<string>();
        for (const action of content) {
          if (action.type === 'native_animation') revealed.add(effects[action.stepIndex].getObjectId());
          for (const cue of action.visualCues ?? []) {
            const effect = cue.effect;
            if (!effect || !['highlight', 'spotlight', 'element_animation'].includes(effect.type)) continue;
            const ids =
              effect.type === 'highlight' ? effect.elementIds : [effect.elementId, ...(effect.elementIds ?? [])];
            for (const id of ids) {
              const target = attentionTarget(slide.cSld.spTree, view!.teachingObjects ?? [], id);
              if (!target.Id) throw new Error('教学提示目标缺少原生对象标识');
              if (
                nativeTargets.has(target.Id) &&
                (effect.type === 'element_animation' || !revealed.has(target.Id))
              )
                throw new Error('请将高亮或聚光讲解放到对象的 PPT 入场动画之后；同一对象不能重复添加教学动画');
            }
          }
        }
      }
      stopAudio(preserveLesson);
      const runId = timelineRun;
      const isCurrent = () => token === epoch && runId === timelineRun && !disposed;
      notice('正在读取本页已有讲解与媒体…');
      let schedule: ReturnType<typeof prepareSpeechNativeSchedule>;
      try {
        const shared = await loadTeachingPreviewRuntime();
        if (!isCurrent()) return;
        schedule = prepareSpeechNativeSchedule(view!.pageNumber, content, view!.teachingObjects ?? [], shared);
        if (schedule && !inSlideshow) throw new Error(t('teachingAnimationSlideshowRequired'));
        const motionSteps =
          schedule?.motions ??
          (inSlideshow
            ? nativeSpeechMotionSteps(view!.pageNumber, content, view!.teachingObjects ?? [], shared)
            : undefined) ??
          [];
        const stepCount = nativeNodes.length || motionSteps.length;
        const nodes: Array<Record<string, unknown>> = [];
        let bytesTotal = 0;
        for (const action of content) {
          if (action.type === 'native_animation') {
            nodes.push(action);
            continue;
          }
          if (action.type === 'delay') {
            nodes.push({ id: action.id, type: 'delay', durationMs: action.durationMs });
            continue;
          }
          if (action.type === 'interaction_checkpoint') {
            nodes.push(structuredClone(action));
            continue;
          }
          if (action.type === 'speech' && !action.audioAssetId)
            throw new Error('本页部分讲解尚未生成语音，请先完成语音生成');
          const audio = await request(action.type === 'speech' ? 'tts:preview' : 'media:preview', {
            ...scope(),
            actionId: action.id,
          });
          if (!isCurrent()) return;
          bytesTotal += audio.bytes.byteLength;
          if (bytesTotal > 256 * 1024 * 1024) throw new Error('本页媒体超过本地预览大小限制');
          const url = URL.createObjectURL(new Blob([audio.bytes], { type: audio.mimeType }));
          timelineUrls.push(url);
          nodes.push(
            action.type === 'video_clip'
              ? { id: action.id, type: 'video_clip', videoUrl: url }
              : {
                  id: action.id,
                  type: action.type,
                  text: action.text,
                  audioAssetId: action.audioAssetId,
                  audioUrl: url,
                  pauseAfterMs: action.pauseAfterMs,
                },
          );
        }
        if (inSlideshow) {
          startingSlideshow = true;
          try {
            teachingSlideshow = createTeachingSlideshow({
              doc: nativeDocument!,
              page: view!.pageNumber,
              timelineStepCount: stepCount || undefined,
              generatedMotions: true,
              schedule,
              onReady: () => {
                if (isCurrent())
                  void startTeaching().catch((error) => {
                    if (isCurrent()) {
                      stopAudio();
                      notice(String(error), true);
                    }
                  });
              },
              onExit: () => {
                if (isCurrent()) {
                  stopAudio();
                  notice('教学放映已停止');
                }
              },
              onError: (error) => {
                if (isCurrent()) {
                  stopAudio();
                  notice(error.message, true);
                }
              },
            });
          } finally {
            startingSlideshow = false;
          }
          const controls = teachingSlideshow.controls;
          const navigate = (number: number) => {
            if (number < 1 || number > view!.project.pages.length) return;
            if (lesson) {
              const current = lesson;
              if (current.navigating || current.targetPage !== undefined) return;
              current.targetPage = number;
              const finish = current.finishPage;
              stopAudio(current);
              finish?.();
            } else {
              stopAudio();
              void playLesson(number);
            }
          };
          for (const [label, command, action] of [
            ['上一页', 'previous', () => navigate(view!.pageNumber - 1)],
            [
              '暂停播放',
              'toggle',
              () => {
                if (previewMode === 'paused') {
                  if (!stepCount) teachingSlideshow?.resume();
                  timeline?.resume();
                } else {
                  if (!stepCount) teachingSlideshow?.pause();
                  timeline?.pause();
                }
              },
            ],
            ['重播本页', 'replay', () => navigate(view!.pageNumber)],
            ['下一页', 'next', () => navigate(view!.pageNumber + 1)],
            ['教师继续', 'teacher', () => timeline?.continueTeacher()],
          ] as const) {
            const control = button(label, action);
            control.dataset.previewCommand = command;
            control.title = label;
            control.disabled =
              command === 'previous'
                ? view!.pageNumber === 1
                : command === 'next'
                  ? view!.pageNumber === view!.project.pages.length
                  : false;
            controls.append(control);
          }
          const seekBar = element('label');
          seekBar.dataset.playbackSeek = '';
          const range = element('input');
          range.type = 'range';
          range.min = '0';
          range.max = '0';
          range.value = '0';
          range.step = '10';
          range.disabled = true;
          range.setAttribute('aria-label', '当前讲解或媒体进度');
          const time = element('output', '0:00 / 0:00');
          time.setAttribute('role', 'timer');
          time.setAttribute('aria-label', '当前媒体时间');
          const seek = (ms: number) => timeline?.seek(Math.max(Number(range.min), Math.min(Number(range.max), ms)));
          range.addEventListener('input', () => seek(Number(range.value)));
          const forward = button('快进 5 秒', () => seek(Number(range.value) + 5000), true);
          forward.dataset.previewCommand = 'forward';
          seekBar.append(range, time, forward);
          controls.append(seekBar);
          controls.append(
            button('退出教学放映', () => {
              stopAudio();
              notice('教学放映已停止');
            }),
          );
          const status = element('span');
          status.dataset.showStatus = '';
          status.setAttribute('role', 'status');
          controls.append(status);
          notice(
            nativeNodes.length
              ? '正在按教学时间轴自动放映 PPT 动画与讲解'
              : '正在放映本页；如有未绑定的原生动画，请使用「下一步 PPT 动画」，随后自动播放教学时间轴',
          );
          renderPreviewControls();
        } else await startTeaching();
        async function startTeaching() {
          if (
            content.some((action) =>
              action.visualCues?.some((cue: AttentionCue) => ['highlight', 'spotlight'].includes(cue.effect?.type)),
            )
          )
            await previewAttention(isCurrent, teachingSlideshow?.canvas);
          if (!isCurrent()) return;
          if (content.some((a) => a.type === 'video_clip')) {
            previewVideo = panel!.ownerDocument.createElement('video');
            previewVideo.controls = true;
            previewVideo.preload = 'metadata';
            previewVideo.hidden = true;
            previewVideo.style.cssText = 'width:100%;max-height:260px;background:#000;';
            previewVideo.setAttribute('aria-label', '本页教学视频');
            if (teachingSlideshow) {
              previewVideo.controls = false;
              previewVideo.style.cssText =
                'position:absolute;inset:0 0 var(--teaching-controls-height,112px);width:100%;height:calc(100% - var(--teaching-controls-height,112px));object-fit:contain;background:#000;z-index:40;';
              teachingSlideshow.surface.append(previewVideo);
            } else panel!.querySelector('[data-preview-command]')!.parentElement!.after(previewVideo);
          }
          const updateMediaProgress = (progress: { timeMs: number; durationMs: number } | null) => {
            const bar = teachingSlideshow?.controls.querySelector('[data-playback-seek]');
            if (!bar) return;
            const range = bar.querySelector('input')!;
            const forward = bar.querySelector('button')!;
            range.disabled = forward.disabled = !progress || progress.durationMs <= 0;
            range.min = String(motionSteps.length && progress ? Math.min(progress.timeMs, progress.durationMs) : 0);
            range.max = String(progress?.durationMs ?? 0);
            range.value = String(progress?.timeMs ?? 0);
            const clock = (ms: number) =>
              `${Math.floor(ms / 60000)}:${String(Math.floor(ms / 1000) % 60).padStart(2, '0')}`;
            bar.querySelector('output')!.textContent =
              `${clock(progress?.timeMs ?? 0)} / ${clock(progress?.durationMs ?? 0)}`;
            range.title = motionSteps.length
              ? '动画同步播放支持向前快进；回看请重播本页'
              : '拖动调整当前讲解或媒体进度';
          };
          timeline = shared.createTeachingTimelinePreview({
            video: previewVideo ?? undefined,
            pageId: view!.pageId,
            nodes,
            speechMotions: motionSteps.length ? motionSteps : undefined,
            native: stepCount
              ? {
                  sourceSha256: view!.project.source!.sha256,
                  stepCount,
                  next: (index: number) => teachingSlideshow!.next(index),
                  pause: () => teachingSlideshow?.pause(),
                  resume: () => teachingSlideshow?.resume(),
                }
              : undefined,
            onMode(mode) {
              if (!isCurrent()) return;
              previewMode = mode;
              renderPreviewControls();
              if (mode === 'waiting_interaction') notice('等待教师继续，点击「教师继续」后播放下一段');
              else if (mode === 'paused') notice('本页教学已暂停');
              else if (mode === 'playing')
                notice(teachingSlideshow ? '正在幻灯片放映' : '正在画布预览；PPT 动画请使用原生动画预览');
            },
            onMediaProgress: updateMediaProgress,
            onProgress(progress) {
              if (!isCurrent()) return;
              updateMediaProgress(progress);
              const action = progress ? content.find((a) => a.id === progress.actionId) : undefined;
              attention?.show(
                progress
                  ? (action?.visualCues ?? []).filter(
                      (cue: AttentionCue) =>
                        ['highlight', 'spotlight'].includes(cue.effect?.type) &&
                        cue.atMs <= progress.timeMs &&
                        progress.timeMs < cue.endMs,
                    )
                  : [],
              );
            },
            onInteraction(question) {
              const parent = teachingSlideshow?.controls ?? panel!;
              parent.querySelector('[data-teaching-question]')?.remove();
              if (!question || !isCurrent()) return;
              const card = element('section');
              card.dataset.teachingQuestion = question.checkpointId;
              card.setAttribute('aria-label', '课堂提问');
              card.style.cssText =
                'max-width:600px;max-height:60vh;overflow:auto;background:white;color:#222;padding:16px;';
              if (teachingSlideshow) card.style.cssText += 'position:absolute;bottom:64px;right:12px;';
              card.append(element('strong', question.question));
              if (question.response) {
                const response = question.response;
                const choices = element('fieldset'),
                  result = element('p');
                result.setAttribute('role', 'status');
                choices.append(element('legend', response.options.length ? '选择答案' : '课堂回应'));
                if (response.options.length) {
                  const selected = new Set<number>();
                  for (const [index, option] of response.options.entries()) {
                    const label = element('label'),
                      field = element('input');
                    label.style.display = 'block';
                    field.type = response.kind === 'multiple' ? 'checkbox' : 'radio';
                    field.name = `answer-${question.checkpointId}`;
                    field.addEventListener('change', () => {
                      if (response.kind !== 'multiple') selected.clear();
                      if (field.checked) selected.add(index);
                      else selected.delete(index);
                      result.textContent = '';
                    });
                    label.append(field, element('span', option));
                    choices.append(label);
                  }
                  choices.append(
                    button('检查答案', () => {
                      result.textContent = !selected.size
                        ? '请先选择答案'
                        : shared.evaluateQuestionAnswer(response, [...selected])
                          ? '回答正确'
                          : '再想一想，可以重新选择';
                    }),
                    result,
                  );
                } else {
                  const field = element('textarea');
                  field.maxLength = 4000;
                  field.setAttribute('aria-label', '课堂回应');
                  choices.append(field, element('p', '讨论或完成活动后，由教师结合参考答案评价。'));
                }
                card.append(choices);
              }
              const answer = element('details');
              answer.append(element('summary', '参考答案'), element('p', question.referenceAnswer.summary));
              const points = element('ul');
              for (const point of question.referenceAnswer.keyPoints) points.append(element('li', point));
              answer.append(points);
              card.append(
                answer,
                button('继续教学', () => timeline?.continueTeacher()),
              );
              parent.append(card);
            },
            onNode(id) {
              if (!isCurrent()) return;
              panel?.querySelectorAll<HTMLElement>('[data-action-id]').forEach((node) => {
                if (node.dataset.actionId === id) node.dataset.playing = '';
                else node.removeAttribute('data-playing');
              });
            },
            onError(error) {
              if (isCurrent()) {
                if (nativeNodes.length) {
        assertNativeTeachingAnimations(view!.pageNumber);
                  stopAudio();
                  notice(`播放失败：${error.message}。请重新开始教学放映。`, true);
                } else notice(`播放失败：${error.message}。可继续重试或停止预览。`, true);
              }
            },
            onComplete() {
              if (isCurrent()) {
                const finish = preserveLesson?.finishPage;
                stopAudio(preserveLesson);
                notice('本页教学播放完成');
                finish?.();
              }
            },
          });
          timeline.start();
        }
      } catch (error) {
        if (!isCurrent()) return;
        stopAudio();
        throw error;
      } finally {
        if (!isCurrent() || !teachingSlideshow) schedule?.restore();
      }
    }, false);

  const playLesson = async (startPage = 1) => {
    if (!view || !context || busy || running || lesson || nativeSlideshow || draftError) return;
    if (changed()) {
      await run(saveCurrent);
      if (changed() || busy || running || draftError || !view || !context) return;
    }
    const api = getEditorApi() as any;
    const slides = api?.WordControl?.m_oLogicDocument?.Slides;
    const source = view.project.source?.sha256;
    const revision = view.project.teachingRevision;
    const pageIds = view.project.pages.map((page) => page.pageId);
    if (
      !source ||
      !Array.isArray(slides) ||
      slides.length !== pageIds.length ||
      new Set(pageIds).size !== pageIds.length ||
      typeof api.goToPage !== 'function'
    ) {
      notice('请先保存并重新打开完整课件，再放映整课教学', true);
      return;
    }
    const savedSlides = [...slides],
      initialPage = context.page;
    const cover = element('section');
    cover.dataset.teachingLesson = '';
    cover.setAttribute('aria-label', '整课教学放映');
    cover.style.cssText =
      'position:fixed;inset:0;z-index:9999;background:#111;color:#fff;display:flex;align-items:center;justify-content:center;gap:16px;';
    const status = element('span');
    status.setAttribute('role', 'status');
    const stop = () => {
      stopAudio();
      notice('整课教学放映已停止');
    };
    const stopButton = button('停止整课放映', stop);
    cover.append(status, stopButton);
    cover.addEventListener('keydown', (event) => {
      if (event.key === 'Escape') stop();
      if (event.key === 'Tab') {
        event.preventDefault();
        stopButton.focus();
      }
      event.stopPropagation();
    });
    nativeDocument!.body.append(cover);
    stopButton.focus();
    const current: LessonRun = { cover, page: 0, total: slides.length, navigating: false };
    lesson = current;
    let completed = false;
    const assertCurrent = () => {
      const actualSlides = api.WordControl?.m_oLogicDocument?.Slides;
      if (
        lesson !== current ||
        disposed ||
        getEditorApi() !== api ||
        actualSlides?.length !== savedSlides.length ||
        savedSlides.some((slide, index) => actualSlides[index] !== slide)
      )
        throw new Error('整课放映来源已变化');
    };
    try {
      stopAudio(current);
      for (let number = startPage; number <= current.total; number = current.targetPage ?? number + 1) {
        current.targetPage = undefined;
        assertCurrent();
        current.page = number;
        status.textContent = `正在准备第 ${number}/${current.total} 页教学…`;
        current.navigating = true;
        try {
          api.goToPage(number - 1);
          const target = deckwiseEditorBridge.getContext(binding);
          if (target.page !== number) throw new Error('无法切换到整课放映的目标页');
          await load(target, current);
          assertCurrent();
          const actual = deckwiseEditorBridge.getContext(binding);
          if (!view) throw new Error(statusText || '无法读取目标页教学');
          if (
            actual.page !== number ||
            actual.editGeneration !== target.editGeneration ||
            view.pageNumber !== number ||
            view.project.source?.sha256 !== source ||
            view.project.teachingRevision !== revision ||
            !pageIds.includes(view.pageId) ||
            changed() ||
            draftError
          )
            throw new Error('课件或教学版本已变化，或存在未保存的教学草稿；整课放映已停止');
        } finally {
          current.navigating = false;
        }
        // Pages without teaching steps have nothing to run; the shared timeline requires at least one node.
        if (!actions.length) continue;
        let finished = false;
        const done = new Promise<Error | undefined>((resolve) => {
          current.finishPage = (error) => {
            finished = true;
            resolve(error);
          };
        });
        await playTimeline(true, current);
        assertCurrent();
        if (!teachingSlideshow && !finished) throw new Error(statusText || `第 ${number} 页教学未能启动`);
        const failure = await done;
        if (failure) throw failure;
        current.finishPage = undefined;
      }
      assertCurrent();
      completed = true;
      lesson = null;
      cover.remove();
      // Restore only after a successful owned run, never after manual navigation or cancellation.
      api.goToPage(initialPage! - 1);
      await load(deckwiseEditorBridge.getContext(binding));
      if (!view) throw new Error(statusText || '原编辑页读取失败');
      notice('整课教学放映完成');
    } catch (error) {
      if (lesson === current) {
        stopAudio();
        notice(`第 ${current.page}/${current.total} 页放映停止：${String(error)}`, true);
      } else if (completed && !disposed) notice(`整课已放映完成，但返回编辑页失败：${String(error)}`, true);
    } finally {
      renderControls();
    }
  };

  function attentionControls(action: Action, index: number) {
    const section = element('section');
    section.className = 'teaching-attention';
    section.setAttribute('aria-label', t('teachingSpotlight'));
    section.append(element('strong', t('teachingFocusTitle')));
    const cues: AttentionCue[] = Array.isArray(action.visualCues) ? action.visualCues : [];
    const allowed = editable() && !!action.audioAssetId;
    const objects = (view?.teachingObjects ?? []).filter(
      (object) => object.status === 'mapped' && object.effectTargetEligible,
    );
    const objectViews = teachingObjectChoices(view!.pageNumber, objects);
    const loadThumbnails: Array<() => void> = [];
    const form = element('div');
    form.dataset.cueForm = '';
    form.hidden = true;
    let editingId: string | null = null;
    const kind = element('select');
    kind.setAttribute('aria-label', '教学提示类型');
    kind.append(new Option(t('teachingSpotlight'), 'spotlight'), new Option(t('teachingHighlight'), 'highlight'));
    const targets = element('fieldset');
    targets.append(element('legend', '目标对象（可多选）'));
    targets.append(element('p', t('teachingObjectPickerHint')));
    const choices = objectViews.map((object) => {
      const check = element('input');
      check.type = 'checkbox';
      check.value = object.objectId;
      loadThumbnails.push(
        appendTeachingObjectChoice(section.ownerDocument, targets, check, object, (message) => notice(message, true)),
      );
      return check;
    });
    const field = (label: string, input: HTMLElement) => {
      const row = element('label', label);
      row.append(input);
      return row;
    };
    const start = element('input'),
      end = element('input');
    for (const [input, name, min] of [
      [start, t('teachingStartSeconds'), '0'],
      [end, t('teachingEndSeconds'), '0.001'],
    ] as const) {
      input.type = 'number';
      input.min = min;
      input.max = String((action.durationMs || 3_600_000) / 1000);
      input.step = '0.1';
      input.setAttribute('aria-label', name);
    }
    const style = element('select');
    style.setAttribute('aria-label', '高亮样式');
    style.append(new Option('描边', 'outline'), new Option('填充', 'fill'), new Option('发光', 'glow'));
    const color = element('input');
    color.type = 'color';
    color.setAttribute('aria-label', '高亮颜色');
    const opacity = element('input');
    opacity.type = 'number';
    opacity.min = '0';
    opacity.max = '100';
    opacity.step = '1';
    opacity.setAttribute('aria-label', '效果不透明度百分比');
    const highlight = element('div');
    highlight.append(field('高亮样式', style), field('高亮颜色', color));
    kind.addEventListener('change', () => {
      highlight.hidden = kind.value !== 'highlight';
    });
    const submit = button(
      t('teachingCueAdd'),
      () => {
        const ids = choices.filter((choice) => choice.checked).map((choice) => choice.value);
        const atMs = Math.round(start.valueAsNumber * 1000),
          endMs = Math.round(end.valueAsNumber * 1000),
          alpha = opacity.valueAsNumber;
        if (
          !allowed ||
          actions[index]?.audioAssetId !== action.audioAssetId ||
          !ids.length ||
          !Number.isInteger(atMs) ||
          !Number.isInteger(endMs) ||
          atMs < 0 ||
          endMs <= atMs ||
          endMs > (action.durationMs || 3_600_000) ||
          !Number.isInteger(alpha) ||
          alpha < 0 ||
          alpha > 100 ||
          (!editingId && cues.length >= 200)
        ) {
          notice('请选择目标对象，并填写有效的时间和不透明度', true);
          return;
        }
        const cue: AttentionCue = {
          id: editingId ?? crypto.randomUUID(),
          atMs,
          endMs,
          effect:
            kind.value === 'highlight'
              ? { type: 'highlight', elementIds: ids, style: style.value, color: color.value, opacity: alpha / 100 }
              : {
                  type: 'spotlight',
                  elementId: ids[0],
                  ...(ids.length > 1 ? { elementIds: ids.slice(1) } : {}),
                  dimOpacity: alpha / 100,
                },
        };
        edit(
          actions.map((entry, i) =>
            i === index
              ? {
                  ...entry,
                  visualCues: editingId
                    ? cues.map((existing) => (existing.id === editingId ? cue : existing))
                    : [...cues, cue],
                }
              : entry,
          ),
        );
        notice('效果已加入本页教学草稿；可直接预览，保存教学后保留');
      },
      !allowed,
    );
    const openForm = (cue?: AttentionCue, type = 'spotlight') => {
      editingId = cue?.id ?? null;
      kind.value = cue?.effect.type ?? type;
      const ids = cue
        ? cue.effect.type === 'highlight'
          ? (cue.effect.elementIds ?? [])
          : [cue.effect.elementId, ...(cue.effect.elementIds ?? [])]
        : [];
      choices.forEach((choice) => {
        choice.checked = ids.includes(choice.value);
      });
      start.value = String((cue?.atMs ?? 0) / 1000);
      end.value = String((cue?.endMs ?? Math.min(action.durationMs || 1000, 1000)) / 1000);
      style.value = cue?.effect.style ?? 'outline';
      color.value = cue?.effect.color ?? '#f59e0b';
      opacity.value = String(
        Math.round(
          (cue
            ? cue.effect.type === 'highlight'
              ? (cue.effect.opacity ?? 0.8)
              : (cue.effect.dimOpacity ?? 0.6)
            : 0.8) * 100,
        ),
      );
      highlight.hidden = kind.value !== 'highlight';
      submit.textContent = t(cue ? 'teachingCueSave' : 'teachingCueAdd');
      form.hidden = false;
      if (!cue) {
        const logic = (getEditorApi() as any)?.WordControl?.m_oLogicDocument;
        const selection = logic?.GetCurrentController?.()?.selectedObjects ?? [];
        if (selection.length) {
          try {
            const selected = selectedAttentionTargets(
              logic?.Slides?.[(view?.pageNumber ?? 1) - 1]?.cSld?.spTree ?? [],
              selection,
              objects,
            );
            choices.forEach((choice) => (choice.checked = selected.includes(choice.value)));
          } catch (error) {
            notice(String(error), true);
          }
        }
      }
      loadThumbnails.forEach((load) => load());
      kind.focus();
    };
    for (const cue of cues) {
      if (!['highlight', 'spotlight'].includes(cue.effect?.type)) continue;
      const ids =
        cue.effect.type === 'highlight'
          ? (cue.effect.elementIds ?? [])
          : [cue.effect.elementId, ...(cue.effect.elementIds ?? [])];
      const row = element('div');
      row.dataset.cueId = cue.id;
      row.append(
        element(
          'p',
          `${cue.effect.type === 'highlight' ? t('teachingHighlightEffect') : t('teachingSpotlightEffect')} · ${cue.atMs / 1000}–${cue.endMs / 1000} s`,
        ),
        element(
          'p',
          ids.map((id) => objectViews.find((object) => object.objectId === id)?.name || '对象映射待确认').join('、'),
        ),
        button(
          '预览此提示',
          () =>
            void run(async () => {
              if (nativeSlideshow) throw new Error('请先退出幻灯片放映，再预览教学效果');
              stopAudio();
              const preview = await previewAttention();
              preview.show([cue]);
              notice('画布正在预览此效果；点击停止预览可关闭');
            }),
          busy,
        ),
        button('编辑此提示', () => openForm(cue), !allowed),
        button(
          '删除此提示',
          () =>
            edit(
              actions.map((entry, i) =>
                i === index ? { ...entry, visualCues: cues.filter((existing) => existing.id !== cue.id) } : entry,
              ),
            ),
          !allowed,
        ),
      );
      row.title = ids.map((id) => objects.find((object) => object.objectId === id)?.name ?? id).join(', ');
      for (const [i, control] of [...row.querySelectorAll('button')].entries())
        decorateTeachingButton(
          control,
          (['play', 'edit', 'delete'] as const)[i],
          i !== 1,
          i === 1 ? t('agentQueueEdit') : undefined,
        );
      section.append(row);
    }
    const unavailable = page()?.locked
      ? t('teachingEffectsLocked')
      : !action.audioAssetId
        ? t('teachingEffectsNeedVoice')
        : !objects.length
          ? t('teachingEffectsNeedObjects')
          : '';
    if (unavailable) section.append(element('p', unavailable));
    const formActions = element('div');
    formActions.className = 'teaching-cue-actions';
    submit.className = 'teaching-primary';
    const cancel = button(t('teachingCueCancel'), () => {
      form.hidden = true;
    });
    cancel.className = 'teaching-secondary';
    formActions.append(submit, cancel);
    form.append(
      kind,
      button('使用画布选中对象', () => {
        try {
          const current = deckwiseEditorBridge.getContext(binding);
          if (typeof current.page !== 'number' || current.page !== view?.pageNumber)
            throw new Error('页面已切换，请等待教学数据更新');
          const api = getEditorApi() as any,
            logic = api?.WordControl?.m_oLogicDocument;
          const ids = selectedAttentionTargets(
            logic?.Slides?.[current.page - 1]?.cSld?.spTree ?? [],
            logic?.GetCurrentController?.()?.selectedObjects ?? [],
            objects,
          );
          choices.forEach((choice) => {
            choice.checked = ids.includes(choice.value);
          });
          notice(`已选择 ${ids.length} 个画布对象`);
        } catch (error) {
          notice(String(error), true);
        }
      }),
      targets,
      field(t('teachingStartSeconds'), start),
      field(t('teachingEndSeconds'), end),
      highlight,
      field('不透明度（%，聚光为背景遮罩）', opacity),
      formActions,
    );
    const effectActions = element('div');
    effectActions.className = 'teaching-effect-actions';
    const addSpotlight = button(t('teachingSpotlight'), () => openForm(), !allowed || !objects.length);
    addSpotlight.className = 'teaching-primary';
    effectActions.append(
      addSpotlight,
      button(t('teachingHighlight'), () => openForm(undefined, 'highlight'), !allowed || !objects.length),
      button(t('teachingEditAnimation'), () => editNativeAnimation(action), !editable() || nativeSlideshow),
    );
    for (const [i, control] of [...effectActions.querySelectorAll('button')].entries())
      decorateTeachingButton(
        control,
        i === 2 ? 'edit' : 'plus',
        false,
        t((['teachingSpotlightEffect', 'teachingHighlightEffect', 'teachingNativeAnimationShort'] as const)[i]),
      );
    effectActions.lastElementChild!.setAttribute('title', t('teachingNativeMotionHint'));
    section.insertBefore(effectActions, section.children[1] ?? null);
    section.append(
      createTeachingAnimationForm({
        doc: section.ownerDocument,
        action: { ...action, text: action.text, visualCues: cues as AnimationCue[] },
        editable: editable(),
        objects: objectViews,
        selection: () => speechSelections.get(action.id),
        native: () =>
          listTeachingNativeAnimations(view!.pageNumber, view?.teachingObjects ?? []).map((choice) => {
            const owner = actions.find(
              (entry) =>
                entry.id !== action.id &&
                entry.visualCues?.some((cue: AttentionCue) => cue.nativeAnimation === choice.objectId),
            );
            return {
              ...choice,
              owner: owner ? `${actions.indexOf(owner) + 1} · ${owner.text?.slice(0, 30) ?? ''}` : undefined,
            };
          }),
        selectedObjects: () => {
          const api = getEditorApi();
          const logic = Reflect.get(Reflect.get(api ?? {}, 'WordControl') ?? {}, 'm_oLogicDocument');
          return selectedAttentionTargets(
            logic?.Slides?.[view!.pageNumber - 1]?.cSld?.spTree ?? [],
            logic?.GetCurrentController?.()?.selectedObjects ?? [],
            objects,
          );
        },
        preview: () => previewMedia(action, true),
        audioDuration: () =>
          previewAudio && Number.isFinite(previewAudio.duration) ? Math.round(previewAudio.duration * 1000) : undefined,
        position: () =>
          previewAudio && Number.isFinite(previewAudio.currentTime)
            ? Math.round(previewAudio.currentTime * 1000)
            : undefined,
        error: (message) => notice(message, true),
        save: (next, moveNative) => {
          stopAudio();
          edit(
            actions.map((entry, i) =>
              i === index
                ? { ...entry, visualCues: next }
                : moveNative.length && Array.isArray(entry.visualCues)
                  ? {
                      ...entry,
                      visualCues: entry.visualCues.filter(
                        (cue: AttentionCue) => !moveNative.includes(cue.nativeAnimation ?? ''),
                      ),
                    }
                  : entry,
            ),
          );
        },
      }),
    );
    if (allowed && objects.length) section.append(form);
    section.append(
      element('p', '选择画布对象或勾选目标，设置时间后添加；试听讲解时效果随音频播放。PPT 动画使用原生动画选项卡。'),
    );
    section.lastElementChild?.classList.add('teaching-attention-hint');
    return section;
  }
  async function previewMedia(action: Action, autoplay: boolean) {
    if (nativeSlideshow) throw new Error('请先退出原生幻灯片放映，再试听语音');
    const requestId = previewRequest;
    if (changed()) await saveCurrent();
    if (requestId !== previewRequest || nativeSlideshow || disposed) return;
    stopAudio();
    const token = epoch,
      runId = timelineRun;
    const currentCard = panel?.querySelector<HTMLElement>('[data-teaching-inspector]');
    if (currentCard?.dataset.actionId !== action.id) return;
    previewAudio = element(action.type === 'video_clip' ? 'video' : 'audio');
    previewAudio.controls = true;
    previewAudio.preload = 'metadata';
    previewAudio.setAttribute('aria-label', teachingStepTitle(action));
    previewAudio.addEventListener('loadedmetadata', () => {
      const duration = currentCard.querySelector('[data-media-preview] > p');
      if (duration && Number.isFinite(player.duration))
        duration.textContent = t('teachingMediaDuration', { seconds: Number(player.duration.toFixed(1)) });
    });
    if (typeof action.gain === 'number') previewAudio.volume = Math.max(0, Math.min(1, action.gain));
    (currentCard.querySelector('[data-media-preview]') ?? currentCard).append(previewAudio);
    const player = previewAudio;
    const audio = await request(action.type === 'speech' ? 'tts:preview' : 'media:preview', {
      ...scope(),
      actionId: action.id,
    });
    if (token !== epoch || runId !== timelineRun || disposed || nativeSlideshow || previewAudio !== player) return;
    previewUrl = URL.createObjectURL(new Blob([audio.bytes], { type: audio.mimeType }));
    const cues = (action.visualCues ?? []).filter((cue: AttentionCue) =>
      ['highlight', 'spotlight'].includes(cue.effect?.type),
    );
    if (cues.length) {
      try {
        const preview = await previewAttention();
        scheduler = new preview.shared.SpeechCueScheduler(
          [],
          () => (previewAudio?.currentTime ?? 0) * 1000,
          () => {},
          undefined,
          (ms) => preview.show(cues.filter((cue: AttentionCue) => cue.atMs <= ms && ms < cue.endMs)),
        );
        const clock = scheduler;
        previewAudio.addEventListener('playing', () => clock.start());
        previewAudio.addEventListener('pause', () => clock.pause());
        previewAudio.addEventListener('waiting', () => clock.pause());
        previewAudio.addEventListener('seeking', () => clock.pause());
        previewAudio.addEventListener('seeked', () => {
          clock.seek();
          if (!previewAudio?.paused) clock.start();
        });
        previewAudio.addEventListener('ended', () => {
          clock.stop();
          preview.show([]);
        });
        clock.seek();
      } catch (error) {
        stopAudio();
        throw error;
      }
    }
    if (token !== epoch || runId !== timelineRun || disposed || nativeSlideshow || !previewAudio) {
      stopAudio();
      return;
    }
    previewAudio.src = previewUrl;
    previewAudio.addEventListener('error', () => {
      if (previewAudio !== player) return;
      stopAudio();
      notice(t('teachingMediaFailed'), true);
    });
    if (autoplay) await previewAudio.play();
  }
  const speechSelections = new Map<string, { start: number; end: number }>();
  function speechCard(action: Action, index: number) {
    const card = element('article');
    card.dataset.actionId = action.id;
    card.append(element('strong', `${index + 1} · 讲解与语音`));
    const text = element('textarea');
    text.rows = 1;
    text.className = 'teaching-speech-text';
    text.value = action.text;
    text.disabled = !editable();
    text.setAttribute('aria-label', `第 ${index + 1} 段讲解`);
    text.addEventListener('select', () => {
      if (text.selectionEnd > text.selectionStart)
        speechSelections.set(action.id, { start: text.selectionStart, end: text.selectionEnd });
    });
    text.addEventListener('input', () => {
      const next: Action = { ...actions[index], text: text.value };
      for (const key of [
        'audioId',
        'audioUrl',
        'audioAssetId',
        'duration',
        'durationMs',
        'visualCues',
        'speechCuePlans',
        'pptxCue',
        'attention',
      ])
        delete next[key];
      if (Array.isArray(action.visualCues)) {
        const retained = action.visualCues
          .filter((cue: AttentionCue) => cue.trigger)
          .map((cue: AttentionCue) => ({ ...cue, trigger: relocateAnimationTrigger(cue.trigger!, text.value) }));
        if (retained.length) next.visualCues = retained;
      }
      speechSelections.delete(action.id);
      stopAudio();
      edit(
        actions.map((entry, i) => (i === index ? next : entry)),
        false,
      );
      const summary = [...(panel?.querySelectorAll<HTMLElement>('.teaching-flow article') ?? [])]
        .find((node) => node.dataset.actionId === action.id)
        ?.querySelector('.teaching-step-summary');
      if (summary) summary.textContent = text.value || t('teachingEmptySpeech');
      panel!.querySelector('[data-teaching-inspector] [data-audio-status]')!.textContent =
        '讲解已修改，保存后可重新生成语音；原语音已失效，选词动画需重新定位';
    });
    card.append(text);
    const status = element(
      'p',
      action.audioAssetId
        ? `语音已就绪${action.durationMs ? ` · ${(action.durationMs / 1000).toFixed(1)} 秒` : ''}`
        : '尚未生成语音',
    );
    status.dataset.audioStatus = '';
    const voiceStatus = element('div');
    voiceStatus.className = 'teaching-voice-status';
    voiceStatus.append(status);
    if (action.audioAssetId)
      voiceStatus.append(
        button(
          '试听语音',
          () =>
            void run(async () => {
              await previewMedia(action, true);
            }),
          busy || !!draftError,
        ),
      );

    const pauseLabel = element('label', t('teachingAfter'));
    pauseLabel.className = 'teaching-after';
    const pauseMode = element('select');
    pauseMode.setAttribute('aria-label', t('teachingAfter'));
    pauseMode.append(
      new Option(t('teachingNone'), 'none'),
      new Option(t('teachingPause'), 'pause'),
      new Option(t('teachingTimed'), 'timed'),
    );
    pauseMode.value = speechPauseMode(action.pauseAfterMs);
    pauseMode.disabled = !editable();
    pauseMode.addEventListener('change', () => {
      const value =
        pauseMode.value === 'none'
          ? 0
          : pauseMode.value === 'pause'
            ? speechPauseMode(action.pauseAfterMs) === 'pause'
              ? action.pauseAfterMs
              : 400
            : Math.max(3000, action.pauseAfterMs || 0);
      edit(actions.map((entry, i) => (i === index ? { ...entry, pauseAfterMs: value } : entry)));
    });
    pauseLabel.append(pauseMode);
    if (pauseMode.value === 'timed') {
      const pause = element('input');
      pause.type = 'number';
      pause.min = '3';
      pause.max = '30';
      pause.step = '0.5';
      pause.value = String(action.pauseAfterMs / 1000);
      pause.disabled = !editable();
      pause.setAttribute('aria-label', t('teachingTimedSeconds'));
      pause.addEventListener('change', () => {
        const ms = Math.round(pause.valueAsNumber * 1000);
        if (!Number.isInteger(ms) || ms < 3000 || ms > 30000) {
          pause.value = String(action.pauseAfterMs / 1000);
          notice(t('teachingPauseRange'), true);
          return;
        }
        edit(actions.map((entry, i) => (i === index ? { ...entry, pauseAfterMs: ms } : entry)));
      });
      pauseLabel.append(pause, element('span', t('teachingSeconds')));
    }
    const pauseHint = element('p', t('teachingAfterHint'));
    pauseHint.className = 'teaching-after-hint';
    pauseHint.hidden = pauseMode.value === 'none';
    pauseLabel.title = t('teachingAfterHint');
    card.append(voiceStatus, pauseLabel, pauseHint);
    const voice = element('details');
    voice.append(element('summary', '语音设置与生成'));
    const modelSelect = element('select');
    modelSelect.setAttribute('aria-label', '语音模型');
    modelSelect.append(new Option('选择语音模型', ''));
    models.forEach((model) => modelSelect.append(new Option(model.model, model.model)));
    const voiceSelect = element('select');
    voiceSelect.setAttribute('aria-label', '音色');
    const populateVoices = () => {
      voiceSelect.replaceChildren(new Option('服务默认音色', ''));
      models
        .find((model) => model.model === modelSelect.value)
        ?.voices.forEach((value) => voiceSelect.append(new Option(value, value)));
    };
    modelSelect.addEventListener('change', populateVoices);
    populateVoices();
    voice.append(
      modelSelect,
      voiceSelect,
      button(
        '获取生成报价',
        () =>
          void run(async () => {
            if (changed()) throw new Error('请先保存讲解，再生成语音');
            const selected = models.find((model) => model.model === modelSelect.value);
            if (!selected || (!voiceSelect.value && !selected.providerDefaultVoice))
              throw new Error('请选择可用模型和音色');
            const prepared = await request('tts:prepare', {
              ...scope(),
              actionId: action.id,
              model: selected.model,
              voice: voiceSelect.value || null,
            });
            const quote = await request('tts:quote', { ...scope(), intentId: prepared.intentId });
            const target = scope();
            // This explicit paid action never runs during automatic reads or retries.
            const confirm = button(`确认生成（最多 ${quote.maxPoints} 点）`, () => {
              confirm.disabled = true;
              void run(async () => {
                if (
                  changed() ||
                  draftError ||
                  !view ||
                  target.viewId !== view.viewId ||
                  target.expectedTeachingRevision !== view.project.teachingRevision
                )
                  throw new Error('讲解或页面已变化，请保存后重新获取报价');
                const task = await request('tts:submit', {
                  ...target,
                  intentId: prepared.intentId,
                  quoteId: quote.quoteId,
                });
                tasks = [...tasks.filter((entry) => entry.intentId !== task.intentId), task];
                render();
              });
            });
            const currentInspector = panel?.querySelector<HTMLElement>('[data-teaching-inspector]');
            if (!currentInspector || currentInspector.dataset.actionId !== action.id)
              throw new Error('Teaching selection changed');
            currentInspector.append(confirm);
            notice('报价已就绪，点击确认后才会生成收费语音');
          }),
        !editable() || changed(),
      ),
    );
    card.append(attentionControls(action, index), voice);
    for (const task of tasks.filter((entry) => entry.actionId === action.id)) {
      const row = element('div');
      row.append(
        element(
          'p',
          task.task?.state === 'ready'
            ? t('teachingVoiceGenerated')
            : task.task?.state === 'failed'
              ? t('teachingVoiceFailed')
              : task.phase === 'submitted'
                ? t('teachingVoiceGenerating')
                : t('teachingVoicePrepared'),
        ),
      );
      const technical = element('small', `${task.model} · ${task.voice ?? '默认音色'}`);
      row.append(technical);
      row.append(
        button(
          '检查生成结果',
          () =>
            void run(async () => {
              const next = await request('tts:status', { ...scope(), intentId: task.intentId });
              tasks = tasks.map((entry) => (entry.intentId === next.intentId ? next : entry));
              render();
            }),
          busy,
        ),
      );
      if (task.task?.state === 'ready' && task.current)
        row.append(
          button(
            '使用生成的语音',
            () =>
              void run(async () => {
                await request('tts:archive', { ...scope(), intentId: task.intentId });
                await applyView(await request('tts:adopt', { ...scope(), intentId: task.intentId }));
              }),
            busy || !!draftError,
          ),
        );
      voice.append(row);
    }
    return card;
  }
  const adoptMedia = (media: Media) =>
    void run(async () => {
      if (!view || changed() || draftError || media.inPage) throw new Error('请先保存本页教学，再加入已有素材');
      const token = epoch,
        key = `${view.viewId}:${view.project.teachingRevision}:${media.kind}:${media.intentId}`;
      if (mediaAdoption?.key !== key) mediaAdoption = { key, actionId: crypto.randomUUID() };
      stopAudio();
      const result = await request('media:adopt', {
        ...scope(),
        kind: media.kind,
        intentId: media.intentId,
        actionId: mediaAdoption.actionId,
      });
      if (disposed || token !== epoch) return;
      mediaAdoption = null;
      if (media.kind === 'music') media.inPage = true;
      await applyView(result);
      notice('已有素材已加入本页，可调整播放顺序');
    });

  function renderMediaLibrary() {
    const details = element('details');
    details.append(element('summary', '添加已有音乐或视频'));
    if (!mediaLibrary.length) {
      details.append(element('p', '当前会话暂无可用的已归档音乐或视频。可在主 Chat 中生成后刷新。'));
      return details;
    }
    const select = element('select');
    select.dataset.mediaLibrary = '';
    select.setAttribute('aria-label', '已有音乐和视频');
    const placeholder = element('option', '选择已有素材');
    placeholder.value = '';
    select.append(placeholder);
    mediaLibrary.forEach((media, index) => {
      const option = element(
        'option',
        `${media.kind === 'music' ? '音乐' : '视频'} · ${media.title.slice(0, 120)}${media.inPage ? '（已在本页）' : ''}`,
      );
      option.value = String(index);
      option.disabled = media.inPage;
      select.append(option);
    });
    select.addEventListener('change', () => {
      mediaAdoption = null;
      renderControls();
    });
    const add = button(
      '加入当前页',
      () => {
        if (select.value) {
          const media = mediaLibrary[Number(select.value)];
          if (media) adoptMedia(media);
        }
      },
      true,
    );
    add.dataset.adoptMedia = '';
    details.append(select, add, element('p', '复用已有文件，不会重新生成。加入后可在时间轴中移动或删除。'));
    return details;
  }

  function editNativeAnimation(action: Action) {
    try {
      stopAudio();
      const api = getEditorApi() as any;
      const slide = api?.WordControl?.m_oLogicDocument?.Slides?.[view!.pageNumber - 1];
      const tab = nativeDocument?.querySelector<HTMLElement>(
        '[role="tab"][aria-controls="animation-panel"], #toolbar [data-tab="animate"], #animate',
      );
      if (!slide || !tab) throw new Error(t('teachingNativeEditorUnavailable'));
      const effect =
        action.type === 'native_animation' ? slide.timing?.getAllAnimEffects?.()[action.stepIndex] : undefined;
      const cue = action.visualCues?.find((cue: AttentionCue) => cue.effect?.type === 'element_animation');
      const shape = effect
        ? slide.cSld.spTree.find((shape: any) => shape.Id === effect.getObjectId())
        : cue?.effect.elementId
          ? attentionTarget(slide.cSld.spTree, view?.teachingObjects ?? [], cue.effect.elementId)
          : undefined;
      if (shape) {
        const controller = api.WordControl.m_oLogicDocument.GetCurrentController();
        controller.resetSelection();
        controller.selectObject(shape, view!.pageNumber - 1);
        api.WordControl.m_oLogicDocument.Document_UpdateSelectionState?.();
      }
      tab.click();
      setOpen(false);
    } catch (error) {
      notice(String(error), true);
    }
  }
  function moveAction(id: string, to: number) {
    if (!editable() || running || timeline) return;
    const from = actions.findIndex((action) => action.id === id);
    if (from < 0 || to < 0 || to >= actions.length || from === to) return;
    const next = [...actions],
      [action] = next.splice(from, 1);
    next.splice(to, 0, action);
    selectedActionId = id;
    edit(next);
  }
  function render(preserveLesson?: LessonRun) {
    if (!panel) return;
    stopAudio(preserveLesson);
    const scrollTop = panel.querySelector('.teaching-flow')?.scrollTop ?? 0;
    panel.replaceChildren();
    renderedEditable = editable();
    if (!actions.some((action) => action.id === selectedActionId)) selectedActionId = actions[0]?.id ?? null;
    const header = element('header');
    header.append(
      element('strong', t('teachingFlow')),
      button(
        '刷新',
        () =>
          void run(async () => {
            await saveCurrent();
            refresh();
          }),
        busy || !!draftError,
      ),
      button('停止预览', () => {
        previewRequest++;
        stopAudio();
        notice('预览已停止');
      }),
    );
    decorateTeachingButton(header.querySelectorAll('button')[0], 'refresh', true);
    decorateTeachingButton(header.querySelectorAll('button')[1], 'stop', true);
    panel.append(header);
    const status = element(
      'p',
      draftError || statusText || (busy ? '正在读取…' : t('teachingStepCount', { count: actions.length })),
    );
    status.dataset.status = '';
    status.title = status.textContent ?? '';
    status.setAttribute('role', draftError || statusError ? 'alert' : 'status');
    const statusBar = element('div');
    statusBar.className = 'teaching-statusbar';
    statusBar.append(status);
    panel.append(statusBar);
    if (!view) {
      panel.append(button('重新读取', refresh, busy));
      return;
    }
    if (draftError || draftRevision || changed() || localDrafts.has(view.pageId)) {
      const recovery = element('div');
      recovery.className = 'teaching-cue-actions';
      const discard = button(t('teachingDiscardDraft'), () => void run(discardChanges), busy || !!lesson);
      discard.className = 'teaching-secondary';
      discard.title = '删除本页未保存修改，恢复已保存内容';
      recovery.append(discard);
      panel.append(recovery);
    }
    const previewControls = element('nav');
    previewControls.setAttribute('aria-label', '本页教学播放');
    previewControls.className = 'teaching-toolbar';
    for (const [label, command, action] of [
      ['画布预览', 'play', () => playTimeline()],
      ['幻灯片放映', 'show', () => playTimeline(true)],
      [
        '放映整课教学',
        'lesson',
        () => {
          void playLesson();
        },
      ],
      ['暂停播放', 'pause', () => timeline?.pause()],
      ['继续播放', 'resume', () => timeline?.resume()],
      ['教师继续', 'teacher', () => timeline?.continueTeacher()],
    ] as const) {
      const control = button(label, action);
      control.dataset.previewCommand = command;
      if (command === 'lesson') control.className = 'teaching-primary';
      decorateTeachingButton(
        control,
        command === 'play'
          ? 'canvas'
          : command === 'show'
            ? 'slide'
            : command === 'lesson'
              ? 'lesson'
              : command === 'pause'
                ? 'pause'
                : command === 'teacher'
                  ? 'continue'
                  : 'play',
        true,
      );
      previewControls.append(control);
    }
    const history = element('nav');
    history.setAttribute('aria-label', t('teachingHistoryHint'));
    history.className = 'teaching-history';
    history.title = t('teachingHistoryHint');
    for (const [label, command, undo] of [
      [t('teachingUndo'), 'undo', true],
      [t('teachingRedo'), 'redo', false],
    ] as const) {
      const control = button(label, () => changeHistory(undo));
      control.dataset.teachingHistory = command;
      decorateTeachingButton(control, undo ? 'undo' : 'redo', true);
      history.append(control);
    }
    header.append(history);
    statusBar.append(previewControls);
    renderControls();
    if (page()?.locked) panel.append(element('p', '本页已锁定，教学内容只读。'));
    const flow = element('div');
    flow.className = 'teaching-flow';
    flow.setAttribute('role', 'list');
    flow.setAttribute('aria-label', t('teachingStepCount', { count: actions.length }));
    const inspector = element('section');
    inspector.className = 'teaching-inspector';
    inspector.dataset.teachingInspector = '';
    if (view?.teaching) {
      const kinds: Record<string, string> = {
        slide: '教学讲解',
        quiz: '提问练习',
        interactive: '课堂互动',
        pbl: '项目活动',
      };
      const info = element('details');
      info.className = 'teaching-page-info';
      const toggle = element('summary');
      const icon = button(t('teachingPageInfo'), () => {});
      decorateTeachingButton(icon, 'info', true);
      toggle.setAttribute('aria-label', t('teachingPageInfo'));
      toggle.title = t('teachingPageInfo');
      toggle.append(...icon.childNodes);
      info.append(
        toggle,
        element('p', `${kinds[view.teaching.pageType] ?? view.teaching.pageType} · ${view.teaching.objective}`),
      );
      statusBar.append(info);
    }
    panel.append(flow);
    if (!actions.length) panel.append(element('p', '本页还没有讲解，可添加讲解或在主 Chat 中生成。'));
    actions.forEach((action, index) => {
      const card = action.type === 'speech' ? speechCard(action, index) : element('article');
      if (action.type !== 'speech')
        card.append(
          element(
            'strong',
            `${index + 1} · ${action.type === 'interaction_checkpoint' ? (action.checkpointKind === 'chapter_summary' ? '课堂提问' : '暂停，等待教师继续') : action.type === 'audio_clip' ? '音乐 / 音频' : action.type === 'video_clip' ? '视频' : action.type === 'music_slot' ? '音乐 · 尚待生成' : action.type === 'delay' ? '停顿' : action.type === 'native_animation' ? `PPT 动画 · 第 ${action.stepIndex + 1} 步` : action.type}`,
          ),
        );
      if (action.type === 'interaction_checkpoint' && action.checkpointKind === 'chapter_summary') {
        for (const [labelText, value, apply] of [
          [
            '问题',
            action.question,
            (value: string, next: Action) => {
              next.question = value;
            },
          ],
          [
            '参考答案',
            action.referenceAnswer.summary,
            (value: string, next: Action) => {
              next.referenceAnswer.summary = value;
              next.context.chapterSummary = value;
            },
          ],
          [
            '答案要点（每行一项）',
            action.referenceAnswer.keyPoints.join('\n'),
            (value: string, next: Action) => {
              next.referenceAnswer.keyPoints = value
                .split('\n')
                .map((s) => s.trim())
                .filter(Boolean);
            },
          ],
        ] as const) {
          const label = element('label', labelText),
            field = element('textarea');
          field.value = value;
          field.disabled = !editable();
          field.setAttribute('aria-label', labelText);
          field.addEventListener('change', () => {
            if (!field.value.trim()) {
              notice('提问和参考答案不能为空', true);
              return;
            }
            const next = structuredClone(actions);
            apply(field.value, next[index]);
            edit(next);
          });
          label.append(field);
          card.append(label);
        }
      }
      if (action.type === 'interaction_checkpoint' && action.checkpointKind === 'chapter_summary') {
        const response = action.response ?? { kind: 'short_answer', options: [], answerIndices: [] };
        const pageId = view?.pageId;
        const saveResponse = async (nextResponse: typeof response) => {
          try {
            const runtime = await loadTeachingPreviewRuntime();
            runtime.evaluateQuestionAnswer(nextResponse, nextResponse.answerIndices);
            if (view?.pageId !== pageId || actions[index]?.id !== action.id || !editable()) return;
            const next = structuredClone(actions);
            next[index].response = nextResponse;
            edit(next);
          } catch {
            notice('题型设置无效：请填写 2–8 个不同选项和有效答案序号；判断题须有两个选项。', true);
          }
        };
        const label = element('label', '题型');
        const kind = element('select');
        kind.setAttribute('aria-label', '题型');
        for (const [value, text] of [
          ['single', '单选'],
          ['multiple', '多选'],
          ['judge', '判断'],
          ['short_answer', '简答'],
          ['discussion', '讨论'],
          ['practice', '活动练习'],
          ['pbl', '项目活动'],
        ])
          kind.append(new Option(text, value));
        kind.value = response.kind;
        kind.disabled = !editable();
        kind.addEventListener('change', () => {
          const closed = ['single', 'multiple', 'judge'].includes(kind.value);
          const options = !closed
            ? []
            : kind.value === 'judge'
              ? ['正确', '错误']
              : response.options.length >= 2
                ? response.options
                : ['选项一', '选项二'];
          void saveResponse({ kind: kind.value, options, answerIndices: closed ? [0] : [] });
        });
        label.append(kind);
        card.append(label);
        if (response.options.length) {
          const optionsLabel = element('label', '选项（每行一项）'),
            options = element('textarea');
          options.setAttribute('aria-label', '选项（每行一项）');
          options.value = response.options.join('\n');
          options.disabled = !editable();
          options.addEventListener(
            'change',
            () =>
              void saveResponse({
                ...response,
                options: options.value
                  .split('\n')
                  .map((s) => s.trim())
                  .filter(Boolean),
              }),
          );
          optionsLabel.append(options);
          card.append(optionsLabel);
          const answerLabel = element('label', '正确答案序号（从 1 开始，多选用逗号分隔）'),
            answer = element('input');
          answer.setAttribute('aria-label', '正确答案序号');
          answer.value = response.answerIndices.map((i: number) => i + 1).join(',');
          answer.disabled = !editable();
          answer.addEventListener(
            'change',
            () =>
              void saveResponse({
                ...response,
                answerIndices: answer.value.split(',').map((s) => Number(s.trim()) - 1),
              }),
          );
          answerLabel.append(answer);
          card.append(answerLabel);
        }
      }
      if (action.type === 'audio_clip' || action.type === 'video_clip') {
        const media = element('section');
        media.dataset.mediaPreview = '';
        media.className = 'teaching-media-preview';
        media.append(
          element(
            'p',
            action.durationMs
              ? t('teachingMediaDuration', { seconds: action.durationMs / 1000 })
              : t('teachingMediaDurationLoading'),
          ),
          button(t('teachingMediaReload'), () => void run(() => previewMedia(action, false)), busy || nativeSlideshow),
        );
        card.append(media);
      }
      if (action.type === 'interaction_checkpoint' && action.checkpointKind !== 'chapter_summary')
        card.append(element('p', t('teachingTeacherPauseHint')));
      if (action.type === 'music_slot') {
        card.append(element('p', action.title), element('p', '此位置的音乐尚未生成，请在主 Chat 中继续完成音乐任务。'));
      }
      if (action.type === 'delay') {
        const label = element('label', '停顿时长（毫秒）');
        const duration = element('input');
        duration.type = 'number';
        duration.min = '0';
        duration.max = '30000';
        duration.step = '100';
        duration.value = String(action.durationMs);
        duration.disabled = !editable();
        duration.addEventListener('change', () => {
          const ms = duration.valueAsNumber;
          if (!Number.isInteger(ms) || ms < 0 || ms > 30000) {
            duration.value = String(action.durationMs);
            notice('停顿时长须为 0–30000 毫秒的整数', true);
            return;
          }
          edit(actions.map((entry, i) => (i === index ? { ...entry, durationMs: ms } : entry)));
        });
        label.append(duration);
        card.append(label);
      }
      card.dataset.actionId = action.id;
      if (action.type === 'native_animation')
        card.append(
          button(t('teachingEditAnimation'), () => editNativeAnimation(action), !editable() || nativeSlideshow),
        );
      const controls = element('div');
      controls.className = 'teaching-node-controls';
      controls.append(
        button(
          '上移',
          () => {
            moveAction(action.id, index - 1);
          },
          !editable() || index === 0,
        ),
        button(
          '下移',
          () => {
            moveAction(action.id, index + 1);
          },
          !editable() || index === actions.length - 1,
        ),
        button(
          '删除',
          () => edit(actions.filter((_, i) => i !== index)),
          !editable() || action.type === 'native_animation',
        ),
      );
      for (const [i, control] of [...controls.querySelectorAll('button')].entries())
        decorateTeachingButton(control, (['up', 'down', 'delete'] as const)[i], true);
      const row = element('article');
      row.dataset.actionId = action.id;
      row.setAttribute('role', 'listitem');
      if (action.id === selectedActionId) row.dataset.selected = '';
      const choose = button('', () => {
        if (timeline || running) return;
        selectedActionId = action.id;
        render();
      });
      choose.className = 'teaching-step';
      choose.setAttribute('aria-pressed', String(action.id === selectedActionId));
      const number = element('span', String(index + 1));
      number.className = 'teaching-step-number';
      const copy = element('span'),
        summary = element('span', teachingStepSummary(action));
      copy.className = 'teaching-step-copy';
      summary.className = 'teaching-step-summary';
      copy.append(element('strong', teachingStepTitle(action)), summary);
      if (action.type === 'speech' && action.pauseAfterMs) {
        const badge = element('span', speechPauseLabel(action.pauseAfterMs));
        badge.className = 'teaching-step-badge';
        copy.append(badge);
      }
      for (const type of ['spotlight', 'highlight']) {
        if (action.visualCues?.some((cue: AttentionCue) => cue.effect?.type === type)) {
          const badge = element(
            'span',
            t(type === 'spotlight' ? 'teachingSpotlightEffect' : 'teachingHighlightEffect'),
          );
          badge.className = 'teaching-step-badge';
          copy.append(badge);
        }
      }
      choose.append(number, copy);
      row.append(choose, controls);
      choose.draggable = editable() && !timeline && !running;
      row.addEventListener('dragstart', (event) => {
        if (!choose.draggable) {
          event.preventDefault();
          return;
        }
        draggedActionId = action.id;
        event.dataTransfer?.setData('text/plain', action.id);
      });
      row.addEventListener('dragover', (event) => {
        if (draggedActionId && editable() && !timeline && !running) {
          event.preventDefault();
          row.dataset.dropTarget = '';
        }
      });
      row.addEventListener('dragleave', () => delete row.dataset.dropTarget);
      row.addEventListener('drop', (event) => {
        if (!draggedActionId) return;
        event.preventDefault();
        delete row.dataset.dropTarget;
        if (draggedActionId) moveAction(draggedActionId, index);
        draggedActionId = null;
      });
      row.addEventListener('dragend', () => {
        draggedActionId = null;
        panel?.querySelectorAll<HTMLElement>('[data-drop-target]').forEach((node) => delete node.dataset.dropTarget);
      });
      flow.append(row);
      if (action.id === selectedActionId) {
        inspector.dataset.actionId = action.id;
        inspector.setAttribute(
          'aria-label',
          t('teachingStepDetail', { step: index + 1, title: teachingStepTitle(action) }),
        );
        card.firstElementChild?.remove();
        inspector.append(...card.childNodes);
        row.append(inspector);
        if (action.type === 'audio_clip' || action.type === 'video_clip')
          queueMicrotask(() => {
            if (
              disposed ||
              lesson ||
              nativeSlideshow ||
              busy ||
              previewAudio?.isConnected ||
              selectedActionId !== action.id ||
              changed()
            )
              return;
            const token = epoch;
            void previewMedia(action, false).catch((error) => {
              if (!disposed && token === epoch && selectedActionId === action.id)
                notice(`${t('teachingMediaFailed')} ${String(error)}`, true);
            });
          });
      }
    });
    const footer = element('footer');
    footer.append(
      button(
        t('teachingAddAfter'),
        () => {
          const index = actions.findIndex((action) => action.id === selectedActionId);
          const next = [...actions];
          const id = crypto.randomUUID();
          next.splice(index + 1, 0, { id, type: 'speech', text: '' });
          selectedActionId = id;
          edit(next);
          panel?.querySelector<HTMLTextAreaElement>('[data-teaching-inspector] textarea')?.focus();
        },
        !editable() || actions.length >= 1000,
      ),
    );
    const addSpeech = footer.querySelector('button')!;
    decorateTeachingButton(addSpeech, 'plus');
    addSpeech.querySelector('span')!.textContent = t('teachingAddSpeech');
    footer.append(renderMediaLibrary());
    const advanced = element('details');
    advanced.append(element('summary', t('teachingEditAnimation')));
    advanced.append(
      button(
        '绑定本页 PPT 动画',
        () => {
          try {
            const source = view?.project.source?.sha256;
            if (!source || deckwisePageLineage.context()?.sourceSha256 !== source)
              throw new Error('请先保存并重新打开 PPT，再绑定当前页动画');
            const { effects } = inspectTeachingAnimations(view!.pageNumber);
            if (!effects.length) throw new Error('本页没有可绑定的 PPT 动画，请先用原生动画功能添加并保存');
            if (actions.length + effects.length > 1000) throw new Error('本页教学节点超过数量限制');
            edit([
              ...actions,
              ...effects.map((_, stepIndex) => ({
                id: crypto.randomUUID(),
                type: 'native_animation',
                version: 1,
                sourceSha256: source,
                stepIndex,
              })),
            ]);
            notice('已添加本页全部 PPT 动画步骤，可上下移动节点，将讲解或媒体排列到动画之间。');
          } catch (error) {
            notice(String(error), true);
          }
        },
        !editable() || nativeSlideshow || actions.some((action) => action.type === 'native_animation'),
      ),
    );
    if (actions.some((action) => action.type === 'native_animation'))
      advanced.append(
        button('移除动画绑定', () => edit(actions.filter((action) => action.type !== 'native_animation')), !editable()),
      );
    footer.append(advanced);
    const saveButton = button('保存本页教学', save, !editable() || !changed());
    saveButton.dataset.save = '';
    decorateTeachingButton(saveButton, 'save', true);
    history.append(saveButton);
    panel.append(footer);
    flow.scrollTop = scrollTop;
    flow.querySelector<HTMLElement>('[data-selected]')?.scrollIntoView?.({ block: 'nearest', inline: 'nearest' });
  }

  function setOpen(expanded: boolean) {
    if (disposed) return;
    const visibilityChanged = open !== expanded;
    if (expanded && panel && (nativeDocument?.getElementById('right-menu')?.getBoundingClientRect().width ?? 0) <= 50)
      nativeDocument?.getElementById('id-right-menu-slide')?.click();
    open = expanded;
    if (!open) {
      clearTimeout(statusTimer);
      stopAudio();
    }
    if (panel) panel.hidden = !open;
    sidebar.update(true, open);
    panelLayout?.sync(open);
    send('document:ui-changed', { teachingOpen: open });
    if (visibilityChanged) {
      lifecycle?.update();
      if (open) pollTasks();
    }
  }
  const sidebar = createSidebarEntry(
    () => setOpen(!open),
    () => '教学时间轴',
  );
  const mount = () => {
    for (const frame of document.querySelectorAll<HTMLIFrameElement>('#app iframe')) {
      const doc = frame.contentDocument;
      const container = doc?.querySelector<HTMLElement>('#right-menu .right-panel');
      if (!doc || !container) continue;
      if (panel?.ownerDocument === doc) return;
      panel?.remove();
      stopNativeClicks?.();
      panelLayout?.dispose();
      nativeDocument = doc;
      panelLayout = createTeachingPanelLayout(doc);
      const style = element('style');
      style.textContent =
        `
        .right-panel:has(> .deckwise-teaching-sidebar) { position:relative; height:100%; }
        .deckwise-teaching-sidebar { position:absolute; inset:0 40px 0 0; z-index:20; background:var(--background-normal,#fff); color:var(--text-normal,#222); overflow:auto; padding:12px; font:13px/1.5 Arial,sans-serif; user-select:text; }
        .deckwise-teaching-sidebar[hidden] { display:none; }
        .deckwise-teaching-sidebar header { display:flex; flex-wrap:wrap; align-items:center; gap:6px; }
        .deckwise-teaching-sidebar header strong { flex:1 0 100%; }
        .deckwise-teaching-sidebar .teaching-attention { margin-top:12px; padding-top:10px; border-top:1px solid #ddd; }
        .deckwise-teaching-sidebar .teaching-attention fieldset { max-height:180px; overflow:auto; margin:8px 0; border:1px solid #bbb; }
        .deckwise-teaching-sidebar .teaching-target { display:flex; align-items:center; gap:6px; }
        .deckwise-teaching-sidebar .teaching-target input { display:inline-block; width:auto; flex:none; }
        .deckwise-teaching-sidebar article[data-playing] { box-shadow:inset 3px 0 #3977c5; padding-left:8px; }
        .deckwise-teaching-sidebar article { padding:12px 0; border-top:1px solid var(--border-regular-control,#ddd); }
        .deckwise-teaching-sidebar textarea,.deckwise-teaching-sidebar input,.deckwise-teaching-sidebar select { display:block; width:100%; box-sizing:border-box; margin:6px 0; padding:6px; color:inherit; background:var(--background-normal,#fff); border:1px solid #bbb; border-radius:4px; user-select:text; }
        .deckwise-teaching-sidebar textarea { resize:vertical; }
        .deckwise-teaching-sidebar button { border:1px solid #bbb; border-radius:4px; background:var(--background-normal,#fff); color:inherit; padding:5px 7px; margin:3px 3px 3px 0; cursor:pointer; }
        .deckwise-teaching-sidebar button:disabled { opacity:.5; cursor:default; }
        .deckwise-teaching-sidebar :focus-visible { outline:2px solid #3977c5; outline-offset:1px; }
        .deckwise-teaching-sidebar [role=alert] { color:#b22; }
        .deckwise-teaching-sidebar p { margin:6px 0; overflow-wrap:anywhere; }
        .deckwise-teaching-sidebar footer { padding-top:12px; border-top:1px solid #ddd; }
      ` + teachingFlowStyles;
      doc.head.append(style);
      panel = element('aside');
      panel.className = 'deckwise-teaching-sidebar';
      panel.setAttribute('aria-label', '教学时间轴');
      panel.hidden = !open;
      // Keep native canvas shortcuts from consuming typing, selection and media controls.
      for (const type of ['keydown', 'keyup', 'keypress', 'pointerdown', 'mousedown', 'wheel'])
        panel.addEventListener(type, (event) => event.stopPropagation());
      container.append(panel);
      const nativeClick = (event: Event) => {
        const target = event.target as Element;
        if (target.closest('#right-menu .tool-menu-btns button:not(.agent-sidebar-entry)')) {
          setOpen(false);
        }
      };
      doc.addEventListener('click', nativeClick);
      const api = getEditorApi();
      const isSlideShow = () => typeof api?.isSlideShow === 'function' && api.isSlideShow() === true;
      const slideshowChanged = () => {
        nativeSlideshow = isSlideShow();
        if (startingSlideshow || teachingSlideshow) return;
        stopAudio();
        if (nativeSlideshow) notice('原生放映已开始，教学预览已停止；原生放映暂不自动播放教学时间轴。');
        else notice('已退出原生放映，可重新预览本页教学');
      };
      const slideshowEvents = ['asc_onStartDemonstration', 'asc_onDemonstrationSlideChanged', 'asc_onEndDemonstration'];
      if (api?.asc_registerCallback && api.asc_unregisterCallback)
        for (const event of slideshowEvents) api.asc_registerCallback(event, slideshowChanged);
      nativeSlideshow = isSlideShow();
      stopNativeClicks = () => {
        for (const event of slideshowEvents) api?.asc_unregisterCallback?.(event, slideshowChanged);
        doc.removeEventListener('click', nativeClick);
        style.remove();
      };
      setOpen(open);
      render();
      send('document:teaching-ready', {});
      return;
    }
  };
  lifecycle = mountPanelContext({
    isActive: () => !disposed && open,
    refresh: () => {
      if (disposed || !open) return;
      mount();
      try {
        if (lesson?.navigating || projectingMotions) return;
        if (hasUnsavedChanges() && !hasNativeTeaching()) {
          notice('PPT 正在自动保存，完成后将同步教学时间轴。');
          return;
        }
        const next = deckwiseEditorBridge.getContext(binding);
        const native = hasNativeTeaching() ? readNativeTeaching() : undefined;
        if (native && view && next.page === view.pageNumber && native.revision === view.project.teachingRevision) {
          context = next;
          view.teachingObjects = nativeTeachingObjects(next.page ?? 1).map(({ nativeShape, ...object }) => ({ ...object, aliases: [] }));
          return;
        }
        if (!context || next.page !== context.page || next.editGeneration !== context.editGeneration ||
            (native && native.revision !== view?.project.teachingRevision)) void load(next);
      } catch {
        /* The next SDK context event retries after initialization. */
      }
    },
  });
  const nativeChanged = () => queueMicrotask(() => lifecycle?.update());
  window.addEventListener('deckwise-native-teaching-changed', nativeChanged);
  sidebar.update(true, true);
  lifecycle?.update();
  pollTasks();
  return {
    refresh,
    setOpen,
    receive(id: string, payload: any) {
      const entry = pending.get(id);
      if (!entry) return;
      pending.delete(id);
      clearTimeout(entry.timer);
      if (payload.error) entry.reject(new Error(payload.error));
      else entry.resolve(payload.result);
    },
    dispose() {
      window.removeEventListener('deckwise-native-teaching-changed', nativeChanged);
      disposed = true;
      epoch++;
      clearTimeout(statusTimer);
      clearTimeout(saveTimer);
      stopAudio();
      sidebar.dispose();
      panelLayout?.dispose();
      lifecycle?.dispose();
      stopNativeClicks?.();
      panel?.remove();
      for (const entry of pending.values()) {
        clearTimeout(entry.timer);
        entry.reject(new Error('Teaching editor closed'));
      }
      pending.clear();
    },
  };
}
