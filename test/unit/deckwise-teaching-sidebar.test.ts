import { LanguageCode, setLanguage } from '@ranuts/shared/i18n';
import { afterEach, beforeEach, expect, it, vi } from 'vitest';
const state = vi.hoisted(() => ({
  refresh: () => {},
  toggle: () => {},
  context: {
    projectId: 'project',
    documentId: 'document',
    sessionId: 'session',
    page: 1,
    editGeneration: 0,
    readOnly: false,
  },
}));
vi.mock('../../lib/agent-plugin/ui/panel-context', () => ({
  mountPanelContext: (options: any) => {
    state.refresh = options.refresh;
    return { update: options.refresh, dispose: vi.fn() };
  },
}));
vi.mock('../../lib/agent-plugin/ui/sidebar-entry', () => ({
  createSidebarEntry: (toggle: () => void) => {
    state.toggle = toggle;
    return { update: vi.fn(), dispose: vi.fn() };
  },
}));
vi.mock('../../lib/deckwise-editor-bridge', () => ({
  deckwiseEditorBridge: { getContext: () => ({ ...state.context }) },
}));
const previewState = vi.hoisted(() => ({ load: vi.fn() }));
vi.mock('../../lib/deckwise-teaching-preview', async (importOriginal) => ({
  ...(await importOriginal<typeof import('../../lib/deckwise-teaching-preview')>()),
  loadTeachingPreviewRuntime: previewState.load,
}));
const slideshowState = vi.hoisted(() => ({
  active: false,
  callbacks: new Map<string, () => void>(),
  api: null as any,
}));
vi.mock('../../lib/agent-plugin/editor-bridge', () => ({
  getEditorApi: () =>
    slideshowState.api ?? {
      isSlideShow: () => slideshowState.active,
      asc_registerCallback: (name: string, callback: () => void) => slideshowState.callbacks.set(name, callback),
      asc_unregisterCallback: (name: string) => slideshowState.callbacks.delete(name),
    },
}));
import { deckwisePageLineage } from '../../lib/deckwise-page-lineage';
import { mountDeckwiseTeachingSidebar } from '../../lib/deckwise-teaching-sidebar';
let sidebar: ReturnType<typeof mountDeckwiseTeachingSidebar>;
let doc: Document;
let calls: Array<{ operation: string; args: any; context: any }>;
let persisted: any[];
let library: Array<{ kind: string; intentId: string; title: string; model: string; inPage: boolean }>;
let adoptionFailure: boolean;
let writeFailure: boolean;
let deferPreview: boolean;
let delayedPreviewId: string | undefined;
let sourceSha: string | undefined;
let teachingRevision: string;
let recoveryOnSecond: boolean;
const first = () => ({
  id: 'speech',
  type: 'speech',
  text: 'Hello',
  audioAssetId: 'tts-old',
  durationMs: 1000,
  visualCues: [{ id: 'cue' }],
});
const flush = async () => {
  for (let i = 0; i < 20; i++) await Promise.resolve();
};
beforeEach(() => {
  setLanguage(LanguageCode.ZH);
  slideshowState.api = null;
  sourceSha = undefined;
  teachingRevision = 'rev';
  recoveryOnSecond = false;
  slideshowState.active = false;
  slideshowState.callbacks.clear();
  state.context = { ...state.context, page: 1, editGeneration: 0 };
  document.body.innerHTML = '<div id="app"><iframe></iframe></div>';
  doc = document.querySelector('iframe')!.contentDocument!;
  doc.body.innerHTML =
    '<div id="right-menu"><div class="right-panel"></div><div class="tool-menu-btns"><button id="id-right-menu-slide"></button><button id="id-right-menu-image"></button></div></div>';
  calls = [];
  library = [];
  adoptionFailure = false;
  writeFailure = false;
  deferPreview = false;
  delayedPreviewId = undefined;
  persisted = [first()];
  const project = () => ({
    teachingRevision,
    source: sourceSha ? { sha256: sourceSha } : undefined,
    pages: [
      { pageId: 'p1', locked: false, actions: persisted },
      {
        pageId: 'p2',
        locked: false,
        actions: [{ id: 's2', type: 'speech', text: 'Page two', audioAssetId: 'second-tts' }],
      },
    ],
  });
  sidebar = mountDeckwiseTeachingSidebar(state.context, (type, payload, id) => {
    if (type !== 'document:teaching-request') return;
    calls.push(payload);
    queueMicrotask(() => {
      let result: any;
      switch (payload.operation) {
        case 'open':
          result = {
            viewId: 'view-' + payload.context.page,
            project: project(),
            pageId: 'p' + payload.context.page,
            pageNumber: payload.context.page,
          };
          break;
        case 'draft:get':
        case 'draft:set':
          result = {
            draft:
              recoveryOnSecond && payload.operation === 'draft:get' && payload.args.viewId === 'view-2'
                ? {
                    revision: 'draft-2',
                    baseline: project(),
                    actions: [{ id: 'unsaved', type: 'speech', text: 'Recovered edit' }],
                  }
                : null,
          };
          break;
        case 'tts:preview':
        case 'media:preview':
          if (deferPreview) {
            delayedPreviewId = id;
            return;
          }
          result = {
            bytes: new Uint8Array([1, 2, 3]),
            mimeType: payload.args.actionId === 'video' ? 'video/mp4' : 'audio/mpeg',
          };
          break;
        case 'media:adopt':
          if (adoptionFailure) {
            adoptionFailure = false;
            sidebar.receive(id!, { error: 'adoption acknowledgement lost' });
            return;
          }
          persisted.push({ id: payload.args.actionId, type: 'video_clip', videoAssetId: 'video-existing' });
          result = { viewId: 'view-1', project: project(), pageId: 'p1', pageNumber: 1 };
          break;
        case 'media:list':
          result = library;
          break;
        case 'tts:models':
        case 'tts:list':
          result = [];
          break;
        case 'write':
          if (writeFailure) {
            sidebar.receive(id!, { error: 'save conflict' });
            return;
          }
          persisted = structuredClone(payload.args.actions);
          result = { viewId: 'view-1', project: project(), pageId: 'p1', pageNumber: 1 };
          break;
      }
      sidebar.receive(id!, { result });
    });
  });
});
afterEach(() => {
  sidebar.dispose();
  vi.unstubAllGlobals();
  vi.useRealTimers();
  document.body.replaceChildren();
  setLanguage(LanguageCode.EN);
});
const click = (label: string) => {
  const button = [...doc.querySelectorAll('button')].find(
    (node) => node.getAttribute('aria-label') === label || node.textContent === label,
  )!;
  expect(button).toBeTruthy();
  expect(button.disabled).toBe(false);
  button.click();
};
it('automatically loads narration and voice in the native property area without a paid request', async () => {
  await flush();
  expect(doc.querySelector('.right-panel [aria-label="教学时间轴"]')).toBeTruthy();
  expect(doc.querySelector('textarea')!.value).toBe('Hello');
  expect(doc.body.textContent).toContain('语音已就绪');
  expect(calls.map((call) => call.operation)).toEqual(['open', 'draft:get', 'tts:models', 'tts:list', 'media:list']);
});
it('edits and saves narration through Host while invalidating stale audio and cues', async () => {
  await flush();
  const input = doc.querySelector('textarea')!;
  input.value = 'New narration';
  input.dispatchEvent(new Event('input'));
  await flush();
  click('保存本页教学');
  await flush();
  expect(persisted).toEqual([{ id: 'speech', type: 'speech', text: 'New narration' }]);
  expect(calls.some((call) => call.operation === 'draft:set' && call.args.draft === null)).toBe(true);
  expect(doc.querySelector('textarea')!.value).toBe('New narration');
});
it('inserts narration after selection and follows page changes with its recovery draft', async () => {
  await flush();
  click('在选中步骤后添加讲解');
  await flush();
  expect(doc.querySelectorAll('article')).toHaveLength(2);
  const draft = calls.find((call) => call.operation === 'draft:set')!;
  expect(draft.args.draft.actions[1].type).toBe('speech');
  state.context.page = 2;
  state.refresh();
  await flush();
  expect(doc.querySelector('textarea')!.value).toBe('Page two');
  expect(calls.filter((call) => call.operation === 'open').map((call) => call.context.page)).toEqual([1, 2]);
  state.context.page = 1;
  state.refresh();
  await flush();
  expect(doc.querySelectorAll('article')).toHaveLength(2);
});
it('native property buttons and teaching share one area and typing never reaches canvas shortcuts', async () => {
  await flush();
  const panel = doc.querySelector<HTMLElement>('.deckwise-teaching-sidebar')!;
  doc.getElementById('id-right-menu-image')!.click();
  expect(panel.hidden).toBe(true);
  state.toggle();
  expect(panel.hidden).toBe(false);
  const canvasShortcut = vi.fn();
  doc.addEventListener('keydown', canvasShortcut);
  doc.querySelector('textarea')!.dispatchEvent(new KeyboardEvent('keydown', { key: ' ', bubbles: true }));
  expect(canvasShortcut).not.toHaveBeenCalled();
  doc.removeEventListener('keydown', canvasShortcut);
});

it('records highlighter cues against eligible Host object IDs without changing audio', async () => {
  await flush();
  // Reopen with a mapped, ready page to exercise the native cue editor.
  sidebar.dispose();
  const cueCalls: any[] = [];
  sidebar = mountDeckwiseTeachingSidebar(state.context, (type, payload, id) => {
    if (type !== 'document:teaching-request') return;
    cueCalls.push(payload);
    queueMicrotask(() => {
      const result =
        payload.operation === 'open'
          ? {
              viewId: 'cue-view',
              pageId: 'p1',
              pageNumber: 1,
              teachingObjects: ['panda', 'monkey'].map((id) => ({
                objectId: id,
                sourceObjectName: id,
                name: id,
                status: 'mapped',
                effectTargetEligible: true,
              })),
              project: {
                teachingRevision: 'rev',
                pages: [
                  { pageId: 'p1', locked: false, reviewState: 'ready', actions: [{ ...first(), visualCues: [] }] },
                ],
              },
            }
          : payload.operation.startsWith('draft:')
            ? { draft: null }
            : [];
      sidebar.receive(id!, { result });
    });
  });
  await flush();
  vi.useFakeTimers();
  click('添加高亮');
  (doc.querySelector('.teaching-target input') as HTMLInputElement).checked = true;
  click('添加教学提示');
  await flush();
  const actions = cueCalls.find((call) => call.operation === 'draft:set').args.draft.actions;
  expect(actions[0].audioAssetId).toBe('tts-old');
  expect(actions[0].visualCues).toEqual([
    {
      id: expect.any(String),
      atMs: 0,
      endMs: 1000,
      effect: { type: 'highlight', elementIds: ['panda'], style: 'outline', color: '#f59e0b', opacity: 0.8 },
    },
  ]);
  expect(doc.querySelector('.teaching-attention')).not.toBeInstanceOf(HTMLDetailsElement);
  expect(doc.querySelector('[data-cue-id]')?.textContent).toContain('panda');
  const originalId = actions[0].visualCues[0].id;
  click('编辑此提示');
  await vi.advanceTimersByTimeAsync(2400);
  expect(cueCalls.some((call) => call.operation === 'write')).toBe(false);
  expect(doc.querySelector<HTMLElement>('[data-cue-form]')!.hidden).toBe(false);
  expect((doc.querySelector('[aria-label="结束 (秒)"]') as HTMLInputElement).value).toBe('1');
  (doc.querySelector('[aria-label="教学提示类型"]') as HTMLSelectElement).value = 'spotlight';
  doc.querySelector('[aria-label="教学提示类型"]')!.dispatchEvent(new Event('change'));
  for (const choice of doc.querySelectorAll<HTMLInputElement>('.teaching-target input')) choice.checked = true;
  (doc.querySelector('[aria-label="开始 (秒)"]') as HTMLInputElement).value = '0.2';
  (doc.querySelector('[aria-label="结束 (秒)"]') as HTMLInputElement).value = '0.2';
  const writes = () => cueCalls.filter((call) => call.operation === 'draft:set');
  const before = writes().length;
  click('更新教学提示');
  await flush();
  expect(writes()).toHaveLength(before);
  expect(doc.querySelector('[role="alert"]')?.textContent).toContain('有效的时间');
  (doc.querySelector('[aria-label="结束 (秒)"]') as HTMLInputElement).value = '0.9';
  (doc.querySelector('[aria-label="效果不透明度百分比"]') as HTMLInputElement).value = '65';
  click('更新教学提示');
  await flush();
  expect(writes().at(-1).args.draft.actions[0]).toMatchObject({
    audioAssetId: 'tts-old',
    visualCues: [
      {
        id: originalId,
        atMs: 200,
        endMs: 900,
        effect: { type: 'spotlight', elementId: 'panda', elementIds: ['monkey'], dimOpacity: 0.65 },
      },
    ],
  });
  expect(doc.querySelector('[data-cue-id]')?.textContent).toContain('聚光');
  click('删除此提示');
  await flush();
  expect(writes().at(-1).args.draft.actions[0].visualCues).toEqual([]);
  click('撤销');
  await flush();
  expect(writes().at(-1).args.draft.actions[0].visualCues[0].id).toBe(originalId);
  click('删除此提示');
  await flush();
  const selected = ['panda', 'monkey'].map((name) => ({ getUniNvProps: () => ({ cNvPr: { name } }) }));
  slideshowState.api = {
    WordControl: {
      m_oLogicDocument: {
        Slides: [{ cSld: { spTree: selected } }],
        GetCurrentController: () => ({ selectedObjects: selected }),
      },
    },
  };
  click('添加聚光');
  expect(doc.querySelector<HTMLSelectElement>('[aria-label="教学提示类型"]')!.value).toBe('spotlight');
  expect([...doc.querySelectorAll<HTMLInputElement>('.teaching-target input')].every((input) => input.checked)).toBe(
    true,
  );
  click('添加教学提示');
  await flush();
  expect(writes().at(-1).args.draft.actions[0]).toMatchObject({
    audioAssetId: 'tts-old',
    visualCues: [{ effect: { type: 'spotlight', elementId: 'panda', elementIds: ['monkey'] } }],
  });
  expect(cueCalls.some((call) => ['tts:submit', 'tts:prepare', 'write'].includes(call.operation))).toBe(false);
});

it('allows stopping during preparation and ignores the late runtime without reading or starting audio', async () => {
  await flush();
  let resolve!: (runtime: any) => void;
  previewState.load.mockReturnValue(
    new Promise((done) => {
      resolve = done;
    }),
  );
  click('画布预览');
  expect(doc.querySelector<HTMLElement>('.deckwise-teaching-sidebar')!.inert).not.toBe(true);
  click('停止预览');
  const create = vi.fn();
  resolve({ createTeachingTimelinePreview: create });
  await flush();
  expect(create).not.toHaveBeenCalled();
  expect(calls.some((call) => call.operation === 'tts:preview')).toBe(false);
  expect(doc.querySelector('[data-status]')!.textContent).toBe('预览已停止');
});

it('prepares adopted music and video in order, then stops and releases their local URLs', async () => {
  await flush();
  persisted = [
    { id: 'music', type: 'audio_clip', audioAssetId: 'music-saved' },
    { id: 'video', type: 'video_clip', videoAssetId: 'video-saved' },
  ];
  state.context.editGeneration++;
  state.refresh();
  await flush();
  const revoke = vi.fn();
  let urlId = 0;
  vi.stubGlobal(
    'URL',
    class extends URL {
      static createObjectURL = () => 'blob:local-' + ++urlId;
      static revokeObjectURL = revoke;
    },
  );
  const dispose = vi.fn(),
    start = vi.fn(),
    create = vi.fn((_input: { nodes: Array<{ type: string }>; video?: HTMLVideoElement }) => ({ start, dispose }));
  previewState.load.mockResolvedValue({ createTeachingTimelinePreview: create });
  click('画布预览');
  await flush();
  expect(doc.querySelector('[role=alert]')?.textContent).toBeUndefined();
  expect(calls.filter((c) => c.operation === 'media:preview').map((c) => c.args.actionId)).toEqual(['music', 'video']);
  expect(create.mock.calls[0][0].nodes.map((n) => n.type)).toEqual(['audio_clip', 'video_clip']);
  expect(create.mock.calls[0][0].video!.getAttribute('aria-label')).toBe('本页教学视频');
  expect(start).toHaveBeenCalledOnce();
  click('停止预览');
  expect(dispose).toHaveBeenCalledOnce();
  expect(revoke).toHaveBeenCalledTimes(2);
  expect(doc.querySelector('video')).toBeNull();
  vi.unstubAllGlobals();
});

it('selects an existing video, preserves the insertion identity after an unknown response, and shows the saved node', async () => {
  library = [{ kind: 'video', intentId: 'existing-video', title: 'Classroom greeting', model: 'local', inPage: false }];
  await flush();
  const select = doc.querySelector<HTMLSelectElement>('[data-media-library]')!;
  select.value = '0';
  select.dispatchEvent(new Event('change'));
  adoptionFailure = true;
  click('加入当前页');
  await flush();
  expect(doc.querySelector('[role=alert]')!.textContent).toContain('acknowledgement lost');
  click('加入当前页');
  await flush();
  const requests = calls.filter((c) => c.operation === 'media:adopt');
  expect(requests).toHaveLength(2);
  expect(requests[0].args.actionId).toBe(requests[1].args.actionId);
  expect(requests[1].args).toMatchObject({
    viewId: 'view-1',
    expectedTeachingRevision: 'rev',
    kind: 'video',
    intentId: 'existing-video',
  });
  expect(persisted.at(-1).id).toBe(requests[1].args.actionId);
  expect(doc.querySelectorAll('article')).toHaveLength(2);
  expect(doc.querySelector('[data-status]')!.textContent).toBe('已有素材已加入本页，可调整播放顺序');
  expect(calls.some((c) => c.operation === 'tts:submit')).toBe(false);
});
it('requires saving pending edits before adding an existing media asset', async () => {
  library = [{ kind: 'video', intentId: 'existing-video', title: 'Classroom greeting', model: 'local', inPage: false }];
  await flush();
  const select = doc.querySelector<HTMLSelectElement>('[data-media-library]')!;
  select.value = '0';
  select.dispatchEvent(new Event('change'));
  const textarea = doc.querySelector('textarea')!;
  textarea.value = 'Unsaved lesson';
  textarea.dispatchEvent(new Event('input'));
  expect(doc.querySelector<HTMLButtonElement>('[data-adopt-media]')!.disabled).toBe(true);
  expect(calls.some((c) => c.operation === 'media:adopt')).toBe(false);
});

const setAfter = (mode: string) => {
  const select = doc.querySelector<HTMLSelectElement>('[aria-label="讲后"]')!;
  expect(select).toBeTruthy();
  select.value = mode;
  select.dispatchEvent(new Event('change'));
};
const latestDraft = () =>
  calls.filter((call) => call.operation === 'draft:set' && call.args.draft).at(-1)!.args.draft.actions;

it('removes standalone pause creation and edits only the selected narration tail', async () => {
  await flush();
  expect(doc.querySelector('.teaching-flow')).toBeTruthy();
  expect(doc.querySelectorAll('[data-teaching-inspector]')).toHaveLength(1);
  expect(doc.querySelector('[data-teaching-inspector]')!.parentElement!.matches('article[data-selected]')).toBe(true);
  expect(doc.querySelector('.teaching-flow')!.contains(doc.querySelector('[data-teaching-inspector]'))).toBe(true);
  expect(
    [...doc.querySelectorAll('button')].some((button) =>
      ['添加停顿', '添加等待教师继续', '课堂提问'].includes(button.textContent!),
    ),
  ).toBe(false);
  expect(doc.querySelector<HTMLSelectElement>('[aria-label="讲后"]')!.value).toBe('none');
  setAfter('pause');
  await flush();
  expect(latestDraft()[0]).toEqual({ ...first(), pauseAfterMs: 400 });
  expect(doc.querySelector('.teaching-flow')!.textContent).toContain('讲后暂停');
  expect(doc.querySelectorAll('article')).toHaveLength(1);
  expect(doc.querySelector('[data-teaching-inspector] input[type=number]')).toBeNull();
  setAfter('timed');
  await flush();
  expect(latestDraft()[0]).toEqual({ ...first(), pauseAfterMs: 3000 });
  const seconds = doc.querySelector<HTMLInputElement>('[data-teaching-inspector] input[type=number]')!;
  expect(seconds.value).toBe('3');
  expect(seconds.min).toBe('3');
  expect(seconds.max).toBe('30');
  for (const value of ['2', '', '31', 'NaN']) {
    seconds.value = value;
    seconds.dispatchEvent(new Event('change'));
    expect(seconds.value).toBe('3');
    expect(latestDraft()[0].pauseAfterMs).toBe(3000);
  }
  seconds.value = '4.5';
  seconds.dispatchEvent(new Event('change'));
  await flush();
  expect(latestDraft()[0]).toEqual({ ...first(), pauseAfterMs: 4500 });
  click('保存本页教学');
  await flush();
  expect(persisted[0]).toEqual({ ...first(), pauseAfterMs: 4500 });
  const create = vi.fn((_input: { nodes: Array<Record<string, unknown>> }) => ({ start: vi.fn(), dispose: vi.fn() }));
  previewState.load.mockResolvedValue({ createTeachingTimelinePreview: create });
  click('画布预览');
  await flush();
  expect(create.mock.calls[0][0].nodes[0]).toMatchObject({ type: 'speech', pauseAfterMs: 4500 });
  expect(calls.filter((c) => c.operation === 'media:preview')).toEqual([]);
  click('停止预览');
  setAfter('none');
  await flush();
  expect(latestDraft()[0]).toEqual({ ...first(), pauseAfterMs: 0 });
});

it('keeps a saved narration tail undoable and preserves the audio and cues', async () => {
  await flush();
  setAfter('pause');
  await flush();
  click('保存本页教学');
  await flush();
  click('撤销');
  await flush();
  expect(latestDraft()[0]).toEqual(first());
  click('重做');
  await flush();
  expect(latestDraft()[0]).toEqual({ ...first(), pauseAfterMs: 400 });
});

it('inserts at the selected step and keeps one detail editor while reordering', async () => {
  await flush();
  click('在选中步骤后添加讲解');
  await flush();
  const second = latestDraft()[1].id;
  doc.querySelector<HTMLButtonElement>('.teaching-step')!.click();
  click('在选中步骤后添加讲解');
  await flush();
  const inserted = latestDraft()[1].id;
  expect(latestDraft().map((action: any) => action.id)).toEqual(['speech', inserted, second]);
  expect(doc.querySelectorAll('textarea')).toHaveLength(1);
  const row = doc.querySelectorAll('article')[1];
  [...row.querySelectorAll('button')].find((button) => button.textContent === '下移')!.click();
  await flush();
  expect(latestDraft().map((action: any) => action.id)).toEqual(['speech', second, inserted]);
  expect(doc.querySelector('[data-teaching-inspector]')!.getAttribute('data-action-id')).toBe(inserted);
  const source = doc.querySelectorAll('article')[2].querySelector('.teaching-step')!;
  source.dispatchEvent(new Event('dragstart', { bubbles: true }));
  doc.querySelectorAll('article')[0].dispatchEvent(new Event('drop', { cancelable: true }));
  await flush();
  expect(latestDraft().map((action: any) => action.id)).toEqual([inserted, 'speech', second]);
  expect(doc.querySelector('[data-teaching-inspector]')!.getAttribute('data-action-id')).toBe(inserted);
});

it('undoes draft text including its original audio and cues, redoes, and clears redo after a new edit', async () => {
  await flush();
  const input = doc.querySelector('textarea')!;
  input.value = 'Changed';
  input.dispatchEvent(new Event('input'));
  await flush();
  click('撤销');
  await flush();
  expect(doc.querySelector('textarea')!.value).toBe('Hello');
  const draftCalls = () => calls.filter((c) => c.operation === 'draft:set' && c.args.draft);
  expect(draftCalls().at(-1)!.args.draft.actions).toEqual([first()]);
  expect(doc.querySelector<HTMLButtonElement>('[data-save]')!.disabled).toBe(true);
  click('重做');
  await flush();
  expect(doc.querySelector('textarea')!.value).toBe('Changed');
  expect(draftCalls().at(-1)!.args.draft.actions).toEqual([{ id: 'speech', type: 'speech', text: 'Changed' }]);
  click('撤销');
  await flush();
  click('在选中步骤后添加讲解');
  await flush();
  expect(doc.querySelector<HTMLButtonElement>('[data-teaching-history=redo]')!.disabled).toBe(true);
  expect(calls.some((c) => ['write', 'tts:submit', 'media:adopt'].includes(c.operation))).toBe(false);
});
it('undoes node add, reorder and delete without touching native history, and retains history after saving', async () => {
  await flush();
  click('在选中步骤后添加讲解');
  await flush();
  const delayId = doc.querySelectorAll<HTMLElement>('article')[1].dataset.actionId;
  doc.querySelectorAll('article')[1].querySelectorAll<HTMLButtonElement>('.teaching-node-controls button')[0].click();
  await flush();
  expect(doc.querySelector('article')!.dataset.actionId).toBe(delayId);
  click('撤销');
  await flush();
  expect(doc.querySelectorAll<HTMLElement>('article')[1].dataset.actionId).toBe(delayId);
  doc.querySelectorAll('article')[1].querySelectorAll<HTMLButtonElement>('.teaching-node-controls button')[2].click();
  await flush();
  expect(doc.querySelectorAll('article')).toHaveLength(1);
  click('撤销');
  await flush();
  expect(doc.querySelectorAll('article')).toHaveLength(2);
  click('保存本页教学');
  await flush();
  expect(persisted[1]).toMatchObject({ id: delayId, type: 'speech' });
  expect(doc.querySelector<HTMLButtonElement>('[data-teaching-history=undo]')!.disabled).toBe(false);
  expect(doc.querySelector<HTMLButtonElement>('[data-teaching-history=redo]')!.disabled).toBe(false);
});
it('does not carry draft undo across native page or source context changes', async () => {
  await flush();
  click('在选中步骤后添加讲解');
  await flush();
  state.context.page = 2;
  state.refresh();
  await flush();
  expect(doc.querySelector<HTMLButtonElement>('[data-teaching-history=undo]')!.disabled).toBe(true);
  expect(doc.querySelector('textarea')!.value).toBe('Page two');
});

it('bounds draft undo history to twenty edits', async () => {
  await flush();
  for (let i = 0; i < 22; i++) {
    click('在选中步骤后添加讲解');
    await flush();
  }
  expect(doc.querySelectorAll('article')).toHaveLength(23);
  for (let i = 0; i < 20; i++) {
    click('撤销');
    await flush();
  }
  expect(doc.querySelectorAll('article')).toHaveLength(3);
  expect(doc.querySelector<HTMLButtonElement>('[data-teaching-history=undo]')!.disabled).toBe(true);
  expect(doc.querySelector<HTMLButtonElement>('[data-teaching-history=redo]')!.disabled).toBe(false);
});

it('retains draft undo after a failed authoritative save', async () => {
  await flush();
  click('在选中步骤后添加讲解');
  await flush();
  writeFailure = true;
  click('保存本页教学');
  await flush();
  expect(doc.querySelector('[role=alert]')!.textContent).toContain('save conflict');
  expect(persisted).toEqual([first()]);
  click('撤销');
  await flush();
  expect(doc.querySelectorAll('article')).toHaveLength(1);
  expect(doc.querySelector('textarea')!.value).toBe('Hello');
});
it('preserves recovery content but clears undo after a native source epoch changes', async () => {
  await flush();
  click('在选中步骤后添加讲解');
  await flush();
  state.context.editGeneration++;
  state.refresh();
  await flush();
  expect(doc.querySelectorAll('article')).toHaveLength(2);
  expect(doc.querySelector<HTMLButtonElement>('[data-teaching-history=undo]')!.disabled).toBe(true);
});

it('cancels teaching media on native slideshow start and restores preview only after exit', async () => {
  await flush();
  const dispose = vi.fn(),
    start = vi.fn();
  previewState.load.mockResolvedValue({ createTeachingTimelinePreview: () => ({ start, dispose }) });
  click('画布预览');
  await flush();
  expect(start).toHaveBeenCalledOnce();
  slideshowState.active = true;
  slideshowState.callbacks.get('asc_onStartDemonstration')!();
  expect(dispose).toHaveBeenCalledOnce();
  expect(doc.querySelector<HTMLButtonElement>('[data-preview-command=play]')!.disabled).toBe(true);
  expect(doc.querySelector('[role=status]')!.textContent).toContain('暂不自动播放教学时间轴');
  slideshowState.active = false;
  slideshowState.callbacks.get('asc_onEndDemonstration')!();
  expect(doc.querySelector<HTMLButtonElement>('[data-preview-command=play]')!.disabled).toBe(false);
  expect(start).toHaveBeenCalledOnce();
  sidebar.dispose();
  expect(slideshowState.callbacks.size).toBe(0);
});

it('rejects a late speech preview response after entering and exiting native slideshow', async () => {
  await flush();
  deferPreview = true;
  click('试听语音');
  await flush();
  expect(delayedPreviewId).toBeTruthy();
  slideshowState.active = true;
  slideshowState.callbacks.get('asc_onStartDemonstration')!();
  slideshowState.active = false;
  slideshowState.callbacks.get('asc_onEndDemonstration')!();
  sidebar.receive(delayedPreviewId!, { result: { bytes: new Uint8Array([1]), mimeType: 'audio/mpeg' } });
  await flush();
  expect(doc.querySelector('audio')).toBeNull();
  expect(doc.querySelector('[role=status]')!.textContent).toContain('已退出原生放映');
});

const lessonSdk = () => {
  vi.useFakeTimers();
  sourceSha = 'a'.repeat(64);
  const callbacks = new Map<string, Set<() => void>>();
  const manager = {
    Mode: false,
    SlideNum: -1,
    DemonstrationDiv: null as HTMLElement | null,
    Transition: { IsPlaying: () => false },
    IsMainSeqFinished: () => true,
  };
  const notify = (event: string) => {
    for (const callback of callbacks.get(event) ?? []) callback();
  };
  const native = (slideshowState.api = {
    WordControl: {
      DemonstrationManager: manager,
      m_oLogicDocument: { Slides: [0, 1].map(() => ({ cSld: { spTree: [] } })) },
    },
    isSlideShow: () => manager.Mode,
    asc_registerCallback: (event: string, callback: () => void) => {
      if (!callbacks.has(event)) callbacks.set(event, new Set());
      callbacks.get(event)!.add(callback);
    },
    asc_unregisterCallback: (event: string, callback: () => void) => callbacks.get(event)?.delete(callback),
    goToPage: (page: number) => {
      if (state.context.page !== page + 1) {
        state.context.page = page + 1;
        state.context.editGeneration++;
      }
      queueMicrotask(state.refresh);
    },
    StartDemonstration: vi.fn((id: string, page: number) => {
      manager.Mode = true;
      manager.SlideNum = page;
      manager.DemonstrationDiv = doc.getElementById(id);
      notify('asc_onStartDemonstration');
      notify('asc_onDemonstrationSlideChanged');
    }),
    EndDemonstration: () => {
      manager.Mode = false;
      notify('asc_onEndDemonstration');
    },
    DemonstrationPause: () => {},
    DemonstrationPlay: () => {},
  });
  vi.stubGlobal(
    'URL',
    class extends URL {
      static createObjectURL = () => 'blob:lesson-media';
      static revokeObjectURL = vi.fn();
    },
  );
  const played: any[] = [];
  previewState.load.mockResolvedValue({
    createTeachingTimelinePreview: (input: any) => ({
      start: () => {
        played.push(input);
        input.onMode('playing');
      },
      dispose: vi.fn(),
    }),
  });
  state.context.editGeneration++;
  state.refresh();
  return { native, played };
};
it.each([1, 2])('plays every native page and restores starting page %s without writes', async (initialPage) => {
  const { native, played } = lessonSdk();
  await flush();
  if (initialPage === 2) {
    state.context.page = 2;
    state.context.editGeneration++;
    state.refresh();
    await flush();
  }
  click('放映整课教学');
  await vi.advanceTimersByTimeAsync(32);
  expect(
    played.map((input) => input.pageId),
    doc.querySelector('[data-status]')?.textContent + JSON.stringify(calls.map((c) => c.operation)),
  ).toEqual(['p1']);
  played[0].onComplete();
  await vi.advanceTimersByTimeAsync(32);
  expect(played.map((input) => input.pageId)).toEqual(['p1', 'p2']);
  expect(calls.some((call) => call.operation === 'tts:models' && call.context.page === 2)).toBe(true);
  expect(calls.some((call) => call.operation === 'media:list' && call.context.page === 2)).toBe(true);
  expect(calls.filter((call) => call.operation === 'tts:preview').map((call) => call.context.page)).toEqual([1, 2]);
  played[1].onComplete();
  await vi.advanceTimersByTimeAsync(32);
  expect(doc.querySelector('[data-teaching-lesson]')).toBeNull();
  expect(doc.querySelector('[data-status]')!.textContent).toBe('整课教学放映完成');
  expect(state.context.page).toBe(initialPage);
  expect(native.StartDemonstration).toHaveBeenCalledTimes(2);
  expect(doc.querySelector('textarea')!.disabled).toBe(false);
  expect([...doc.querySelectorAll('button')].find((node) => node.textContent === '绑定本页 PPT 动画')!.disabled).toBe(
    false,
  );
  expect(calls.some((call) => ['write', 'tts:submit', 'tts:prepare', 'media:adopt'].includes(call.operation))).toBe(
    false,
  );
});
it.each(['teaching', 'source'])(
  'stops rather than playing a later page from a different %s revision',
  async (change) => {
    const { played } = lessonSdk();
    await flush();
    click('放映整课教学');
    await vi.advanceTimersByTimeAsync(32);
    if (change === 'teaching') teachingRevision = 'changed';
    else sourceSha = 'b'.repeat(64);
    played[0].onComplete();
    await vi.advanceTimersByTimeAsync(32);
    expect(played).toHaveLength(1);
    expect(doc.querySelector('[data-teaching-lesson]')).toBeNull();
    expect(doc.querySelector('[role=alert]')!.textContent).toContain('版本已变化');
  },
);
it('does not resurrect lesson playback from a media response after cancelling preparation', async () => {
  const { native, played } = lessonSdk();
  await flush();
  deferPreview = true;
  click('放映整课教学');
  await vi.advanceTimersByTimeAsync(32);
  expect(delayedPreviewId).toBeTruthy();
  click('停止整课放映');
  sidebar.receive(delayedPreviewId!, { result: { bytes: new Uint8Array([1]), mimeType: 'audio/mpeg' } });
  await vi.advanceTimersByTimeAsync(32);
  expect(played).toHaveLength(0);
  expect(native.StartDemonstration).not.toHaveBeenCalled();
  expect(doc.querySelector('[data-teaching-lesson]')).toBeNull();
});
it('cancels the whole lesson on unexpected editor navigation while playing', async () => {
  const { played } = lessonSdk();
  await flush();
  click('放映整课教学');
  await vi.advanceTimersByTimeAsync(32);
  state.context.page = 2;
  state.context.editGeneration++;
  state.refresh();
  await flush();
  expect(doc.querySelector('[data-teaching-lesson]')).toBeNull();
  played[0].onComplete();
  await vi.advanceTimersByTimeAsync(32);
  expect(played).toHaveLength(1);
  expect(state.context.page).toBe(2);
});

it('preserves an unseen page recovery draft and stops rather than playing its older committed teaching', async () => {
  const { played } = lessonSdk();
  await flush();
  click('放映整课教学');
  await vi.advanceTimersByTimeAsync(32);
  recoveryOnSecond = true;
  played[0].onComplete();
  await vi.advanceTimersByTimeAsync(32);
  expect(played).toHaveLength(1);
  expect(doc.querySelector('[data-status]')!.textContent).toContain('未保存');
  expect(doc.querySelector('textarea')!.value).toBe('Recovered edit');
  expect(calls.some((call) => ['write', 'draft:set', 'tts:submit'].includes(call.operation))).toBe(false);
});

it('reports a failed return to editing after the lesson finishes instead of hiding it', async () => {
  const { native, played } = lessonSdk();
  await flush();
  click('放映整课教学');
  await vi.advanceTimersByTimeAsync(32);
  played[0].onComplete();
  await vi.advanceTimersByTimeAsync(32);
  native.goToPage = () => {
    throw new Error('restoration failed');
  };
  played[1].onComplete();
  await vi.advanceTimersByTimeAsync(32);
  expect(doc.querySelector('[data-teaching-lesson]')).toBeNull();
  expect(doc.querySelector('[role=alert]')!.textContent).toContain('restoration failed');
});

it('binds all saved native steps, persists their exact source and retains native order during editing', async () => {
  const { native } = lessonSdk();
  const slide = native.WordControl.m_oLogicDocument.Slides[0] as any;
  slide.cSld.spTree = [{ Id: 'shape' }];
  const effects = [0, 1].map(() => ({
    cTn: { presetClass: 1, presetID: 10, presetSubtype: 0, nodeType: 2 },
    getObjectId: () => 'shape',
    asc_getDuration: () => 500,
    asc_getDelay: () => 0,
    isIdle: () => true,
    isAtEnd: () => false,
  }));
  slide.timing = { getAllAnimEffects: () => effects, getEffectsSequences: () => [[null, ...effects]] };
  Object.assign(native.WordControl.DemonstrationManager, { OnNextSlideAnimPlayer: vi.fn(), StartAnimation: vi.fn() });
  const lineage = vi
    .spyOn(deckwisePageLineage, 'context')
    .mockReturnValue({ sourceSha256: sourceSha!, pageTokens: ['1', '2'], openedPageTokens: ['1', '2'] });
  try {
    await flush();
    click('绑定本页 PPT 动画');
    await flush();
    expect(doc.querySelectorAll('article')).toHaveLength(3);
    const nodes = () =>
      calls.filter((call) => call.operation === 'draft:set' && call.args.draft).at(-1)!.args.draft.actions;
    expect(nodes().slice(1)).toEqual(
      [0, 1].map((stepIndex) =>
        expect.objectContaining({ type: 'native_animation', version: 1, sourceSha256: sourceSha, stepIndex }),
      ),
    );
    const card = doc.querySelectorAll('article')[2];
    [...card.querySelectorAll('button')].find((button) => button.textContent === '上移')!.click();
    await flush();
    expect(doc.querySelector('[role=alert]')!.textContent).toContain('原始顺序');
    expect(nodes().map((node: any) => node.stepIndex)).toEqual([undefined, 0, 1]);
    click('保存本页教学');
    await flush();
    expect(persisted.slice(1).map((node) => node.stepIndex)).toEqual([0, 1]);
    expect(persisted[1].sourceSha256).toBe(sourceSha);
    const controller = { resetSelection: vi.fn(), selectObject: vi.fn() };
    Object.assign(native.WordControl.m_oLogicDocument, { GetCurrentController: () => controller });
    const tab = doc.createElement('button');
    tab.id = 'animate';
    const openAnimation = vi.fn();
    tab.addEventListener('click', openAnimation);
    doc.body.append(tab);
    doc.querySelectorAll<HTMLButtonElement>('.teaching-step')[1].click();
    click('编辑原生动画');
    expect(controller.resetSelection).toHaveBeenCalledOnce();
    expect(controller.selectObject).toHaveBeenCalledWith(slide.cSld.spTree[0], 0);
    expect(openAnimation).toHaveBeenCalledOnce();
    expect(doc.querySelector<HTMLElement>('.deckwise-teaching-sidebar')!.hidden).toBe(true);
    state.toggle();
    const previewLoads = previewState.load.mock.calls.length;
    click('画布预览');
    await flush();
    expect(doc.querySelector('[role=alert]')!.textContent).toContain('幻灯片放映');
    expect(previewState.load).toHaveBeenCalledTimes(previewLoads);
    click('移除动画绑定');
    await flush();
    click('保存本页教学');
    await flush();
    expect(persisted).toEqual([first()]);
    lineage.mockReturnValue(undefined);
    click('绑定本页 PPT 动画');
    await flush();
    expect(doc.querySelector('[role=alert]')!.textContent).toContain('保存并重新打开');
    expect(doc.querySelectorAll('article')).toHaveLength(1);
  } finally {
    lineage.mockRestore();
  }
});

it('refreshes selection bindings on the same saved page without replacing the timeline or rereading media catalogs', async () => {
  sourceSha = 'a'.repeat(64);
  await flush();
  click('刷新');
  await flush();
  const text = doc.querySelector('textarea')!;
  text.value = 'Draft survives selection';
  text.dispatchEvent(new Event('input'));
  await flush();
  const catalogs = calls.filter((call) => ['tts:models', 'tts:list', 'media:list'].includes(call.operation)).length;
  state.context = { ...state.context, editGeneration: state.context.editGeneration + 1 };
  state.refresh();
  await flush();
  expect(doc.querySelector('textarea')).toBe(text);
  expect(text.value).toBe('Draft survives selection');
  expect(calls.filter((call) => ['tts:models', 'tts:list', 'media:list'].includes(call.operation))).toHaveLength(
    catalogs,
  );
});

it('automatically commits teaching edits after idle without losing the narration cursor', async () => {
  await flush();
  vi.useFakeTimers();
  const input = doc.querySelector('textarea')!;
  input.focus();
  input.value = 'Updated narration';
  input.setSelectionRange(4, 4);
  input.dispatchEvent(new Event('input'));
  await flush();
  await vi.advanceTimersByTimeAsync(1200);
  await flush();
  expect(persisted[0].text).toBe('Updated narration');
  expect(doc.activeElement).toBe(doc.querySelector('textarea'));
  expect(doc.querySelector('textarea')!.selectionStart).toBe(4);
  expect(doc.querySelector('[role=status]')!.textContent).toContain('已保存');
});
it('commits a recovered draft before canvas preview without requiring a manual save', async () => {
  await flush();
  setAfter('pause');
  await flush();
  state.context.editGeneration++;
  state.refresh();
  await flush();
  const start = vi.fn();
  previewState.load.mockResolvedValue({ createTeachingTimelinePreview: () => ({ start, dispose: vi.fn() }) });
  click('画布预览');
  await flush();
  expect(persisted).toHaveLength(1);
  expect(start).toHaveBeenCalledOnce();
});
it('keeps the draft and reports an automatic save failure instead of starting preview', async () => {
  await flush();
  setAfter('pause');
  await flush();
  writeFailure = true;
  previewState.load.mockClear();
  click('画布预览');
  await flush();
  expect(doc.querySelector('[role=alert]')!.textContent).toContain('save conflict');
  expect(doc.querySelectorAll('article')).toHaveLength(1);
  expect(persisted).toHaveLength(1);
  expect(previewState.load).not.toHaveBeenCalled();
});

it('keeps speech controls in the current sidebar after saving before audition', async () => {
  await flush();
  setAfter('pause');
  await flush();
  const media = doc.defaultView!.HTMLMediaElement.prototype;
  const play = vi.spyOn(media, 'play').mockResolvedValue();
  const pause = vi.spyOn(media, 'pause').mockImplementation(() => {});
  const load = vi.spyOn(media, 'load').mockImplementation(() => {});
  vi.stubGlobal('URL', { createObjectURL: () => 'blob:test', revokeObjectURL: vi.fn() });
  try {
    click('试听语音');
    await flush();
    expect(persisted).toHaveLength(1);
    expect(doc.querySelector('[data-teaching-inspector] audio')).toBeTruthy();
    expect(play).toHaveBeenCalledOnce();
    sidebar.dispose();
  } finally {
    play.mockRestore();
    pause.mockRestore();
    load.mockRestore();
  }
});
