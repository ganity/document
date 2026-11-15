import { createHash } from 'node:crypto';
import { readFile, writeFile } from 'node:fs/promises';
import { test, expect } from './lib/l0';
import { toBase64 } from './lib/ooxml';

test('Main lesson planning and restarted audio recovery save and reopen a real native PPT', async ({ page }, info) => {
  test.setTimeout(240000);
  const fixture = process.env.DECKWISE_NATIVE_GENERATED_FIXTURE;
  const moduleUrl = process.env.DECKWISE_ENRICHMENT_RECOVERY_URL;
  test.skip(!fixture || !moduleUrl, 'Requires a native lesson and the current Desktop Main integration bundle');
  const { CoursewareNativeTools, CoursewareNativeAudio, CoursewareCandidateCoordinator, adoptNativeMaterials } = await import(moduleUrl!);
  const binding = { projectId: 'enrichment-recovery', documentId: 'lesson', sessionId: 'isolated' };
  await page.goto('/embed-demo.html');
  await page.evaluate(() => {
    document.querySelector('#status')!.textContent = 'loading';
    document.querySelector('#log')!.setAttribute('style', 'display:none');
    (document.querySelector('#editorFrame') as HTMLIFrameElement).src = '/editor?embed=1&deckwise=1&locale=zh-CN';
  });
  await expect(page.locator('#status')).toHaveText('ready', { timeout: 60000 });
  await page.evaluate(async ({ data, binding }) => post('document:open-buffer', {
    fileName: 'native.pptx', buffer: Uint8Array.from(atob(data), c => c.charCodeAt(0)).buffer,
    editorBinding: binding, trackPageLineage: true,
  }), { data: toBase64(await readFile(fixture + '/native.pptx')), binding });
  await page.waitForFunction(() => !!window.__ooFrames.readyEditor(), { timeout: 60000 });
  const teaching = { native: async (operation: string, args: any = {}, n = 1) => {
    const bytes = args.command?.bytes ? Array.from(new Uint8Array(args.command.bytes)) : undefined;
    return page.evaluate(async ({ operation, args, n, bytes, binding }) => {
      if (bytes) args.command.bytes = new Uint8Array(bytes).buffer;
      const editor = await post('document:get-context', binding);
      const result = (await post('document:native-teaching', { binding, operation: 'route', routeOperation: operation, args, page: n })).routed;
      return operation === 'narration-context' ? { value: { ...result.value, editor } } : result;
    }, { operation, args: bytes ? { ...args, command: { ...args.command, bytes: undefined } } : args, n, bytes, binding });
  } };
  const before = (await teaching.native('open')).value;
  expect(before.project.pages.length).toBeGreaterThanOrEqual(2);
  const script = before.project.pages.map((_: unknown, i: number) => `P${i + 1}: Exact teacher script ${i + 1}.`).join('\n');
  await adoptNativeMaterials({ teaching, view: before, materials: { apply: true, script: { text: script, policy: 'preserve' } }, check() {} });
  const captures = await Promise.all([1, 2].map(n => teaching.native('narration-context', {}, n)));
  const quotes = captures.map(c => c.value.text.trim().slice(0, 80));
  expect(quotes.every(q => q.length > 0)).toBe(true);
  const lessonText = `Teach: ${quotes[0]}\nReview: ${quotes[1]}\nA circle is round.`;
  let modelCalls = 0, saveCount = 0, savedBytes: number[] = [];
  const candidates = new CoursewareCandidateCoordinator(teaching, async () => {
    modelCalls++;
    return { modelKey: 'local/test', text: JSON.stringify({ units: [
      { title: 'Teach', evidence: `Teach: ${quotes[0]}`, kind: 'teaching', mappings: [{ pageId: before.project.pages[0].pageId, evidence: quotes[0] }] },
      { title: 'Review', evidence: `Review: ${quotes[1]}`, kind: 'teaching', mappings: [{ pageId: before.project.pages[1].pageId, evidence: quotes[1] }] },
      { title: 'Quiz', evidence: 'A circle is round.', kind: 'quiz', mappings: [] },
    ] }) };
  }, undefined, async () => ({ modelKey: 'local/test' }));
  const save = async () => {
    savedBytes = await page.evaluate(async () => Array.from(new Uint8Array(await (await post('document:save')).file.arrayBuffer())));
    await writeFile(info.outputPath('recovered-native.pptx'), new Uint8Array(savedBytes)); saveCount++;
  };
  const toolsFor = (audio: any) => new CoursewareNativeTools(teaching, candidates, audio, teaching.native, () => 'chat', undefined, undefined, undefined, save);
  let tools = toolsFor({});
  const sourceSha256 = (await teaching.native('open')).value.project.source.sha256;
  const plan = await tools.execute('chat', 'courseware_narration_plan', { sourceSha256, apply: true, lessonPlan: { text: lessonText } }, () => true);
  expect(plan.value.saved).toBe(true); expect(modelCalls).toBe(1);
  let view = (await teaching.native('open')).value;
  const preservedScript = structuredClone(view.enrichment.script);
  const inserted = await tools.execute('chat', 'courseware_activity_insert', { sourceSha256: view.project.source.sha256, apply: true,
    unitId: 'lesson-unit-3', afterPageId: before.project.pages[0].pageId, question: 'Which shape is round?',
    referenceEvidence: 'A circle is round.', response: { kind: 'single', options: ['circle', 'square'], answerIndices: [0] } }, () => true);
  expect(inserted.value.saved).toBe(true);

  // Only the remote service is simulated: real Main journals, validation, native adoption and saves run below.
  const records = new Map<string, any>(), tasks = new Map<string, any>(); let submits = 0, quoteCount = 0;
  const wave = Buffer.alloc(16044);
  wave.write('RIFF', 0); wave.writeUInt32LE(16036, 4); wave.write('WAVEfmt ', 8); wave.writeUInt32LE(16, 16);
  wave.writeUInt16LE(1, 20); wave.writeUInt16LE(1, 22); wave.writeUInt32LE(8000, 24); wave.writeUInt32LE(16000, 28);
  wave.writeUInt16LE(2, 32); wave.writeUInt16LE(16, 34); wave.write('data', 36); wave.writeUInt32LE(16000, 40);
  for (let i = 0; i < 8000; i++) wave.writeInt16LE(Math.round(Math.sin(i * Math.PI * 2 * 440 / 8000) * 1000), 44 + i * 2);
  const sha256 = createHash('sha256').update(wave).digest('hex');
  const host = { async call(method: string, args: any) {
    if (method === 'courseware.mediaRelay.scope') return { sessionId: 'chat', teamProjectId: null };
    if (method === 'courseware.mediaRelay.read') return structuredClone(records.get(args.key) ?? null);
    expect(method).toBe('courseware.mediaRelay.write');
    expect(args.expectedRevision).toBe(records.get(args.key)?.revision ?? 0);
    const value = { revision: args.expectedRevision + 1, value: structuredClone(args.value) }; records.set(args.key, value); return value;
  } };
  const auth = { getCoursewarePrincipal: () => ({ key: 'a'.repeat(64), assertCurrent() {} }), coursewareMedia: {
    getTtsQuote: async () => { quoteCount++; return { model: 'test-voice', maxChars: 4000, maxPoints: 1, expiresAt: new Date(Date.now() + 60000).toISOString() }; },
    submitConfirmedTts: async (key: string, request: any) => {
      submits++;
      const canonical = [request.kind, request.model, request.input, request.voice ?? null].reduce((s, f) => `${s}:${f === null ? -1 : f.length}:${f ?? ''}`, 'courseware-media-v1');
      const task = { model: request.model, taskId: `task-${submits}`, idempotencyKey: key, requestHash: createHash('sha256').update(canonical).digest('hex'),
        state: submits === 1 ? 'ready' : 'running', feeState: 'settled', reservedPoints: 1, actualPoints: 1, errorCode: null,
        assetRef: { mime: 'audio/wav', sha256, byteLength: wave.length } };
      tasks.set(task.taskId, task); return task;
    },
    get: async (ref: any) => tasks.get(ref.taskId),
    download: async (ref: any) => ({ task: tasks.get(ref.taskId), bytes: wave }),
  } };
  const newAudio = () => new CoursewareNativeAudio(auth, () => host, () => 'chat', teaching.native, async () => 1000);
  const originalIds = before.project.pages.slice(0, 2).map((p: any) => p.pageId);
  for (const pageId of originalIds) {
    view = (await teaching.native('open')).value;
    const n = view.project.pages.findIndex((p: any) => p.pageId === pageId) + 1;
    const current = (await teaching.native('open', {}, n)).value;
    const audio = newAudio(), scope = { viewId: current.viewId, expectedTeachingRevision: current.project.teachingRevision };
    const intent = await audio.execute('tts:prepare', { ...scope, actionId: current.project.pages[n - 1].actions.find((a: any) => a.type === 'speech').id,
      model: 'test-voice', voice: null }, n, current);
    const quote = await audio.execute('tts:quote', { ...scope, intentId: intent.intentId }, n, current);
    await audio.execute('tts:submit', { ...scope, intentId: intent.intentId, quoteId: quote.quoteId }, n, current);
  }
  tools = toolsFor(newAudio()); // Recreate Main's coordinator from the durable journal.
  view = (await teaching.native('open')).value;
  const recovery = { kind: 'tts', recover: true, sourceSha256: view.project.source.sha256,
    pageNumbers: originalIds.map((id: string) => view.project.pages.findIndex((p: any) => p.pageId === id) + 1) };
  const partial = (await tools.execute('chat', 'courseware_media_refresh', recovery, () => true)).value;
  expect(partial.saved).toBe(true); expect(partial.audioReady).toBe(false);
  expect(partial.pages.map((p: any) => p.state)).toEqual(['adopted', 'running']);
  tasks.get('task-2').state = 'ready';
  view = (await teaching.native('open')).value;
  const complete = (await tools.execute('chat', 'courseware_media_refresh', { ...recovery, sourceSha256: view.project.source.sha256 }, () => true)).value;
  expect(complete.audioReady).toBe(true); expect(complete.pages.map((p: any) => p.state)).toEqual(['retained', 'adopted']);
  expect(submits).toBe(2); expect(quoteCount).toBe(2); expect(saveCount).toBe(4);
  view = (await teaching.native('open')).value;
  expect(view.enrichment.script).toEqual(preservedScript);
  await page.evaluate(async ({ savedBytes, binding }) => post('document:open-buffer', {
    fileName: 'recovered-native.pptx', buffer: new Uint8Array(savedBytes).buffer, editorBinding: binding, trackPageLineage: true,
  }), { savedBytes, binding });
  await page.waitForFunction(() => !!window.__ooFrames.readyEditor(), { timeout: 60000 });
  const reopened = (await teaching.native('open')).value;
  expect(reopened.enrichment).toEqual(view.enrichment);
  for (const id of originalIds) expect(reopened.project.pages.find((p: any) => p.pageId === id).actions.find((a: any) => a.type === 'speech').audioAssetId).toBe(`tts-${sha256}`);
  expect((await toolsFor({}).execute('chat', 'courseware_narration_plan', { sourceSha256: reopened.project.source.sha256,
    apply: true, lessonPlan: { text: lessonText } }, () => true)).value.reused).toBe(true);
  expect(modelCalls).toBe(1);
});
