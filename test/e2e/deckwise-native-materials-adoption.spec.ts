import { readFile, writeFile } from 'node:fs/promises';
import { test, expect } from './lib/l0';
import { toBase64 } from './lib/ooxml';

for (const mode of ['script', 'plan', 'both'] as const) test(`real native ${mode} adoption uses Main services and survives reopen`, async ({ page }, testInfo) => {
  test.setTimeout(240000);
  const fixture = process.env.DECKWISE_NATIVE_GENERATED_FIXTURE;
  const moduleUrl = process.env.DECKWISE_MATERIALS_URL;
  test.skip(!fixture || !moduleUrl, 'Requires copied native lesson and current Main materials module');
  const { adoptNativeMaterials, insertNativeActivity } = await import(moduleUrl!);
  const binding = { projectId: `adoption-${mode}`, documentId: 'lesson', sessionId: 'isolated' };
  await page.goto('/embed-demo.html');
  await page.evaluate(() => {
    document.querySelector('#status')!.textContent = 'loading';
    document.querySelector('#log')!.setAttribute('style', 'display:none');
    (document.querySelector('#editorFrame') as HTMLIFrameElement).src = '/editor?embed=1&deckwise=1&locale=zh-CN';
  });
  await expect(page.locator('#status')).toHaveText('ready', { timeout: 60000 });
  await page.evaluate(async ({ data, binding }) => post('document:open-buffer', {
    fileName: 'adoption.pptx', buffer: Uint8Array.from(atob(data), c => c.charCodeAt(0)).buffer,
    editorBinding: binding, trackPageLineage: true,
  }), { data: toBase64(await readFile(fixture + '/native.pptx')), binding });
  await page.waitForFunction(() => !!window.__ooFrames.readyEditor(), { timeout: 60000 });
  const teaching = { native: async (operation: string, args = {}, n = 1) => page.evaluate(async ({ operation, args, n, binding }) => {
    const editor = await post('document:get-context', binding);
    const result = (await post('document:native-teaching', { binding, operation: 'route', routeOperation: operation, args, page: n })).routed;
    return operation === 'narration-context' ? { value: { ...result.value, editor } } : result;
  }, { operation, args, n, binding }) };
  const before = (await teaching.native('open')).value;
  const originalPages = before.project.pages.map((p: any) => p.pageId);
  const text = before.project.pages.map((_: unknown, i: number) => `P${i + 1}: Exact supplied words ${i + 1}. 1.25 kg. Watch the video!`).join('\n');
  const materials = { apply: false, ...(mode !== 'plan' ? { script: { text, policy: 'preserve' } } : {}),
    ...(mode !== 'script' ? { lessonPlan: { text: 'Review: A circle is round.', units: [{ id: 'round', title: 'Round shapes',
      evidence: 'A circle is round.', pageIds: [], kind: 'quiz' }] } } : {}) };
  const preview = await adoptNativeMaterials({ teaching, view: before, materials, check() {} });
  expect(preview.preview).toBe(true);
  expect((await teaching.native('open')).value.project.teachingRevision).toBe(before.project.teachingRevision);
  const adopted = await adoptNativeMaterials({ teaching, view: before, materials: { ...materials, apply: true }, check() {} });
  expect(adopted.failures).toEqual([]);
  let view = (await teaching.native('open')).value;
  expect(view.enrichment.originalPages.map((p: any) => p.pageId)).toEqual(originalPages);
  if (mode !== 'script') {
    const args = { apply: true, unitId: 'round', afterPageId: originalPages[0], question: 'Which shape is round?',
      referenceEvidence: 'A circle is round.', response: { kind: 'single', options: ['circle', 'square'], answerIndices: [0] } };
    const inserted = await insertNativeActivity({ teaching, view, args, check() {} });
    expect(inserted.applied).toBe(true);
    view = (await teaching.native('open')).value;
    expect(view.project.pages[1].pageId).toBe(inserted.pageId);
    expect((await insertNativeActivity({ teaching, view, args, check() {} })).reused).toBe(true);
    expect((await teaching.native('open')).value.project.pages.length).toBe(originalPages.length + 1);
  }
  // Deterministic local media exercises verified adoption; no provider is called.
  await page.evaluate(async ({ binding, video }) => {
    const wave = new ArrayBuffer(16044), data = new DataView(wave);
    const ascii = (at: number, s: string) => [...s].forEach((c, i) => data.setUint8(at + i, c.charCodeAt(0)));
    ascii(0, 'RIFF'); data.setUint32(4, 16036, true); ascii(8, 'WAVE'); ascii(12, 'fmt '); data.setUint32(16, 16, true);
    data.setUint16(20, 1, true); data.setUint16(22, 1, true); data.setUint32(24, 8000, true); data.setUint32(28, 16000, true);
    data.setUint16(32, 2, true); data.setUint16(34, 16, true); ascii(36, 'data'); data.setUint32(40, 16000, true);
    for (let i = 0; i < 8000; i++) data.setInt16(44 + i * 2, Math.sin(i * Math.PI * 2 * 440 / 8000) * 1000, true);
    const read = async () => (await post('document:native-teaching', { binding, operation: 'read' })).teaching;
    const initial = await read();
    for (const page of initial.pages) for (const action of page.actions) {
      if (!['speech', 'music_slot', 'video_slot'].includes(action.type) || action.audioAssetId) continue;
      const current = await read(), isVideo = action.type === 'video_slot';
      const bytes = isVideo ? Uint8Array.from(atob(video), c => c.charCodeAt(0)).buffer : wave;
      const sha256 = Array.from(new Uint8Array(await crypto.subtle.digest('SHA-256', bytes)), b => b.toString(16).padStart(2, '0')).join('');
      await post('document:native-teaching', { binding, operation: 'adopt-audio', command: {
        documentId: current.documentId, expectedRevision: current.revision, pageId: page.pageId, actionId: action.id,
        assetId: `local-${page.pageId}-${action.id}`, sha256, mimeType: isVideo ? 'video/mp4' : 'audio/wav', bytes, durationMs: 1000,
        ...(action.type === 'speech' ? {} : { kind: isVideo ? 'video' : 'music' }),
      } });
    }
  }, { binding, video: toBase64(await readFile(fixture + '/video.mp4')) });
  view = (await teaching.native('open')).value;
  const bytes = await page.evaluate(async () => Array.from(new Uint8Array(await (await post('document:save')).file.arrayBuffer())));
  await writeFile(testInfo.outputPath('native.pptx'), new Uint8Array(bytes));
  await page.evaluate(async ({ bytes, binding }) => post('document:open-buffer', {
    fileName: 'adoption-reopened.pptx', buffer: new Uint8Array(bytes).buffer, editorBinding: binding, trackPageLineage: true,
  }), { bytes, binding });
  await page.waitForFunction(() => !!window.__ooFrames.readyEditor(), { timeout: 60000 });
  const reopened = (await teaching.native('open')).value;
  expect(reopened.enrichment).toEqual(view.enrichment);
  for (const segment of reopened.enrichment.script?.segments ?? []) expect(reopened.project.pages.find((p: any) => p.pageId === segment.pageId)
    .actions.find((a: any) => a.id === segment.actionId).text).toBe(segment.text);
  if (mode !== 'script') {
    const assessment = (await teaching.native('assessment-context', {}, 2)).value;
    expect(assessment.classroom.scenes.find((s: any) => s.content.type === 'quiz' && s.content.questions[0].question === 'Which shape is round?')
      .content.questions[0].answer).toEqual(['A']);
  }
});
