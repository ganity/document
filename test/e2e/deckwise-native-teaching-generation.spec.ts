import { readFile, writeFile } from 'node:fs/promises';
import { test, expect } from './lib/l0';
import { toBase64 } from './lib/ooxml';

test('generated lesson closes the native edit, media, save, reopen and playback loop', async ({ page }, info) => {
  test.setTimeout(240000);
  const fixture = process.env.DECKWISE_NATIVE_GENERATED_FIXTURE;
  test.skip(!fixture, 'Requires a lesson produced by the real Desktop generation/Host test');
  const original = await readFile(fixture + '/native.pptx');
  const initial = JSON.parse(await readFile(fixture + '/teaching.json', 'utf8'));
  const video = toBase64(await readFile(fixture + '/video.mp4'));
  const binding = { projectId: 'real-fixture', documentId: 'family', sessionId: 'isolated' };
  if (process.env.DECKWISE_NATIVE_MODULE) await page.addInitScript(() => {
    if (window.parent === window) return;
    const listeners = new Set<(command: unknown) => void>();
    window.addEventListener('message', event => {
      if (event.source === window.parent && event.origin === location.origin) listeners.forEach(listener => listener(event.data));
    });
    window.deckwiseEditorHost = {
      onCommand(listener) { listeners.add(listener); return () => { listeners.delete(listener); }; },
      sendEvent(event) {
        if (event.type === 'message') window.parent.postMessage(event.message, location.origin);
        else if (event.type === 'saved') window.parent.postMessage({ type: 'document:saved', id: event.id,
          payload: { file: new Blob([event.bytes!]), fileName: event.fileName, pageLineage: event.pageLineage } }, location.origin);
        else if (event.type === 'save-requested') window.parent.postMessage({ type: 'document:save-requested', payload: {} }, location.origin);
      },
    };
  });
  await page.goto('/embed-demo.html');
  // The demo prints full native payloads; its growing text layout is outside the editor's navigation budget.
  await page.locator('#log').evaluate(node => { node.style.display = 'none'; });
  await page.evaluate(() => {
    document.querySelector('#status')!.textContent = 'loading';
    (window as any).nativeHostCalls = [];
    window.addEventListener('message', event => {
      if (event.data?.type !== 'document:teaching-request') return;
      const operation = event.data.payload.operation;
      (window as any).nativeHostCalls.push(operation);
      if (!['tts:models', 'tts:list'].includes(operation)) throw new Error('Unexpected legacy teaching request: ' + operation);
      (document.querySelector('#editorFrame') as HTMLIFrameElement).contentWindow!.postMessage({
        type: 'document:teaching-response', id: event.data.id, payload: { result: [] },
      }, location.origin);
    });
    (document.querySelector('#editorFrame') as HTMLIFrameElement).src = '/editor?embed=1&deckwise=1&locale=zh-CN';
  });
  await expect(page.locator('#status')).toHaveText('ready', { timeout: 60000 });
  await page.evaluate(
    async ({ data, binding }) => {
      await post('document:open-buffer', {
        fileName: 'generated-native.pptx',
        buffer: Uint8Array.from(atob(data), (c) => c.charCodeAt(0)).buffer,
        editorBinding: binding,
        trackPageLineage: true,
      });
    },
    { data: toBase64(original), binding },
  );
  await page.waitForFunction(() => !!window.__ooFrames.readyEditor(), { timeout: 60000 });
  const native = page.frames().find(f => f.name() === 'frameEditor')!;
  const before = await page.evaluate(async binding => (await post('document:native-teaching', { binding, operation: 'read' })).teaching, binding);
  expect(before.documentId).toBe(initial.documentId);
  expect(before.pages).toEqual(initial.pages);
  const generatedView = await page.evaluate(async binding => (await post('document:native-teaching', { binding, operation: 'route', routeOperation: 'open', page: 1 })).routed.value, binding);
  expect(generatedView.teaching.objective).toBe('Say Hello');
  expect(generatedView.teachingRequirements).toHaveLength(3);
  expect(before.pages[1].actions[1].response.answerIndices).toEqual([1]);
  // A teacher edits generated narration using the same native history as slide edits.
  await page.evaluate(async binding => {
    const { teaching } = await post('document:native-teaching', { binding, operation: 'read' });
    const actions = structuredClone(teaching.pages[0].actions);
    actions[0].text += ' Edited in Office.';
    await post('document:native-teaching', { binding, operation: 'replace-page', command: {
      documentId: teaching.documentId, expectedRevision: teaching.revision, pageId: teaching.pages[0].pageId, actions,
    } });
  }, binding);
  await native.evaluate(() => (Reflect.get(window, 'editor') as unknown as { Undo(): void }).Undo());
  expect(await page.evaluate(async binding => (await post('document:native-teaching', { binding, operation: 'read' })).teaching.pages[0].actions[0].text, binding)).toBe(initial.pages[0].actions[0].text);
  await native.evaluate(() => (Reflect.get(window, 'editor') as unknown as { Redo(): void }).Redo());
  const mediaResult = await page.evaluate(async ({ binding, video }) => {
    // Local deterministic PCM stands in for the authorized provider response; the real adoption path verifies bytes.
    const sampleRate = 8000, count = sampleRate, wave = new ArrayBuffer(44 + count * 2), view = new DataView(wave);
    const ascii = (at: number, value: string) => [...value].forEach((c, i) => view.setUint8(at + i, c.charCodeAt(0)));
    ascii(0, 'RIFF'); view.setUint32(4, wave.byteLength - 8, true); ascii(8, 'WAVE'); ascii(12, 'fmt '); view.setUint32(16, 16, true);
    view.setUint16(20, 1, true); view.setUint16(22, 1, true); view.setUint32(24, sampleRate, true); view.setUint32(28, sampleRate * 2, true);
    view.setUint16(32, 2, true); view.setUint16(34, 16, true); ascii(36, 'data'); view.setUint32(40, count * 2, true);
    for (let i = 0; i < count; i++) view.setInt16(44 + i * 2, Math.sin(i * Math.PI * 2 * 440 / sampleRate) * 1000, true);
    const videoBytes = Uint8Array.from(atob(video), c => c.charCodeAt(0)).buffer;
    const read = async () => (await post('document:native-teaching', { binding, operation: 'read' })).teaching;
    let teaching = await read(), adopted = 0;
    const initialView = (await post('document:native-teaching', { binding, operation: 'route', routeOperation: 'open', page: 1 })).routed.value;
    const generatedTarget = initialView.teachingObjects.find((o: any) => o.effectTargetEligible).objectId;
    const generatedActions = structuredClone(teaching.pages[0].actions);
    generatedActions[0].attention = { type: 'highlight', objectIds: [generatedTarget], animation: 'pulse' };
    await post('document:native-teaching', { binding, operation: 'replace-page', command: {
      documentId: teaching.documentId, expectedRevision: teaching.revision, pageId: teaching.pages[0].pageId, actions: generatedActions,
    } });
    teaching = await read();
    for (let i = 0; i < teaching.pages.length; i++) {
      for (const action of teaching.pages[i].actions) {
        if (!['speech', 'music_slot', 'video_slot'].includes(action.type)) continue;
        teaching = await read();
        const bytes = action.type === 'video_slot' ? videoBytes : wave;
        const sha256 = Array.from(new Uint8Array(await crypto.subtle.digest('SHA-256', bytes)), b => b.toString(16).padStart(2, '0')).join('');
        await post('document:native-teaching', { binding, operation: 'adopt-audio', command: {
          documentId: teaching.documentId, expectedRevision: teaching.revision, pageId: teaching.pages[i].pageId,
          actionId: action.id, assetId: 'generated-media-' + adopted++, sha256, mimeType: action.type === 'video_slot' ? 'video/mp4' : 'audio/wav',
          bytes, durationMs: 1000, ...(action.type === 'speech' ? {} : { kind: action.type === 'video_slot' ? 'video' : 'music' }),
        } });
      }
    }
    const result = await post('document:native-teaching', { binding, operation: 'route', routeOperation: 'open', page: 1 });
    const current = result.routed.value, objectId = current.teachingObjects.find((o: any) => o.effectTargetEligible).objectId;
    const actions = structuredClone(current.project.pages[0].actions);
    if (actions[0].attention !== undefined || actions[0].visualCues?.length !== 2) throw new Error('Generated attention was not resolved by real media adoption');
    actions[0].visualCues = [
      ...actions[0].visualCues,
      { id: 'generated-spotlight', atMs: 0, endMs: 1000, effect: { type: 'spotlight', elementId: objectId, dimOpacity: 0.6 } },
    ];
    await post('document:native-teaching', { binding, operation: 'route', routeOperation: 'write', page: 1,
      args: { viewId: current.viewId, expectedTeachingRevision: current.project.teachingRevision, actions } });
    return { teaching: await read(), objectId, adopted };
  }, { binding, video });
  expect(mediaResult.adopted).toBe(5);
  await page.evaluate(async binding => { (document.querySelector('#editorFrame') as HTMLIFrameElement).contentWindow!.postMessage({ type: 'document:teaching-enable', payload: { binding } }, location.origin); }, binding);
  const office = page.frameLocator('iframe#editorFrame').frameLocator('iframe[name="frameEditor"]');
  const panel = office.getByRole('complementary', { name: '教学时间轴' });
  await expect(panel).toBeVisible({ timeout: 15000 });
  await panel.getByRole('button', { name: '试听语音', exact: true }).first().click();
  await expect(office.locator('[data-attention-kind="highlight"]')).toBeVisible();
  await expect(office.locator('[data-attention-kind="spotlight"]')).toBeVisible();
  await expect.poll(() => panel.locator('audio').evaluate((audio: HTMLAudioElement) => audio.currentTime)).toBeGreaterThan(0);
  await panel.locator('audio').evaluate((audio: HTMLAudioElement) => audio.pause());
  const exported = await page.evaluate(async () => Array.from(new Uint8Array(await (await post('document:save')).file.arrayBuffer())));
  const output = info.outputPath('generated-native-edited.pptx');
  await writeFile(output, new Uint8Array(exported));
  if (fixture) await writeFile(fixture + '/edited.pptx', new Uint8Array(exported));
  await info.attach('generated-native-edited', { path: output, contentType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation' });
  await page.evaluate(async ({ bytes, binding }) => {
    await post('document:open-buffer', { fileName: 'reopened-generated.pptx', buffer: new Uint8Array(bytes).buffer, editorBinding: binding, trackPageLineage: true });
  }, { bytes: exported, binding });
  await page.waitForFunction(() => !!window.__ooFrames.readyEditor(), { timeout: 60000 });
  const reopened = await page.evaluate(async binding => {
    const { teaching } = await post('document:native-teaching', { binding, operation: 'read' });
    let audioCount = 0;
    for (const asset of teaching.media.filter((a: any) => a.mimeType.startsWith('audio/'))) {
      const { audio } = await post('document:native-teaching', { binding, operation: 'audio', assetId: asset.id });
      const context = new AudioContext(), decoded = await context.decodeAudioData(await audio.arrayBuffer());
      if (decoded.duration <= 0) throw new Error('Generated embedded speech cannot decode');
      await context.close(); audioCount++;
    }
    return { teaching, audioCount };
  }, binding);
  expect(reopened.audioCount).toBe(4);
  expect(reopened.teaching.pages).toEqual(mediaResult.teaching.pages);
  expect(reopened.teaching.identity.pages[0].objects.some((o: any) => o.logicalId === mediaResult.objectId)).toBe(true);
  await page.evaluate(async binding => { (document.querySelector('#editorFrame') as HTMLIFrameElement).contentWindow!.postMessage({ type: 'document:teaching-enable', payload: { binding } }, location.origin); }, binding);
  const reopenedNative = page.frames().find(f => f.name() === 'frameEditor')!;
  await reopenedNative.evaluate(() => (Reflect.get(window, 'editor') as unknown as { goToPage(index: number): void }).goToPage(1));
  await expect(panel.locator(`article[data-action-id="${initial.pages[1].actions[0].id}"]`)).toBeVisible();
  await panel.getByRole('button', { name: '画布预览', exact: true }).click();
  const question = office.locator('[data-teaching-question]');
  await expect(question).toContainText('Choose the greeting', { timeout: 15000 });
  await question.getByRole('radio').nth(1).check();
  await question.getByRole('button', { name: '检查答案', exact: true }).click();
  await expect(question).toContainText('回答正确');
  await question.getByRole('button', { name: /^(继续教学|教师继续)$/ }).click();
  await expect(question).toHaveCount(0);
  await expect(panel.getByRole('alert')).toHaveCount(0);
});
