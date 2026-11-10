import { resolve } from 'node:path';
import { readFile, writeFile, mkdir, copyFile } from 'node:fs/promises';
import { test, expect } from './lib/l0';
import { toBase64, zipEntryNames, zipEntryText } from './lib/ooxml';

test('real Part A teaching and adopted audio survive native edit, undo and PPTX reopen', async ({ page }, info) => {
  test.setTimeout(240000);
  const fixture = process.env.DECKWISE_NATIVE_REAL_FIXTURE;
  test.skip(!fixture, 'Requires an explicitly copied real courseware fixture');
  if (process.env.DECKWISE_PREVIEW_RUNTIME) {
    const directory = resolve(`dist-e2e-${process.env.E2E_PORT}`, '.deckwise');
    await mkdir(directory, { recursive: true });
    await copyFile(process.env.DECKWISE_PREVIEW_RUNTIME, resolve(directory, 'teaching-preview.mjs'));
  }
  const inventory = JSON.parse(await readFile(fixture + '/report.json', 'utf8'));
  const original = await readFile(fixture + '/native.pptx');
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
        fileName: 'family-native.pptx',
        buffer: Uint8Array.from(atob(data), (c) => c.charCodeAt(0)).buffer,
        editorBinding: binding,
        trackPageLineage: true,
      });
    },
    { data: toBase64(original), binding },
  );
  await page.waitForFunction(() => !!window.__ooFrames.readyEditor(), { timeout: 60000 });
  const report = await page.evaluate(async (binding) => {
    const { teaching } = await post('document:native-teaching', { binding, operation: 'read' });
    const first = teaching.pages[0],
      actions = structuredClone(first.actions);
    const speech = actions.find((a: { type: string }) => a.type === 'speech');
    let totalAudioBytes = 0;
    for (const asset of teaching.media) {
      const { audio } = await post('document:native-teaching', { binding, operation: 'audio', assetId: asset.id });
      const bytes = await audio.arrayBuffer();
      if (bytes.byteLength !== asset.byteLength) throw new Error('Embedded audio length mismatch');
      const decoded = await new OfflineAudioContext(1, 1, 44100).decodeAudioData(bytes.slice(0));
      if (!Number.isFinite(decoded.duration) || decoded.duration <= 0) throw new Error('Embedded audio cannot decode');
      totalAudioBytes += bytes.byteLength;
    }
    const audio = await post('document:native-teaching', { binding, operation: 'audio', assetId: speech.audioAssetId });
    const bytes = await audio.audio.arrayBuffer();
    let staleRejected = false;
    try {
      await post('document:native-teaching', {
        binding,
        operation: 'replace-page',
        command: { documentId: teaching.documentId, expectedRevision: 'stale', pageId: first.pageId, actions },
      });
    } catch {
      staleRejected = true;
    }
    if (!staleRejected) throw new Error('Stale teaching mutation was accepted');
    actions.push({ id: 'native-validation-delay', type: 'delay', durationMs: 500 });
    const { teaching: edited } = await post('document:native-teaching', {
      binding,
      operation: 'replace-page',
      command: { documentId: teaching.documentId, expectedRevision: teaching.revision, pageId: first.pageId, actions },
    });
    return {
      pages: teaching.pages.length,
      media: teaching.media.length,
      totalAudioBytes,
      audioBytes: bytes.byteLength,
      before: teaching.revision,
      after: edited.revision,
      pageId: first.pageId,
    };
  }, binding);
  expect(report.pages).toBe(inventory.pages);
  expect(report.media).toBe(inventory.media);
  expect(report.totalAudioBytes).toBe(inventory.mediaBytes);
  expect(report.audioBytes).toBeGreaterThan(100);
  expect(report.after).not.toBe(report.before);
  const native = page.frames().find((f) => f.name() === 'frameEditor')!;
  await native.evaluate(() => (Reflect.get(window, 'editor') as unknown as { Undo(): void }).Undo());
  expect(
    await page.evaluate(
      async (binding) => (await post('document:native-teaching', { binding, operation: 'read' })).teaching.revision,
      binding,
    ),
  ).toBe(report.before);
  await native.evaluate(() => (Reflect.get(window, 'editor') as unknown as { Redo(): void }).Redo());
  await page.evaluate(async binding => { (document.querySelector('#editorFrame') as HTMLIFrameElement).contentWindow!.postMessage({ type: 'document:teaching-enable', payload: { binding } }, location.origin); }, binding);
  const office = page.frameLocator('iframe#editorFrame').frameLocator('iframe[name="frameEditor"]');
  const panel = office.getByRole('complementary', { name: '教学时间轴' });
  await expect(panel).toBeVisible({ timeout: 15000 });
  await expect(office.locator('.teaching-flow article').first()).toBeVisible();
  const pages = await page.evaluate(async binding => (await post('document:native-teaching', { binding, operation: 'read' })).teaching.pages.map((page: any) => ({ firstActionId: page.actions[0]?.id, questions: page.actions.filter((action: any) => action.type === 'interaction_checkpoint') })), binding);
  const navigationTimings: Array<{ page: number; elapsedMs: number }> = [];
  for (const index of [2, 1, 3, 4, 5, 0]) {
    const start = Date.now();
    await native.evaluate(index => (Reflect.get(window, 'editor') as unknown as { goToPage(index: number): void }).goToPage(index), index);
    if (pages[index].firstActionId) {
      await expect(panel.locator(`.teaching-flow > article[data-action-id="${pages[index].firstActionId}"]`)).toBeVisible();
    } else {
      await expect(panel.getByText('本页还没有讲解，可添加讲解或在主 Chat 中生成。', { exact: true })).toBeVisible();
    }
    const elapsedMs = Date.now() - start;
    navigationTimings.push({ page: index + 1, elapsedMs });
    expect(elapsedMs, `Page ${index + 1} teaching navigation must not freeze the editor`).toBeLessThan(1500);
  }
  await info.attach('teaching-navigation-timings', { body: JSON.stringify(navigationTimings), contentType: 'application/json' });
  // The copied Quiz has a saved answer key but no optional explanation text.
  for (const name of ['画布预览', '幻灯片放映', '放映整课教学']) {
    await native.evaluate(() => (Reflect.get(window, 'editor') as unknown as { goToPage(index: number): void }).goToPage(4));
    await expect(panel.locator(`article[data-action-id="${pages[4].firstActionId}"]`)).toBeVisible();
    await panel.getByRole('button', { name, exact: true }).click();
    if (name === '放映整课教学') {
      // Empty page 3 advances immediately to page 4.
      let currentIndex = 0;
      for (const index of [1, 3, 4]) {
        await expect(panel.locator(`article[data-action-id="${pages[currentIndex].firstActionId}"][data-playing]`)).toBeVisible({ timeout: 20000 });
        await expect(office.locator('[data-teaching-slideshow]')).toBeVisible({ timeout: 20000 });
        await office.locator('[data-teaching-slideshow]').getByRole('button', { name: '下一页', exact: true }).click();
        await expect.poll(() => native.evaluate(() => (Reflect.get(window, 'editor') as unknown as { getCurrentPage(): number }).getCurrentPage())).toBe(index);
        currentIndex = index;
      }
    }
    const question = office.locator('[data-teaching-question]');
    await expect(question).toBeVisible({ timeout: 20000 });
    await expect(panel.getByRole('alert')).toHaveCount(0);
    if (process.env.DECKWISE_NATIVE_MODULE) {
      await expect(question.getByRole('button', { name: '检查答案', exact: true })).toBeDisabled();
    } else {
      await question.getByRole('button', { name: '检查答案', exact: true }).click();
      await expect(question.getByRole('status')).toHaveText('请先选择答案');
    }
    await question.getByRole('radio').nth(0).check();
    await question.getByRole('button', { name: '检查答案', exact: true }).click();
    await expect(question.getByRole('status')).toHaveText('再想一想，可以重新选择');
    await question.getByRole('radio').nth(1).check();
    await question.getByRole('button', { name: '检查答案', exact: true }).click();
    await expect(question.getByRole('status')).toHaveText('回答正确');
    await question.getByRole('button', { name: /^(继续教学|教师继续)$/ }).click();
    if (name === '放映整课教学') {
      await expect(question).toBeVisible({ timeout: 20000 });
      await expect(question).toContainText('谁来举起相册');
      await question.getByRole('button', { name: /^(继续教学|教师继续)$/ }).click();
    }
    await expect(question).toHaveCount(0);
    await expect(panel.getByRole('alert')).toHaveCount(0);
    const show = office.locator('[data-teaching-slideshow]');
    if (await show.count()) await show.getByRole('button', { name: '退出教学放映', exact: true }).click();
  }
  await native.evaluate(() => (Reflect.get(window, 'editor') as unknown as { goToPage(index: number): void }).goToPage(0));
  await expect(panel.locator(`article[data-action-id="${pages[0].firstActionId}"]`)).toBeVisible();
  // Exercise the copied lesson's original nested targets and native animations before replacing any cues.
  for (const name of ['画布预览', '幻灯片放映', '放映整课教学']) {
    await panel.getByRole('button', { name, exact: true }).click();
    const show = office.locator('[data-teaching-slideshow]');
    await expect(show).toBeVisible({ timeout: 20000 });
    await expect(office.locator('.teaching-flow article[data-playing]')).toHaveCount(1, { timeout: 20000 });
    await expect(office.locator('[data-attention-kind="spotlight"]')).toBeVisible({ timeout: 20000 });
    await expect(panel.getByRole('alert')).toHaveCount(0);
    await show.getByRole('button', { name: '退出教学放映', exact: true }).click();
    await expect(show).toHaveCount(0);
  }
  const narration = panel.getByRole('textbox', { name: '第 1 段讲解' });
  await expect(narration).toBeEnabled();
  const beforeText = await narration.inputValue();
  await narration.fill(beforeText + ' 验证原生编辑');
  await expect.poll(async () => page.evaluate(async binding => (await post('document:native-teaching', { binding, operation: 'read' })).teaching.pages[0].actions[0].text, binding)).toBe(beforeText + ' 验证原生编辑');
  await narration.blur();
  await panel.getByRole('button', { name: '撤销', exact: true }).click();
  await expect.poll(async () => page.evaluate(async binding => (await post('document:native-teaching', { binding, operation: 'read' })).teaching.pages[0].actions[0].text, binding)).toBe(beforeText);
  await expect(panel.getByRole('button', { name: '试听语音', exact: true })).toBeVisible();
  await panel.getByRole('button', { name: '试听语音', exact: true }).click();
  await expect.poll(() => panel.locator('audio').evaluate((audio: HTMLAudioElement) => audio.currentTime)).toBeGreaterThan(0);
  await panel.locator('audio').evaluate((audio: HTMLAudioElement) => audio.pause());
  const targetId = await page.evaluate(async binding => {
    const result = await post('document:native-teaching', { binding, operation: 'route', routeOperation: 'open', page: 1 });
    const view = result.routed.value;
    const target = view.teachingObjects.find((object: any) => object.effectTargetEligible);
    if (!target) throw new Error('Real slide has no native attention target');
    const actions = structuredClone(view.project.pages[0].actions);
    actions[0].visualCues = [
      { id: 'real-highlight', atMs: 0, endMs: 5000, effect: { type: 'highlight', elementIds: [target.objectId], style: 'outline', color: '#ff9900' } },
      { id: 'real-spotlight', atMs: 0, endMs: 5000, effect: { type: 'spotlight', elementId: target.objectId, dimOpacity: 0.6 } },
    ];
    await post('document:native-teaching', { binding, operation: 'route', routeOperation: 'write', page: 1,
      args: { viewId: view.viewId, expectedTeachingRevision: view.project.teachingRevision, actions } });
    return target.objectId;
  }, binding);
  await expect.poll(() => panel.getByRole('button', { name: '试听语音', exact: true }).isEnabled()).toBe(true);
  await panel.getByRole('button', { name: '试听语音', exact: true }).click();
  await expect(office.locator('[data-attention-kind="highlight"]')).toBeVisible();
  await expect(office.locator('[data-attention-kind="spotlight"]')).toBeVisible();
  await panel.locator('audio').evaluate((audio: HTMLAudioElement) => audio.pause());
  const hostCalls = await page.evaluate(() => (window as any).nativeHostCalls as string[]);
  expect(hostCalls.every(operation => ['tts:models', 'tts:list'].includes(operation))).toBe(true);
  const exported = await page.evaluate(async () =>
    Array.from(new Uint8Array(await (await post('document:save')).file.arrayBuffer())),
  );
  const output = new Uint8Array(exported);
  await writeFile(info.outputPath('real-native-edited.pptx'), output);
  await info.attach('real-native-edited', { path: info.outputPath('real-native-edited.pptx'), contentType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation' });
  const xml = (
    await Promise.all(
      zipEntryNames(output)
        .filter((p) => p.startsWith('customXml/') && p.endsWith('.xml'))
        .map((p) => zipEntryText(output, p)),
    )
  ).join('');
  expect(xml).toContain('native-validation-delay');
  await page.evaluate(
    async ({ bytes, binding }) => {
      await post('document:open-buffer', {
        fileName: 'reopened.pptx',
        buffer: new Uint8Array(bytes).buffer,
        editorBinding: binding,
        trackPageLineage: true,
      });
    },
    { bytes: exported, binding },
  );
  await page.waitForFunction(() => !!window.__ooFrames.readyEditor(), { timeout: 60000 });
  const reopened = await page.evaluate(async (binding) => {
    const { teaching } = await post('document:native-teaching', { binding, operation: 'read' });
    const action = teaching.pages[0].actions.find((a: { type: string }) => a.type === 'speech');
    const { audio } = await post('document:native-teaching', {
      binding,
      operation: 'audio',
      assetId: action.audioAssetId,
    });
    return { revision: teaching.revision, audioBytes: (await audio.arrayBuffer()).byteLength, cues: action.visualCues, identity: teaching.identity, assessments: teaching.pages.slice(4).map((page: any) => page.actions.filter((action: any) => action.type === "interaction_checkpoint")) };
  }, binding);
  expect(reopened.audioBytes).toBe(report.audioBytes);
  expect(reopened.cues[0].effect.elementIds).toEqual([targetId]);
  expect(reopened.identity.pages[0].objects.some((object: any) => object.logicalId === targetId)).toBe(true);
  if (inventory.assessmentPages) expect(reopened.assessments.every((actions: any[]) => actions.length > 0)).toBe(true);
  expect(reopened.assessments).toEqual(pages.slice(4).map((page: any) => page.questions));
  await info.attach('real-fixture-report', { body: JSON.stringify(report), contentType: 'application/json' });
});
