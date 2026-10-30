import { mkdir, copyFile, readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { buildPptx, toBase64, makeStoredZip, zipEntryNames, zipEntryText } from './lib/ooxml';
import { expect, test } from './lib/l0';
const binding = { projectId: 'teaching-probe', documentId: 'fixture', sessionId: 'isolated' };

test('native teaching sidebar edits the Host timeline and reloads saved narration', async ({ page }) => {
  test.setTimeout(240_000);
  const runtimePath = process.env.DECKWISE_PREVIEW_RUNTIME;
  if (runtimePath) {
    const directory = resolve(process.env.E2E_PORT ? `dist-e2e-${process.env.E2E_PORT}` : 'dist', '.deckwise');
    await mkdir(directory, { recursive: true });
    await copyFile(runtimePath, resolve(directory, 'teaching-preview.mjs'));
  }
  const original = buildPptx('Native teaching sidebar');
  const entries = await Promise.all(
    zipEntryNames(original).map(async (name) => ({ name, data: (await zipEntryText(original, name))! })),
  );
  entries.push({
    name: 'ppt/slides/slide2.xml',
    data: entries
      .find((entry) => entry.name === 'ppt/slides/slide1.xml')!
      .data.replace('Native teaching sidebar', 'Second teaching page'),
  });
  entries.push({
    name: 'ppt/slides/_rels/slide2.xml.rels',
    data: entries.find((entry) => entry.name === 'ppt/slides/_rels/slide1.xml.rels')!.data,
  });
  entries.find((entry) => entry.name === 'ppt/presentation.xml')!.data = entries
    .find((entry) => entry.name === 'ppt/presentation.xml')!
    .data.replace('</p:sldIdLst>', '<p:sldId id="257" r:id="rId3"/></p:sldIdLst>');
  entries.find((entry) => entry.name === 'ppt/_rels/presentation.xml.rels')!.data = entries
    .find((entry) => entry.name === 'ppt/_rels/presentation.xml.rels')!
    .data.replace(
      '</Relationships>',
      '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide2.xml"/></Relationships>',
    );
  entries.find((entry) => entry.name === '[Content_Types].xml')!.data = entries
    .find((entry) => entry.name === '[Content_Types].xml')!
    .data.replace(
      '</Types>',
      '<Override PartName="/ppt/slides/slide2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>',
    );
  const twoPages = makeStoredZip(entries);
  await page.goto('/embed-demo.html');
  await page.evaluate(() => {
    document.querySelector('#status')!.textContent = 'loading';
    (document.querySelector('#editorFrame') as HTMLIFrameElement).src = '/editor?embed=1&deckwise=1&locale=zh-CN';
  });
  await expect(page.locator('#status')).toHaveText('ready', { timeout: 60_000 });
  await page.evaluate(
    async ({ b64, binding }) => {
      const bytes = Uint8Array.from(atob(b64), (c) => c.charCodeAt(0));
      (window as any).fixtureHash = Array.from(new Uint8Array(await crypto.subtle.digest('SHA-256', bytes)), (b) =>
        b.toString(16).padStart(2, '0'),
      ).join('');
      await post('document:open-buffer', {
        fileName: 'teaching.pptx',
        buffer: bytes.buffer,
        editorBinding: binding,
        trackPageLineage: true,
      });
    },
    { b64: toBase64(twoPages), binding },
  );
  await page.waitForFunction(() => !!window.__ooFrames.readyEditor(), { timeout: 60_000 });
  if (runtimePath) {
    const editorFrame = page.frames().find((frame) => frame.name() === 'frameEditor')!;
    await editorFrame.evaluate(() => {
      const api = (window as any).editor,
        logic = api.WordControl.m_oLogicDocument;
      const controller = logic.GetCurrentController();
      controller.resetSelection();
      controller.selectObject(logic.Slides[0].cSld.spTree[0], 0);
      api.asc_AddAnimation(1, 10, 0, null, false, false);
    });
    await page.evaluate(async (binding) => {
      const saved = await post('document:save', {});
      const buffer = await saved.file.arrayBuffer();
      (window as any).fixtureHash = Array.from(new Uint8Array(await crypto.subtle.digest('SHA-256', buffer)), (b) =>
        b.toString(16).padStart(2, '0'),
      ).join('');
      await post('document:open-buffer', {
        fileName: 'teaching.pptx',
        buffer,
        editorBinding: binding,
        trackPageLineage: true,
      });
    }, binding);
    await page.waitForFunction(() => !!window.__ooFrames.readyEditor(), { timeout: 60_000 });
  }
  const media = {
    music: (await readFile(new URL('./fixtures/native-teaching/music.mp3', import.meta.url))).toString('base64'),
    video: (await readFile(new URL('./fixtures/native-teaching/video.mp4', import.meta.url))).toString('base64'),
  };
  await page.evaluate(
    async ({ binding, media }) => {
      await post('document:get-context', binding);
      const frame = document.querySelector('#editorFrame') as HTMLIFrameElement;
      const project = {
        teachingRevision: 'rev-1',
        source: { sha256: (window as any).fixtureHash },
        pages: [
          {
            pageId: 'p1',
            locked: false,
            reviewState: 'ready',
            actions: [
              {
                id: 's1',
                type: 'speech',
                text: 'Hello, class!',
                audioAssetId: 'tts-old',
                durationMs: 3000,
                visualCues: [
                  { id: 'c1', atMs: 200, endMs: 2200, effect: { type: 'highlight', elementIds: ['title'] } },
                  { id: 'c2', atMs: 250, endMs: 2150, effect: { type: 'spotlight', elementId: 'title' } },
                ],
              },
              {
                id: 'wait',
                type: 'interaction_checkpoint',
                checkpointId: 'teacher-wait',
                checkpointKind: 'teacher_continue',
              },
              { id: 'delay', type: 'delay', durationMs: 1200 },
              { id: 'music', type: 'audio_clip', audioAssetId: 'music-existing' },
              { id: 'video', type: 'video_clip', videoAssetId: 'video-existing' },
              { id: 'video2', type: 'video_clip', videoAssetId: 'video-existing' },
              { id: 's2', type: 'speech', text: 'Goodbye, class!', audioAssetId: 'tts-next', durationMs: 3000 },
            ],
          },
          {
            pageId: 'p2',
            locked: false,
            reviewState: 'ready',
            actions: [
              {
                id: 'page2-speech',
                type: 'speech',
                text: 'Second teaching page',
                audioAssetId: 'page2-tts',
                durationMs: 3000,
              },
              {
                id: 'page2-wait',
                type: 'interaction_checkpoint',
                checkpointId: 'page2-wait',
                checkpointKind: 'teacher_continue',
              },
            ],
          },
        ],
      };
      (window as any).teachingFixture = project;
      let draft: any = null;
      (window as any).teachingOperations = [];
      window.addEventListener('message', (event) => {
        if (
          event.source !== frame.contentWindow ||
          event.origin !== location.origin ||
          event.data?.type !== 'document:teaching-request'
        )
          return;
        const { operation, args } = event.data.payload;
        (window as any).teachingOperations.push(operation);
        const number = event.data.payload.context.page;
        const view = () => ({
          viewId: `view-${number}`,
          pageId: `p${number}`,
          pageNumber: number,
          project: structuredClone(project),
          teachingObjects: [
            {
              objectId: 'title',
              sourceShapeId: '2',
              sourceObjectName: 'Title 1',
              name: 'Title',
              status: 'mapped',
              effectTargetEligible: true,
            },
          ],
        });
        let result: any;
        switch (operation) {
          case 'open':
            result = view();
            break;
          case 'draft:get':
            result = { draft };
            break;
          case 'draft:set':
            draft = args.draft ? { ...args.draft, baseline: structuredClone(project), revision: 'draft-rev' } : null;
            result = { draft };
            break;
          case 'write':
            project.pages[0].actions = args.actions;
            project.teachingRevision = 'rev-2';
            result = view();
            break;
          case 'media:list':
            result = [
              {
                kind: 'video',
                intentId: 'saved-video',
                title: 'Existing classroom video',
                model: 'fixture',
                inPage: false,
              },
            ];
            break;
          case 'media:adopt':
            (window as any).lastMediaAction = args.actionId;
            project.pages[0].actions.push({
              id: args.actionId,
              type: 'video_clip',
              videoAssetId: 'video-existing',
            });
            project.teachingRevision = 'rev-media';
            result = view();
            break;
          case 'media:preview': {
            const isVideo = project.pages[0].actions.some((a) => a.id === args.actionId && a.type === 'video_clip');
            result = {
              bytes: Uint8Array.from(atob(isVideo ? media.video : media.music), (c) => c.charCodeAt(0)),
              mimeType: isVideo ? 'video/mp4' : 'audio/mpeg',
            };
            break;
          }
          case 'tts:preview': {
            const samples = 24000,
              bytes = new Uint8Array(44 + samples * 2),
              wav = new DataView(bytes.buffer);
            for (const [at, str] of [
              [0, 'RIFF'],
              [8, 'WAVE'],
              [12, 'fmt '],
              [36, 'data'],
            ] as const)
              for (let i = 0; i < str.length; i++) bytes[at + i] = str.charCodeAt(i);
            wav.setUint32(4, bytes.length - 8, true);
            wav.setUint32(16, 16, true);
            wav.setUint16(20, 1, true);
            wav.setUint16(22, 1, true);
            wav.setUint32(24, 8000, true);
            wav.setUint32(28, 16000, true);
            wav.setUint16(32, 2, true);
            wav.setUint16(34, 16, true);
            wav.setUint32(40, samples * 2, true);
            result = { bytes, mimeType: 'audio/wav' };
            break;
          }
          case 'tts:models':
          case 'tts:list':
            result = [];
            break;
          default:
            throw new Error('Unexpected paid operation ' + operation);
        }
        frame.contentWindow!.postMessage(
          { type: 'document:teaching-response', id: event.data.id, payload: { result } },
          location.origin,
        );
      });
      frame.contentWindow!.postMessage({ type: 'document:teaching-enable', payload: { binding } }, location.origin);
    },
    { binding, media },
  );
  const native = page.frameLocator('#editorFrame').frameLocator('#app iframe');
  const panel = native.getByRole('complementary', { name: '教学时间轴' });
  await expect(panel).toBeVisible({ timeout: 30_000 });
  const narration = panel.getByRole('textbox', { name: '第 1 段讲解' });
  await expect(narration).toHaveValue('Hello, class!');
  await expect(panel.locator('.teaching-flow article')).toHaveCount(7);
  await expect(panel.locator('[data-teaching-inspector]')).toHaveCount(1);
  await expect(panel.locator('textarea')).toHaveCount(1);
  await expect(panel.locator('article[data-selected] > [data-teaching-inspector]')).toHaveCount(1);
  expect((await panel.boundingBox())!.width).toBeGreaterThanOrEqual(350);
  const canvasBounds = (await native.locator('#editor_sdk').boundingBox())!;
  expect(canvasBounds.x + canvasBounds.width).toBeLessThanOrEqual((await panel.boundingBox())!.x + 2);
  expect(await panel.locator('[data-teaching-inspector]').evaluate((node) => getComputedStyle(node).overflowY)).toBe(
    'visible',
  );
  await page.locator('#editorFrame').screenshot({ path: 'test-results-native-teaching-redesign.png' });
  await expect(panel.getByRole('button', { name: '添加停顿', exact: true })).toHaveCount(0);
  await expect(panel.getByRole('button', { name: '添加等待教师继续', exact: true })).toHaveCount(0);
  const after = panel.getByLabel('讲后', { exact: true });
  await after.selectOption('pause');
  await expect(panel.getByRole('status')).toHaveText('本页教学已保存');
  expect(await page.evaluate(() => (window as any).teachingFixture.pages[0].actions[0].pauseAfterMs)).toBe(400);
  await after.selectOption('timed');
  const pauseSeconds = panel.getByRole('spinbutton', { name: '讲后停顿秒数', exact: true });
  await expect(pauseSeconds).toHaveValue('3');
  await pauseSeconds.fill('3.5');
  await pauseSeconds.blur();
  await expect(panel.getByRole('status')).toHaveText('本页教学已保存');
  expect(await page.evaluate(() => (window as any).teachingFixture.pages[0].actions[0].pauseAfterMs)).toBe(3500);
  expect(await page.evaluate(() => (window as any).teachingFixture.pages[0].actions[0].audioAssetId)).toBe('tts-old');
  await after.selectOption('none');
  await expect(panel.getByRole('status')).toHaveText('本页教学已保存');

  const nativeFrame = page.frames().find((f) => f.name() === 'frameEditor')!;
  if (runtimePath) {
    const history = await nativeFrame.evaluate(() => (window as any).AscCommon.History.Index);
    await panel.getByRole('button', { name: '幻灯片放映', exact: true }).click();
    const show = native.getByRole('region', { name: '本页教学放映' });
    await expect(show).toBeVisible();
    expect(
      await nativeFrame.evaluate(() => (window as any).editor.WordControl.DemonstrationManager.IsMainSeqFinished(0)),
    ).toBe(false);
    await expect(panel.locator('article[data-playing]')).toHaveCount(0);
    await expect(show.locator('[data-native-animation-controls]')).toContainText('PPT 动画 0/1');
    await show.getByRole('button', { name: '暂停 PPT 动画', exact: true }).click();
    await expect(show.getByRole('button', { name: '下一步 PPT 动画', exact: true })).toBeDisabled();
    await show.getByRole('button', { name: '继续 PPT 动画', exact: true }).click();
    await show.getByRole('button', { name: '下一步 PPT 动画', exact: true }).click();
    await expect(show.locator('[data-native-animation-controls]')).toHaveCount(0);
    await expect(show.locator('[data-attention-kind=spotlight]')).toBeVisible();
    await show.getByRole('button', { name: '暂停播放', exact: true }).click();
    await nativeFrame.evaluate(() => {
      (window as any).teachingClickErrors = [];
      (window as any).editor.asc_registerCallback('asc_onError', (...args: any[]) =>
        (window as any).teachingClickErrors.push(args),
      );
    });
    for (let i = 0; i < 3; i++) await show.locator('[data-deckwise-teaching-canvas] canvas').first().click();
    expect(await nativeFrame.evaluate(() => (window as any).teachingClickErrors)).toEqual([]);
    expect(await nativeFrame.evaluate(() => (window as any).editor.WordControl.DemonstrationManager.SlideNum)).toBe(0);
    await expect(show).toBeVisible();
    await expect(show.locator('[data-attention-kind=highlight]')).toBeVisible();
    const actual = await show.locator('[data-teaching-preview] svg').boundingBox();
    const expected = await nativeFrame.evaluate(() => {
      const manager = (window as any).editor.WordControl.DemonstrationManager;
      const rect = manager.Transition.Rect,
        canvas = manager.Canvas.getBoundingClientRect();
      return {
        x: canvas.x + (rect.x * canvas.width) / manager.Canvas.width,
        y: canvas.y + (rect.y * canvas.height) / manager.Canvas.height,
        width: (rect.w * canvas.width) / manager.Canvas.width,
        height: (rect.h * canvas.height) / manager.Canvas.height,
      };
    });
    const frameBounds = await page.frameLocator('#editorFrame').locator('#app iframe').boundingBox();
    expect(Math.abs(actual!.x - expected.x - frameBounds!.x)).toBeLessThan(2);
    expect(Math.abs(actual!.y - expected.y - frameBounds!.y)).toBeLessThan(2);
    expect(Math.abs(actual!.width - expected.width)).toBeLessThan(2);
    expect(Math.abs(actual!.height - expected.height)).toBeLessThan(2);
    await expect(show.getByRole('status')).toHaveText('本页教学已暂停');
    await show.getByRole('button', { name: '继续播放', exact: true }).click();
    await expect(show.getByRole('button', { name: '教师继续', exact: true })).toBeEnabled({ timeout: 10000 });
    await show.getByRole('button', { name: '教师继续', exact: true }).click();
    await expect(panel.locator('article[data-action-id=music]')).toHaveAttribute('data-playing', '', {
      timeout: 10000,
    });
    const video = show.getByLabel('本页教学视频');
    await expect(video).toBeVisible({ timeout: 10000 });
    await show.getByRole('button', { name: '暂停播放', exact: true }).click();
    expect(await video.evaluate((v: HTMLVideoElement) => v.paused)).toBe(true);
    await show.getByRole('button', { name: '继续播放', exact: true }).click();
    await expect(panel.locator('article[data-action-id=video2]')).toHaveAttribute('data-playing', '', {
      timeout: 10000,
    });
    await expect(panel.locator('article[data-action-id=s2]')).toHaveAttribute('data-playing', '', { timeout: 10000 });
    await expect(show).toHaveCount(0, { timeout: 10000 });
    await expect(panel.getByRole('status')).toHaveText('本页教学播放完成');
    expect(await nativeFrame.evaluate(() => (window as any).AscCommon.History.Index)).toBe(history);
    await panel.getByRole('button', { name: '幻灯片放映', exact: true }).click();
    await show.getByRole('button', { name: '下一步 PPT 动画', exact: true }).click();
    await expect(show.locator('[data-attention-kind=highlight]')).toBeVisible();
    await show.getByRole('button', { name: '退出教学放映', exact: true }).click();
    await expect(show).toHaveCount(0);
    await expect(native.locator('[data-teaching-preview]')).toHaveCount(0);
    await expect(panel.locator('article[data-playing]')).toHaveCount(0);
    await panel.getByRole('button', { name: '幻灯片放映', exact: true }).click();
    await show.getByRole('button', { name: '下一步 PPT 动画', exact: true }).click();
    await expect(show.locator('[data-attention-kind=highlight]')).toBeVisible();
    await show.locator('[data-deckwise-teaching-canvas] canvas').first().press('Escape');
    await expect(show).toHaveCount(0);
    await expect(native.locator('[data-teaching-preview]')).toHaveCount(0);
    await expect(panel.locator('article[data-playing]')).toHaveCount(0);
  }
  if (runtimePath) {
    await panel.getByRole('button', { name: '放映整课教学', exact: true }).click();
    const firstPage = native.locator('[data-teaching-slideshow][data-teaching-page="1"]');
    await expect(firstPage).toBeVisible();
    await firstPage.getByRole('button', { name: '下一步 PPT 动画', exact: true }).click();
    await expect(firstPage.getByRole('button', { name: '教师继续', exact: true })).toBeEnabled({ timeout: 10000 });
    await firstPage.getByRole('button', { name: '教师继续', exact: true }).click();
    const secondPage = native.locator('[data-teaching-slideshow][data-teaching-page="2"]');
    await expect(secondPage).toBeVisible({ timeout: 20000 });
    await expect(panel.locator('article[data-action-id=page2-speech]')).toHaveAttribute('data-playing', '');
    await expect(secondPage.getByRole('button', { name: '教师继续', exact: true })).toBeEnabled({ timeout: 10000 });
    await secondPage.getByRole('button', { name: '教师继续', exact: true }).click();
    await expect(native.locator('[data-teaching-slideshow]')).toHaveCount(0);
    await expect(native.locator('[data-teaching-lesson]')).toHaveCount(0);
    await expect(panel.getByRole('status')).toHaveText('整课教学放映完成');
    await expect(narration).toHaveValue('Hello, class!');
    await panel.getByRole('button', { name: '放映整课教学', exact: true }).click();
    await expect(firstPage).toBeVisible();
    await firstPage.getByRole('button', { name: '下一步 PPT 动画', exact: true }).click();
    await expect(firstPage.locator('[data-attention-kind=highlight]')).toBeVisible();
    await firstPage.locator('[data-deckwise-teaching-canvas] canvas').first().press('Escape');
    await expect(native.locator('[data-teaching-slideshow]')).toHaveCount(0);
    await expect(native.locator('[data-teaching-lesson]')).toHaveCount(0);
    await expect(panel.locator('article[data-playing]')).toHaveCount(0);
    await expect(native.locator('[data-teaching-preview]')).toHaveCount(0);
  }
  if (runtimePath) {
    await panel.getByRole('button', { name: '画布预览', exact: true }).click();
    await expect(panel.locator('article[data-action-id=s1]')).toHaveAttribute('data-playing', '');
    await nativeFrame.evaluate(() => {
      const host = document.createElement('div');
      host.id = 'teaching-native-slideshow-probe';
      host.style.cssText = 'position:fixed;inset:0;z-index:10000;background:black';
      document.body.append(host);
      (window as any).editor.StartDemonstration(host.id, 0);
    });
    await expect(panel.getByRole('status')).toHaveText(
      '原生放映已开始，教学预览已停止；原生放映暂不自动播放教学时间轴。',
    );
    await expect(panel.locator('article[data-playing]')).toHaveCount(0);
    await expect(panel.getByRole('button', { name: '画布预览', exact: true })).toBeDisabled();
    await expect(native.locator('[data-teaching-preview]')).toHaveCount(0);
    await expect(panel.locator('video')).toHaveCount(0);
    await nativeFrame.evaluate(() => {
      (window as any).editor.EndDemonstration();
      document.getElementById('teaching-native-slideshow-probe')!.remove();
    });
    await expect(panel.getByRole('status')).toHaveText('已退出原生放映，可重新预览本页教学');
    await expect(panel.getByRole('button', { name: '画布预览', exact: true })).toBeEnabled();
  }
  if (runtimePath) {
    await panel.getByText('添加聚光', { exact: true }).first().click();
    const history = await nativeFrame.evaluate(() => (window as any).AscCommon.History.Index);
    await panel.getByRole('button', { name: '预览此提示', exact: true }).first().click();
    const overlay = native.locator('[data-teaching-preview]');
    await expect(overlay.locator('[data-attention-kind=highlight]')).toBeVisible();
    for (const zoom of [50, 100, 150]) {
      await nativeFrame.evaluate((zoom) => (window as any).editor.zoom(zoom), zoom);
      await expect
        .poll(() => overlay.locator('svg').evaluate((node) => parseFloat((node as SVGElement).style.width)))
        .toBeGreaterThan(100);
      // Overlay and native SDK use the same CSS viewport coordinates, including editor chrome.
      const actual = await overlay.locator('svg').boundingBox();
      const expected = await nativeFrame.evaluate(() => {
        const wc = (window as any).editor.WordControl,
          logic = wc.m_oLogicDocument;
        return {
          from: wc.m_oDrawingDocument.ConvertCoordsToCursor3(0, 0, true),
          to: wc.m_oDrawingDocument.ConvertCoordsToCursor3(logic.GetWidthMM(), logic.GetHeightMM(), true),
        };
      });
      const frameBounds = await page.frameLocator('#editorFrame').locator('#app iframe').boundingBox();
      expect(Math.abs(actual!.x - expected.from.X - frameBounds!.x)).toBeLessThan(2);
      expect(Math.abs(actual!.width - (expected.to.X - expected.from.X))).toBeLessThan(2);
    }
    await page.screenshot({ path: 'test-results-native-attention.png' });
    await panel.getByRole('button', { name: '停止预览', exact: true }).click();
    await expect(overlay).toHaveCount(0);
    await nativeFrame.evaluate(() => (window as any).editor.zoom(70));
    await panel.getByRole('button', { name: '试听语音', exact: true }).first().click();
    const audio = panel.locator('audio');
    await expect(overlay.locator('[data-attention-kind=spotlight]')).toBeVisible();
    await audio.evaluate((node: HTMLAudioElement) => node.pause());
    await expect.poll(() => audio.evaluate((node: HTMLAudioElement) => node.paused)).toBe(true);
    await audio.evaluate((node: HTMLAudioElement) => {
      node.currentTime = 2.5;
    });
    await expect(overlay.locator('[data-attention-kind]')).toHaveCount(0);
    await audio.evaluate((node: HTMLAudioElement) => {
      node.currentTime = 0.7;
    });
    await expect(overlay.locator('[data-attention-kind=highlight]')).toBeVisible();
    expect(await nativeFrame.evaluate(() => (window as any).AscCommon.History.Index)).toBe(history);
    await panel.getByRole('button', { name: '停止预览', exact: true }).click();
    await expect(overlay).toHaveCount(0);
  }
  if (runtimePath) {
    await panel.getByRole('button', { name: '画布预览', exact: true }).click();
    await expect(panel.locator('article[data-action-id=s1]')).toHaveAttribute('data-playing', '');
    await panel.getByRole('button', { name: '暂停播放', exact: true }).click();
    await expect(panel.getByRole('status')).toHaveText('本页教学已暂停');
    await panel.getByRole('button', { name: '继续播放', exact: true }).click();
    await expect(panel.getByRole('button', { name: '教师继续', exact: true })).toBeEnabled({ timeout: 10000 });
    await expect(panel.getByRole('status')).toHaveText('等待教师继续，点击「教师继续」后播放下一段');
    await expect(panel.getByRole('button', { name: '继续播放', exact: true })).toBeDisabled();
    await panel.getByRole('button', { name: '教师继续', exact: true }).click();
    await expect(panel.locator('article[data-action-id=delay]')).toHaveAttribute('data-playing', '');
    await panel.getByRole('button', { name: '暂停播放', exact: true }).click();
    await expect(panel.getByRole('status')).toHaveText('本页教学已暂停');
    await page.waitForTimeout(1400);
    await expect(panel.locator('article[data-action-id=delay]')).toHaveAttribute('data-playing', '');
    await expect(panel.locator('[data-preview-command=teacher]')).toBeDisabled();
    await panel.getByRole('button', { name: '继续播放', exact: true }).click();
    await expect(panel.locator('article[data-action-id=music]')).toHaveAttribute('data-playing', '');
    const video = panel.getByLabel('本页教学视频');
    await expect(video).toBeVisible({ timeout: 10000 });
    await expect(panel.locator('[data-preview-command=teacher]')).toBeDisabled();
    await panel.getByRole('button', { name: '暂停播放', exact: true }).click();
    await expect.poll(() => video.evaluate((v: HTMLVideoElement) => v.paused)).toBe(true);
    const pausedAt = await video.evaluate((v: HTMLVideoElement) => v.currentTime);
    await expect(panel.getByRole('status')).toHaveText('本页教学已暂停');
    expect(await video.evaluate((v: HTMLVideoElement) => v.currentTime)).toBeCloseTo(pausedAt, 1);
    await panel.getByRole('button', { name: '继续播放', exact: true }).click();
    await expect(panel.locator('article[data-action-id=video2]')).toHaveAttribute('data-playing', '', {
      timeout: 10000,
    });
    await expect(panel.getByRole('button', { name: '暂停播放', exact: true })).toBeEnabled();
    await expect(panel.locator('article[data-action-id=s2]')).toHaveAttribute('data-playing', '', { timeout: 10000 });
    await expect(video).toBeHidden();
    await expect(panel.getByRole('status')).toHaveText('本页教学播放完成', { timeout: 10000 });
    await expect(native.locator('[data-teaching-preview]')).toHaveCount(0);
  }
  if (runtimePath) {
    const history = await nativeFrame.evaluate(() => (window as any).AscCommon.History.Index);
    const before = await panel.locator('article').count();
    await panel.locator('footer > details').last().locator('summary').click();
    await panel.getByRole('button', { name: '绑定本页 PPT 动画', exact: true }).click();
    await expect(panel.locator('article')).toHaveCount(before + 1);
    const nodeId = await panel.locator('article').last().getAttribute('data-action-id');
    const node = panel.locator(`article[data-action-id="${nodeId}"]`);
    await expect(node).toContainText('PPT 动画 · 第 1 步');
    await node.locator('.teaching-step').click();
    for (let i = 0; i < before; i++) await node.getByRole('button', { name: '上移', exact: true }).click();
    await expect(panel.locator('article').first()).toHaveAttribute('data-action-id', nodeId!);
    await panel.getByRole('button', { name: '保存本页教学', exact: true }).click();
    await expect(panel.getByRole('button', { name: '保存本页教学', exact: true })).toBeDisabled();
    await expect(panel.locator('article').first()).toContainText('PPT 动画 · 第 1 步');
    await panel.locator('[data-teaching-inspector]').getByRole('button', { name: '编辑原生动画', exact: true }).click();
    await expect(panel).toBeHidden();
    expect((await native.locator('#right-menu').boundingBox())!.width).toBeLessThan(350);
    await expect(native.locator('#animation-panel')).toBeVisible();
    expect(
      await nativeFrame.evaluate(() => {
        const logic = (window as any).editor.WordControl.m_oLogicDocument;
        return logic.GetCurrentController().selectedObjects.includes(logic.Slides[0].cSld.spTree[0]);
      }),
    ).toBe(true);
    expect(await nativeFrame.evaluate(() => (window as any).AscCommon.History.Index)).toBe(history);
    await nativeFrame.evaluate(() => {
      const logic = (window as any).editor.WordControl.m_oLogicDocument;
      logic.GetCurrentController().resetSelection();
      logic.Document_UpdateSelectionState();
    });
    await native.getByRole('button', { name: '教学时间轴', exact: true }).click();
    await expect(panel).toBeVisible();
    await panel.getByRole('button', { name: '幻灯片放映', exact: true }).click();
    const show = native.locator('[data-teaching-slideshow]');
    await expect(show).toBeVisible();
    await expect(show.locator('[data-native-animation-controls]')).toHaveCount(0);
    await expect(show.locator('[data-attention-kind=highlight]')).toBeVisible({ timeout: 10000 });
    expect(
      await nativeFrame.evaluate(() => (window as any).editor.WordControl.DemonstrationManager.IsMainSeqFinished(0)),
    ).toBe(true);
    await show.getByRole('button', { name: '暂停播放', exact: true }).click();
    await expect(show.getByRole('status')).toHaveText('本页教学已暂停');
    await show.locator('[data-deckwise-teaching-canvas] canvas').first().click();
    await expect(show).toBeVisible();
    await expect(show.getByRole('status')).toHaveText('本页教学已暂停');
    await show.getByRole('button', { name: '继续播放', exact: true }).click();
    await expect(show.getByRole('button', { name: '教师继续', exact: true })).toBeEnabled({ timeout: 10000 });
    await show.locator('[data-deckwise-teaching-canvas] canvas').first().press('Escape');
    await expect(show).toHaveCount(0);
    await expect(native.locator('[data-teaching-preview]')).toHaveCount(0);
    await panel.locator('footer > details').last().locator('summary').click();
    await panel.getByRole('button', { name: '移除动画绑定', exact: true }).click();
    await panel.getByRole('button', { name: '保存本页教学', exact: true }).click();
    await expect(panel.locator('article')).toHaveCount(before);
    await expect(panel.getByRole('button', { name: '保存本页教学', exact: true })).toBeDisabled();
    expect(await nativeFrame.evaluate(() => (window as any).AscCommon.History.Index)).toBe(history);
  }
  await panel.getByText('添加已有音乐或视频', { exact: true }).click();
  await panel.getByLabel('已有音乐和视频').selectOption('0');
  await panel.getByRole('button', { name: '加入当前页', exact: true }).click();
  await expect(panel.getByRole('status')).toHaveText('已有素材已加入本页，可调整播放顺序');
  await expect(panel.locator('article')).toHaveCount(8);
  if (runtimePath) {
    const addedId = await page.evaluate(() => (window as any).lastMediaAction);
    await panel.getByRole('button', { name: '画布预览', exact: true }).click();
    await expect(panel.getByRole('button', { name: '教师继续', exact: true })).toBeEnabled({ timeout: 10000 });
    await panel.getByRole('button', { name: '教师继续', exact: true }).click();
    await expect(panel.locator(`article[data-action-id="${addedId}"]`)).toHaveAttribute('data-playing', '', {
      timeout: 20000,
    });
    await expect(panel.getByRole('status')).toHaveText('本页教学播放完成', { timeout: 10000 });
  }
  if (runtimePath) {
    const sdkHistory = await nativeFrame.evaluate(() => (window as any).AscCommon.History.Index);
    const cue = panel.locator('[data-cue-id=c1]');
    const opens = await page.evaluate(
      () => (window as any).teachingOperations.filter((op: string) => op === 'open').length,
    );
    await narration.evaluate((node) => {
      (node as any).selectionIdentity = true;
    });
    await nativeFrame.evaluate(() => {
      const logic = (window as any).editor.WordControl.m_oLogicDocument;
      const controller = logic.GetCurrentController();
      controller.resetSelection();
      controller.selectObject(logic.Slides[0].cSld.spTree[0], 0);
      document.dispatchEvent(new Event('pointerup'));
    });
    await expect
      .poll(() => page.evaluate(() => (window as any).teachingOperations.filter((op: string) => op === 'open').length))
      .toBeGreaterThan(opens);
    await expect(panel.getByRole('button', { name: '刷新', exact: true })).toBeEnabled();
    expect(await narration.evaluate((node) => (node as any).selectionIdentity)).toBe(true);
    await cue.getByRole('button', { name: '编辑此提示', exact: true }).click();
    await panel.getByRole('button', { name: '使用画布选中对象', exact: true }).first().click();
    await expect(panel.getByRole('status')).toHaveText('已选择 1 个画布对象');
    await panel.getByLabel('高亮样式', { exact: true }).first().selectOption('fill');
    await panel.getByLabel('高亮颜色', { exact: true }).first().fill('#123456');
    await panel.getByLabel('效果不透明度百分比', { exact: true }).first().fill('70');
    await panel.getByRole('button', { name: '更新教学提示', exact: true }).click();
    await cue.getByRole('button', { name: '预览此提示', exact: true }).click();
    const overlay = native.locator('[data-teaching-preview]');
    await expect(overlay.locator('[data-attention-kind=highlight]')).toHaveAttribute('fill', '#123456');
    await expect(overlay.locator('[data-attention-kind=highlight]')).toHaveAttribute('opacity', '0.7');
    await panel.getByRole('button', { name: '停止预览', exact: true }).click();
    await cue.getByRole('button', { name: '编辑此提示', exact: true }).click();
    await panel.getByLabel('教学提示类型', { exact: true }).first().selectOption('spotlight');
    await panel.getByLabel('效果不透明度百分比', { exact: true }).first().fill('65');
    await panel.getByRole('button', { name: '更新教学提示', exact: true }).click();
    await expect(cue).toContainText('聚光');
    await cue.getByRole('button', { name: '预览此提示', exact: true }).click();
    await expect(overlay.locator('[data-attention-kind=spotlight]')).toBeVisible();
    await panel.getByRole('button', { name: '停止预览', exact: true }).click();
    await expect(panel.getByRole('status')).toHaveText('本页教学已保存');
    await expect(panel.getByRole('button', { name: '画布预览', exact: true })).toBeEnabled();
    await panel.getByRole('button', { name: '刷新', exact: true }).click();
    await expect(cue).toContainText('聚光');
    expect(await nativeFrame.evaluate(() => (window as any).AscCommon.History.Index)).toBe(sdkHistory);
  }
  await narration.fill('Autosaved narration');
  await narration.evaluate((node: HTMLTextAreaElement) => node.setSelectionRange(4, 4));
  await expect(panel.getByRole('status')).toHaveText('本页教学已保存');
  await expect(narration).toBeFocused();
  expect(await narration.evaluate((node: HTMLTextAreaElement) => node.selectionStart)).toBe(4);
  const beforeDraftHistory = await nativeFrame.evaluate(() => (window as any).AscCommon.History.Index);
  const previousNarration = await narration.inputValue();
  await narration.fill('Hello, friends!');
  await panel.getByRole('button', { name: '撤销', exact: true }).click();
  await expect(narration).toHaveValue(previousNarration);
  await panel.getByRole('button', { name: '重做', exact: true }).click();
  await expect(narration).toHaveValue('Hello, friends!');
  expect(await nativeFrame.evaluate(() => (window as any).AscCommon.History.Index)).toBe(beforeDraftHistory);
  await panel.getByLabel('讲后', { exact: true }).selectOption('pause');
  await expect(panel.locator('article')).toHaveCount(8);
  await expect(panel.getByRole('status')).toHaveText('本页教学已保存');
  await expect(panel.getByRole('button', { name: '撤销', exact: true })).toBeEnabled();
  await expect(panel.getByRole('button', { name: '重做', exact: true })).toBeDisabled();
  await page.evaluate((binding) => {
    (document.querySelector('#editorFrame') as HTMLIFrameElement).contentWindow!.postMessage(
      { type: 'document:teaching-enable', payload: { binding } },
      location.origin,
    );
  }, binding);
  await expect(narration).toHaveValue('Hello, friends!');
  await expect(panel.locator('article')).toHaveCount(8);
  await native.locator('#id-right-menu-slide').click({ timeout: 5000 });
  await expect(panel).toBeHidden();
  await native.getByRole('button', { name: '教学时间轴', exact: true }).click();
  await expect(panel).toBeVisible();
  expect(await page.evaluate(() => (window as any).teachingOperations)).not.toContain('tts:submit');
  // A native edit must save and reload teaching without the user pressing Refresh.
  await page.evaluate(() => {
    const frame = (document.querySelector('#editorFrame') as HTMLIFrameElement).contentWindow!;
    (window as any).nativeAutomaticSave = null;
    window.addEventListener('message', async (event) => {
      if (event.source !== frame || event.origin !== location.origin || event.data?.type !== 'document:save-requested')
        return;
      const saved = await post('document:save', { requireHostCommit: true });
      const project = (window as any).teachingFixture;
      project.source.sha256 = saved.pageLineage.outputSha256;
      project.source.revision = saved.pageLineage.outputSha256;
      project.teachingRevision = 'native-save-revision';
      for (const p of project.pages)
        for (const action of p.actions)
          if (action.type === 'native_animation') action.sourceSha256 = saved.pageLineage.outputSha256;
      (window as any).fixtureHash = saved.pageLineage.outputSha256;
      (window as any).nativeAutomaticSave = saved.pageLineage;
      frame.postMessage(
        { id: saved.pageLineage.saveRequestId, type: 'document:commit-save', payload: { success: true } },
        location.origin,
      );
    });
  });
  const opensBeforeSave = await page.evaluate(
    () => (window as any).teachingOperations.filter((op: string) => op === 'open').length,
  );
  await nativeFrame.evaluate(() => {
    const w = window as any,
      api = w.editor,
      logic = api.WordControl.m_oLogicDocument;
    const shape = logic.Slides[0].cSld.spTree[0];
    logic.StartAction(w.AscDFH.historydescription_Presentation_ApplyShapeProps);
    shape.spPr.xfrm.setOffX(shape.spPr.xfrm.offX + 5);
    logic.Recalculate();
    logic.FinalizeAction();
    api.CheckChangedDocument();
    document.dispatchEvent(new Event('pointerup'));
  });
  await expect.poll(() => page.evaluate(() => !!(window as any).nativeAutomaticSave), { timeout: 60_000 }).toBe(true);
  await expect
    .poll(() => page.evaluate(() => (window as any).teachingOperations.filter((op: string) => op === 'open').length))
    .toBeGreaterThan(opensBeforeSave);
  await expect(narration).toHaveValue('Hello, friends!');
  await expect(panel.getByRole('button', { name: '刷新', exact: true })).toBeEnabled();
  await page.screenshot({ path: 'test-results-native-teaching.png' });
});
