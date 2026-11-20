import { readFile, writeFile } from 'node:fs/promises';
import { createHash } from 'node:crypto';
import { test, expect } from './lib/l0';
import { toBase64 } from './lib/ooxml';

test('background teaching region preserves pixels/audio and survives native save/reopen', async ({ page }, info) => {
  test.setTimeout(180000);
  const fixture = process.env.DECKWISE_REGION_FIXTURE;
  test.skip(!fixture, 'Requires the saved four-page water-cycle real-model fixture');
  const original = await readFile(fixture!);
  const binding = { projectId: 'region-fixture', documentId: 'water-cycle', sessionId: 'isolated' };
  await page.goto('/embed-demo.html');
  await page.locator('#log').evaluate((node) => {
    node.style.display = 'none';
  });
  await page.evaluate(() => {
    document.querySelector('#status')!.textContent = 'loading';
    (document.querySelector('#editorFrame') as HTMLIFrameElement).src = '/editor?embed=1&deckwise=1&locale=zh-CN';
  });
  await expect(page.locator('#status')).toHaveText('ready', { timeout: 60000 });
  await page.evaluate(
    async ({ data, binding }) => {
      await post('document:open-buffer', {
        fileName: 'water-cycle.pptx',
        buffer: Uint8Array.from(atob(data), (c) => c.charCodeAt(0)).buffer,
        editorBinding: binding,
        trackPageLineage: true,
      });
    },
    { data: toBase64(original), binding },
  );
  await page.waitForFunction(() => !!window.__ooFrames.readyEditor(), { timeout: 60000 });
  const before = await page.evaluate(async (binding) => {
    const call = async (routeOperation: string) =>
      (await post('document:native-teaching', { binding, operation: 'route', routeOperation, page: 3 })).routed.value;
    return { view: await call('open'), png: (await call('narration-visual')).dataUrl };
  }, binding);
  if (!before.view.backgroundRegionRepair) {
    const diagnostic = await page
      .frames()
      .find((f) => f.name() === 'frameEditor')!
      .evaluate(() => {
        const w = window as any,
          api = w.editor,
          logic = api.WordControl.m_oLogicDocument,
          slide = logic.Slides[2],
          c = w.AscCommon;
        return {
          bg: slide.cSld.Bg,
          canEdit: api.canEdit?.(),
          logicEdit: logic.CanEdit?.(),
          master: logic.IsMasterMode?.(),
          slideShow: api.isSlideShow?.(),
          longAction: api.isLongAction?.(),
          bgDuplicate: typeof slide.cSld.Bg?.createDuplicate,
          bgMethod: typeof slide.changeBackground,
          longPoint: c.History._getLongPointIndex?.(),
          globalLock: c.CollaborativeEditing?.Get_GlobalLock(),
          bgLockType: c.changestype_SlideBg,
          drawingLockType: c.changestype_Drawing_Props,
          bgLock: logic.IsSelectionLocked(c.changestype_SlideBg, [slide]),
          drawingLock: logic.IsSelectionLocked(c.changestype_Drawing_Props),
        };
      });
    await writeFile(info.outputPath('capability.json'), JSON.stringify({ diagnostic, view: before.view }, null, 2));
  }
  expect(before.view.backgroundRegionRepair).toBe(true);
  const action = before.view.project.pages[2].actions.find((a: any) =>
    a.visualCues?.some((c: any) => c.effect?.effect === 'fadeIn'),
  );
  expect(action?.audioAssetId).toBeTruthy();
  await expect(
    page.evaluate(
      async ({ binding, view, actionId }) => {
        await post('document:native-teaching', {
          binding,
          operation: 'route',
          routeOperation: 'repair-background-region',
          page: 3,
          args: {
            viewId: view.viewId,
            expectedTeachingRevision: view.project.teachingRevision,
            expectedVisualSha256: '0'.repeat(64),
            actionId,
            description: 'Stale reviewed region',
            bounds: [0.1, 0.25, 0.8, 0.7],
          },
        });
      },
      { binding, view: before.view, actionId: action.id },
    ),
  ).rejects.toThrow('Reviewed region pixels changed');
  const after = await page.evaluate(
    async ({ binding, view, actionId, expectedVisualSha256 }) => {
      const result = (
        await post('document:native-teaching', {
          binding,
          operation: 'route',
          routeOperation: 'repair-background-region',
          page: 3,
          args: {
            viewId: view.viewId,
            expectedTeachingRevision: view.project.teachingRevision,
            expectedVisualSha256,
            actionId,
            description: 'Complete river and runoff teaching region',
            bounds: [0.1, 0.25, 0.8, 0.7],
          },
        })
      ).routed.value;
      const png = (
        await post('document:native-teaching', {
          binding,
          operation: 'route',
          routeOperation: 'narration-visual',
          page: 3,
        })
      ).routed.value.dataUrl;
      return { view: result, png };
    },
    {
      binding,
      view: before.view,
      actionId: action.id,
      expectedVisualSha256: createHash('sha256')
        .update(Buffer.from(before.png.split(',')[1], 'base64'))
        .digest('hex'),
    },
  );
  expect(after.view.backgroundRegionRepair).toBe(false);
  expect(after.view.project.pages).toHaveLength(4);
  const protectedActions = (view: any) =>
    view.project.pages.map((p: any) =>
      p.actions.map((a: any) => {
        const { visualCues: _visualCues, ...rest } = a;
        return rest;
      }),
    );
  expect(protectedActions(after.view)).toEqual(protectedActions(before.view));
  const added = after.view.teachingObjects.filter(
    (o: any) => !before.view.teachingObjects.some((p: any) => p.objectId === o.objectId),
  );
  expect(added).toHaveLength(1);
  const newAction = after.view.project.pages[2].actions.find((a: any) => a.id === action.id);
  expect(
    newAction.visualCues.some(
      (c: any) => c.effect.type === 'element_animation' && c.effect.elementId === added[0].objectId,
    ),
  ).toBe(true);
  expect(
    newAction.visualCues.some(
      (c: any) => c.effect.type === 'spotlight' && c.effect.elementIds.includes(added[0].objectId),
    ),
  ).toBe(true);
  // Compare browser-decoded rendered pixels, not PNG metadata/compression.
  const difference = await page.evaluate(
    async ({ before, after }) => {
      const decode = async (url: string) => {
        const image = new Image();
        image.src = url;
        await image.decode();
        const c = document.createElement('canvas');
        c.width = image.width;
        c.height = image.height;
        const ctx = c.getContext('2d')!;
        ctx.drawImage(image, 0, 0);
        return ctx.getImageData(0, 0, c.width, c.height).data;
      };
      const a = await decode(before),
        b = await decode(after);
      if (a.length !== b.length) throw new Error('Canvas changed');
      return a.reduce((sum, v, i) => sum + Math.abs(v - b[i]), 0) / a.length;
    },
    { before: before.png, after: after.png },
  );
  expect(difference).toBeLessThan(1);
  await writeFile(info.outputPath('before.png'), Buffer.from(before.png.split(',')[1], 'base64'));
  await writeFile(info.outputPath('after.png'), Buffer.from(after.png.split(',')[1], 'base64'));
  const native = page.frames().find((f) => f.name() === 'frameEditor')!;
  await native.evaluate(() => (Reflect.get(window, 'editor') as unknown as { Undo(): void }).Undo());
  const undone = await page.evaluate(
    async (binding) =>
      (await post('document:native-teaching', { binding, operation: 'route', routeOperation: 'open', page: 3 })).routed
        .value,
    binding,
  );
  expect(undone.project.pages).toEqual(before.view.project.pages);
  expect(undone.teachingObjects).toEqual(before.view.teachingObjects);
  await native.evaluate(() => (Reflect.get(window, 'editor') as unknown as { Redo(): void }).Redo());
  const exported = await page.evaluate(async () =>
    Array.from(new Uint8Array(await (await post('document:save')).file.arrayBuffer())),
  );
  await writeFile(info.outputPath('region-repaired.pptx'), new Uint8Array(exported));
  await page.evaluate(
    async ({ bytes, binding }) => {
      await post('document:open-buffer', {
        fileName: 'region-reopened.pptx',
        buffer: new Uint8Array(bytes).buffer,
        editorBinding: binding,
        trackPageLineage: true,
      });
    },
    { bytes: exported, binding },
  );
  await page.waitForFunction(() => !!window.__ooFrames.readyEditor(), { timeout: 60000 });
  const reopened = await page.evaluate(
    async (binding) =>
      (await post('document:native-teaching', { binding, operation: 'route', routeOperation: 'open', page: 3 })).routed
        .value,
    binding,
  );
  expect(reopened.project.pages).toEqual(after.view.project.pages);
  expect(reopened.teachingObjects.some((o: any) => o.objectId === added[0].objectId && o.effectTargetEligible)).toBe(
    true,
  );
  await writeFile(
    info.outputPath('result.json'),
    JSON.stringify(
      { difference, objectId: added[0].objectId, pages: 4, savedAndReopened: true, audioPreserved: true },
      null,
      2,
    ),
  );
});
