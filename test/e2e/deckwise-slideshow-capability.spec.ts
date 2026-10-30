import { buildPptx, toBase64 } from './lib/ooxml';
import { expect, test } from './lib/l0';

test('native slideshow exposes live animation completion, pause and slide geometry', async ({ page }, testInfo) => {
  test.setTimeout(120_000);
  await page.goto('/embed-demo.html');
  await page.evaluate(() => {
    document.querySelector('#status')!.textContent = 'loading';
    (document.querySelector('#editorFrame') as HTMLIFrameElement).src = '/editor?embed=1&deckwise=1';
  });
  await expect(page.locator('#status')).toHaveText('ready', { timeout: 60_000 });
  await page.evaluate(
    async (b64) => {
      const bytes = Uint8Array.from(atob(b64), (c) => c.charCodeAt(0));
      await post('document:open-buffer', { fileName: 'slideshow.pptx', buffer: bytes.buffer });
    },
    toBase64(buildPptx('Native slideshow capability')),
  );
  await page.waitForFunction(() => !!window.__ooFrames.readyEditor(), { timeout: 60_000 });
  const native = page.frames().find((frame) => frame.name() === 'frameEditor')!;
  const initial = await native.evaluate(() => {
    const api = (window as any).editor,
      logic = api.WordControl.m_oLogicDocument;
    const controller = logic.GetCurrentController(),
      shape = logic.Slides[0].cSld.spTree[0];
    controller.resetSelection();
    controller.selectObject(shape, 0);
    api.asc_AddAnimation(1, 10, 0, null, false, false);
    api.asc_AddAnimation(1, 10, 0, null, false, false);
    const effects = logic.Slides[0].timing.getAllAnimEffects();
    const host = document.createElement('div');
    host.id = 'deckwise-slideshow-probe';
    host.style.cssText = 'position:fixed;inset:0;z-index:10000;background:black';
    document.body.append(host);
    const events: unknown[] = [];
    (window as any).slideshowProbe = { events, history: (window as any).AscCommon.History.Index };
    api.asc_registerCallback('asc_onDemonstrationSlideChanged', (page: number) => events.push({ type: 'page', page }));
    api.asc_registerCallback('asc_onEndDemonstration', () => events.push({ type: 'end' }));
    api.StartDemonstration(host.id, 0);
    const manager = api.WordControl.DemonstrationManager,
      player = manager.GetCurrentAnimPlayer();
    return {
      effects: effects.length,
      mode: manager.Mode,
      page: manager.SlideNum,
      complete: player.isMainSequenceFinished(),
      rect: manager.Transition.Rect,
      playerMethods: Object.getOwnPropertyNames(Object.getPrototypeOf(player)).filter((name) =>
        /play|pause|finish|next|time/i.test(name),
      ),
    };
  });
  expect(initial.effects).toBe(2);
  expect(initial.mode).toBe(true);
  expect(initial.complete).toBe(false);
  await native.evaluate(() => (window as any).editor.DemonstrationNextSlide());
  await expect
    .poll(() =>
      native.evaluate(() =>
        (window as any).editor.WordControl.DemonstrationManager.GetCurrentAnimPlayer().getElapsedTicks(),
      ),
    )
    .toBeGreaterThan(0);
  const pausedTick = await native.evaluate(() => {
    const api = (window as any).editor;
    api.DemonstrationPause();
    return api.WordControl.DemonstrationManager.GetCurrentAnimPlayer().getElapsedTicks();
  });
  await page.waitForTimeout(250);
  expect(
    await native.evaluate(() =>
      (window as any).editor.WordControl.DemonstrationManager.GetCurrentAnimPlayer().getElapsedTicks(),
    ),
  ).toBe(pausedTick);
  await native.evaluate(() => (window as any).editor.DemonstrationPlay());
  await expect
    .poll(() =>
      native.evaluate(() =>
        (window as any).editor.WordControl.m_oLogicDocument.Slides[0].timing.getAllAnimEffects()[0].isAtEnd(),
      ),
    )
    .toBe(true);
  const first = await native.evaluate(() => {
    const api = (window as any).editor,
      manager = api.WordControl.DemonstrationManager;
    const effects = api.WordControl.m_oLogicDocument.Slides[0].timing.getAllAnimEffects();
    return {
      firstFinished: effects[0].isAtEnd(),
      secondIdle: effects[1].isIdle(),
      complete: manager.IsMainSeqFinished(0),
      page: manager.SlideNum,
    };
  });
  expect(first).toEqual({ firstFinished: true, secondIdle: true, complete: false, page: 0 });
  expect(
    await native.evaluate(() => (window as any).editor.WordControl.DemonstrationManager.OnNextSlideAnimPlayer()),
  ).toBe(true);
  await expect
    .poll(() => native.evaluate(() => (window as any).editor.WordControl.DemonstrationManager.IsMainSeqFinished(0)))
    .toBe(true);
  const finished = await native.evaluate(() => {
    const api = (window as any).editor,
      manager = api.WordControl.DemonstrationManager;
    api.DemonstrationPause();
    const paused = { playMode: manager.IsPlayMode, animation: manager.IsPausedAnimation(0), page: manager.SlideNum };
    const rect = { ...manager.Transition.Rect };
    const canvas = manager.Canvas.getBoundingClientRect();
    api.EndDemonstration();
    return {
      paused,
      rect,
      canvas: { x: canvas.x, y: canvas.y, width: canvas.width, height: canvas.height },
      ended: !manager.Mode,
      events: (window as any).slideshowProbe.events,
      historyUnchanged: (window as any).AscCommon.History.Index === (window as any).slideshowProbe.history,
    };
  });
  expect(finished.paused.playMode).toBe(false);
  expect(finished.paused.page).toBe(0);
  expect(finished.rect.w).toBeGreaterThan(0);
  expect(finished.rect.h).toBeGreaterThan(0);
  expect(finished.ended).toBe(true);
  expect(finished.events).toContainEqual({ type: 'page', page: 0 });
  expect(finished.events).toContainEqual({ type: 'end' });
  expect(finished.historyUnchanged).toBe(true);
  await testInfo.attach('native-slideshow-capabilities', {
    body: JSON.stringify({ initial, first, finished }, null, 2),
    contentType: 'application/json',
  });
});
