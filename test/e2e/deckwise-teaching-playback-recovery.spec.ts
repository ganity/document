import { readFile } from 'node:fs/promises';
import { test, expect } from './lib/l0';
import { toBase64 } from './lib/ooxml';

test('saved video teaching finishes without SDK errors and narration follows navigation and sidebar reopening', async ({ page }) => {
  const fixture = process.env.DECKWISE_TEACHING_PLAYBACK_PPTX;
  test.skip(!fixture, 'Requires a real native teaching PPTX with adopted video and speech on page 3');
  test.setTimeout(240_000);
  await page.addInitScript(() => {
    const w = window as any;
    const listeners = new Set<(command: unknown) => void>();
    w.playbackEvents = [];
    w.playbackCommand = (type: string, payload: unknown = {}, id?: string) => {
      for (const listener of listeners) listener({ type, payload, id });
    };
    w.deckwiseEditorHost = {
      onCommand(listener: (command: unknown) => void) { listeners.add(listener); return () => listeners.delete(listener); },
      sendEvent(event: any) {
        w.playbackEvents.push(event);
        const message = event.message;
        if (message?.type !== 'document:teaching-request') return;
        if (!['tts:models', 'tts:list', 'media:list'].includes(message.payload.operation))
          throw new Error('Unexpected host operation: ' + message.payload.operation);
        w.playbackCommand('document:teaching-response', { result: [] }, message.id);
      },
    };
  });
  await page.goto('/module.html?deckwise=1&locale=zh-CN');
  await page.waitForFunction(() => (window as any).playbackEvents.some((e: any) => e.message?.type === 'document:ready'));
  const binding = { projectId: 'playback-regression', documentId: 'saved-native', sessionId: 'isolated' };
  await page.evaluate(({ data, binding }) => (window as any).playbackCommand('document:open-buffer', {
    fileName: 'saved-native.pptx', buffer: Uint8Array.from(atob(data), c => c.charCodeAt(0)).buffer,
    editorBinding: binding, trackPageLineage: true,
  }), { data: toBase64(await readFile(fixture!)), binding });
  await page.waitForFunction(() => !!window.__ooFrames.readyEditor(), undefined, { timeout: 60_000 });
  await page.evaluate(binding => (window as any).playbackCommand('document:teaching-enable', { binding, nativeTeaching: true }), binding);
  const frame = page.frames().find(frame => frame.name() === 'frameEditor')!;
  const panel = frame.locator('.deckwise-teaching-sidebar');
  await expect(panel).toBeVisible();
  const ids = await frame.evaluate(() => {
    const w = window as any;
    const parts = new w.AscBuilder.ApiCustomXmlParts(w.editor.WordControl.m_oLogicDocument).GetByNamespace('urn:deckwise:native-teaching:1');
    return JSON.parse(new DOMParser().parseFromString(parts[0].GetXml(), 'text/xml').documentElement.textContent!).pages
      .map((p: any) => p.actions.map((a: any) => a.id));
  });
  const navigate = async (index: number) => {
    await frame.evaluate(index => (window as any).editor.goToPage(index), index);
    await expect.poll(() => panel.locator('.teaching-flow > article[data-action-id]').evaluateAll(nodes => nodes.map(node => (node as HTMLElement).dataset.actionId))).toEqual(ids[index]);
  };
  await navigate(2);
  await panel.locator('[data-preview-command="show"]').click();
  await expect(frame.locator('[data-teaching-slideshow]')).toBeVisible({ timeout: 30_000 });
  await expect(frame.locator('[data-teaching-slideshow]')).toHaveCount(0, { timeout: 120_000 });
  await expect(panel.locator('[data-status]')).toHaveAttribute('role', 'status');
  expect(await frame.evaluate(() => (window as any).editor.isViewMode)).toBe(false);
  for (const index of [0, 1, 2, 0]) await navigate(index);
  await page.evaluate(() => (window as any).playbackCommand('document:configure-ui', { teachingOpen: false }));
  await frame.evaluate(() => (window as any).editor.goToPage(2));
  await page.evaluate(() => (window as any).playbackCommand('document:configure-ui', { teachingOpen: true }));
  await navigate(2);
});
