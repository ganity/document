import { readFile } from 'node:fs/promises';
import { test, expect } from './lib/l0';
import { toBase64 } from './lib/ooxml';

test('replacing a saved speech motion removes its old native mirror and preserves audio across reopen', async ({ page }, info) => {
  test.setTimeout(180000);
  const fixture = process.env.DECKWISE_NATIVE_MOTION_REPLACEMENT_FIXTURE;
  test.skip(!fixture, 'Requires the preserved real lesson with a saved zoomIn or pulse speech cue and embedded audio');
  const binding = { projectId: 'motion-replacement', documentId: 'lesson', sessionId: 'isolated' };
  await page.goto('/embed-demo.html');
  await page.evaluate(() => {
    document.querySelector('#status')!.textContent = 'loading';
    (document.querySelector('#log') as HTMLElement).style.display = 'none';
    (document.querySelector('#editorFrame') as HTMLIFrameElement).src = '/editor?embed=1&deckwise=1&locale=zh-CN';
  });
  await expect(page.locator('#status')).toHaveText('ready', { timeout: 60000 });
  const open = async (data: string) => {
    await page.evaluate(async ({ data, binding }) => post('document:open-buffer', {
      fileName: 'real-motion.pptx', buffer: Uint8Array.from(atob(data), c => c.charCodeAt(0)).buffer,
      editorBinding: binding, trackPageLineage: true,
    }), { data, binding });
    await page.waitForFunction(() => !!window.__ooFrames.readyEditor(), { timeout: 60000 });
  };
  await open(toBase64(await readFile(fixture!)));
  const read = () => page.evaluate(async binding => (await post('document:native-teaching', { binding, operation: 'read' })).teaching, binding);
  const effects = () => page.frames().find(f => f.name() === 'frameEditor')!.evaluate(() => {
    const slide = (window as any).editor.WordControl.m_oLogicDocument.Slides[0];
    return slide.timing.getAllAnimEffects().map((effect: any) => ({ preset: effect.cTn.presetID, duration: effect.asc_getDuration() }));
  });
  const before = await read(), motions = await effects();
  const actions = structuredClone(before.pages[0].actions);
  const originalEffect = actions.some((a: any) => a.visualCues?.some((c: any) => c.effect?.effect === 'zoomIn')) ? 'zoomIn' : 'pulse';
  const replacementEffect = originalEffect === 'zoomIn' ? 'fadeIn' : 'gentleRotate';
  const originalPreset = originalEffect === 'zoomIn' ? 53 : 26, replacementPreset = originalEffect === 'zoomIn' ? 10 : 32;
  const speech = actions.find((a: any) => a.visualCues?.some((c: any) => c.effect?.effect === originalEffect));
  expect(speech?.audioAssetId).toBeTruthy();
  const cue = speech.visualCues.find((c: any) => c.effect?.effect === originalEffect);
  cue.effect.effect = replacementEffect;
  await page.evaluate(async ({ binding, before, actions }) => post('document:native-teaching', { binding, operation: 'replace-page', command: {
    documentId: before.documentId, expectedRevision: before.revision, pageId: before.pages[0].pageId, actions,
  } }), { binding, before, actions });
  const changed = await read(), updated = await effects();
  expect(updated).toHaveLength(motions.length);
  expect(updated.filter((e: any) => e.preset === originalPreset)).toHaveLength(motions.filter((e: any) => e.preset === originalPreset).length - 1);
  expect(updated.filter((e: any) => e.preset === replacementPreset)).toHaveLength(motions.filter((e: any) => e.preset === replacementPreset).length + 1);
  expect(changed.pages[0].actions).toEqual(actions);
  expect(changed.media).toEqual(before.media);
  expect(changed.pages.slice(1)).toEqual(before.pages.slice(1));
  const native = page.frames().find(f => f.name() === 'frameEditor')!;
  await native.evaluate(() => (window as any).editor.Undo());
  expect(await effects()).toEqual(motions);
  expect((await read()).pages).toEqual(before.pages);
  await native.evaluate(() => (window as any).editor.Redo());
  expect(await effects()).toEqual(updated);
  expect((await read()).pages).toEqual(changed.pages);
  const bytes = await page.evaluate(async () => Array.from(new Uint8Array(await (await post('document:save')).file.arrayBuffer())));
  await info.attach('repaired-real-pptx', { body: Buffer.from(bytes), contentType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation' });
  await open(toBase64(new Uint8Array(bytes)));
  expect((await read()).pages).toEqual(changed.pages);
  expect(await effects()).toEqual(updated);
  const current = await read();
  const removed = structuredClone(current.pages[0].actions);
  removed.find((a: any) => a.id === speech.id).visualCues = speech.visualCues.filter((c: any) => c.id !== cue.id);
  await page.evaluate(async ({ binding, current, removed }) => post('document:native-teaching', { binding, operation: 'replace-page', command: {
    documentId: current.documentId, expectedRevision: current.revision, pageId: current.pages[0].pageId, actions: removed,
  } }), { binding, current, removed });
  expect(await effects()).toHaveLength(updated.length - 1);
  expect((await read()).media).toEqual(before.media);
});
