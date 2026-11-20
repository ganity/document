import { readFile } from 'node:fs/promises';
import { test, expect } from './lib/l0';
import { toBase64 } from './lib/ooxml';

test('replacing a saved speech motion removes its old native mirror and preserves audio across reopen', async ({
  page,
}, info) => {
  test.setTimeout(180000);
  const fixture = process.env.DECKWISE_NATIVE_MOTION_REPLACEMENT_FIXTURE;
  test.skip(!fixture, 'Requires a saved native lesson with embedded narration and a teaching target');
  const binding = { projectId: 'motion-replacement', documentId: 'lesson', sessionId: 'isolated' };
  await page.goto('/embed-demo.html');
  await page.evaluate(() => {
    document.querySelector('#status')!.textContent = 'loading';
    (document.querySelector('#log') as HTMLElement).style.display = 'none';
    (document.querySelector('#editorFrame') as HTMLIFrameElement).src = '/editor?embed=1&deckwise=1&locale=zh-CN';
  });
  await expect(page.locator('#status')).toHaveText('ready', { timeout: 60000 });
  const open = async (data: string) => {
    await page.evaluate(
      async ({ data, binding }) =>
        post('document:open-buffer', {
          fileName: 'real-motion.pptx',
          buffer: Uint8Array.from(atob(data), (c) => c.charCodeAt(0)).buffer,
          editorBinding: binding,
          trackPageLineage: true,
        }),
      { data, binding },
    );
    await page.waitForFunction(() => !!window.__ooFrames.readyEditor(), { timeout: 60000 });
  };
  await open(toBase64(await readFile(fixture!)));
  const read = () =>
    page.evaluate(
      async (binding) => (await post('document:native-teaching', { binding, operation: 'read' })).teaching,
      binding,
    );
  const effects = () =>
    page
      .frames()
      .find((f) => f.name() === 'frameEditor')!
      .evaluate(() => {
        const slide = (window as any).editor.WordControl.m_oLogicDocument.Slides[0];
        return (slide.timing?.getAllAnimEffects() ?? [])
          .map((effect: any) => ({ preset: effect.cTn.presetID, duration: effect.asc_getDuration(), delay: effect.asc_getDelay() }));
      });
  const replace = (teaching: any, actions: any[]) => page.evaluate(
    async ({ binding, teaching, actions }) => post('document:native-teaching', {
      binding, operation: 'replace-page',
      command: { documentId: teaching.documentId, expectedRevision: teaching.revision, pageId: teaching.pages[0].pageId, actions },
    }), { binding, teaching, actions },
  );
  let before = await read(), motions = await effects();
  if (!motions.length) {
    const actions = structuredClone(before.pages[0].actions);
    const speech = actions.find((a: any) => a.type === 'speech' && a.audioAssetId && a.visualCues?.some((c: any) => c.effect.type === 'spotlight' || c.effect.type === 'highlight'));
    const focus = speech?.visualCues.find((c: any) => c.effect.type === 'spotlight' || c.effect.type === 'highlight');
    const elementId = focus?.effect.elementId ?? focus?.effect.elementIds?.[0];
    expect(elementId).toBeTruthy();
    speech.visualCues.push({ id: 'first-native-cue', atMs: 0, endMs: 600, effect: { type: 'element_animation', effect: 'zoomIn', elementId } });
    await replace(before, actions);
    const seeded = await read();
    expect(seeded.media).toEqual(before.media);
    expect(seeded.pages[0].actions).toEqual(actions);
    before = seeded;
    motions = await effects();
    expect(motions).toHaveLength(1);
  }
  const actions = structuredClone(before.pages[0].actions);
  const originalEffect = actions.some((a: any) => a.visualCues?.some((c: any) => c.effect?.effect === 'zoomIn'))
    ? 'zoomIn'
    : 'pulse';
  const replacementEffect = originalEffect === 'zoomIn' ? 'fadeIn' : 'gentleRotate';
  const originalPreset = originalEffect === 'zoomIn' ? 53 : 26,
    replacementPreset = originalEffect === 'zoomIn' ? 10 : 32;
  const speech = actions.find((a: any) => a.visualCues?.some((c: any) => c.effect?.effect === originalEffect));
  expect(speech?.audioAssetId).toBeTruthy();
  const cue = speech.visualCues.find((c: any) => c.effect?.effect === originalEffect);
  cue.effect.effect = replacementEffect;
  delete cue.trigger;
  cue.atMs += 300;
  cue.endMs += 300;
  await replace(before, actions);
  const changed = await read(),
    updated = await effects();
  expect(updated).toHaveLength(motions.length);
  expect(updated.some((e: any) => e.preset === replacementPreset && e.delay === 300)).toBe(true);
  expect(updated.filter((e: any) => e.preset === originalPreset)).toHaveLength(
    motions.filter((e: any) => e.preset === originalPreset).length - 1,
  );
  expect(updated.filter((e: any) => e.preset === replacementPreset)).toHaveLength(
    motions.filter((e: any) => e.preset === replacementPreset).length + 1,
  );
  expect(changed.pages[0].actions).toEqual(actions);
  expect(changed.media).toEqual(before.media);
  expect(changed.pages.slice(1)).toEqual(before.pages.slice(1));
  const native = page.frames().find((f) => f.name() === 'frameEditor')!;
  await native.evaluate(() => (window as any).editor.Undo());
  expect(await effects()).toEqual(motions);
  expect((await read()).pages).toEqual(before.pages);
  await native.evaluate(() => (window as any).editor.Redo());
  expect(await effects()).toEqual(updated);
  expect((await read()).pages).toEqual(changed.pages);
  const bytes = await page.evaluate(async () =>
    Array.from(new Uint8Array(await (await post('document:save')).file.arrayBuffer())),
  );
  await info.attach('repaired-real-pptx', {
    body: Buffer.from(bytes),
    contentType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
  });
  await open(toBase64(new Uint8Array(bytes)));
  expect((await read()).pages).toEqual(changed.pages);
  expect(await effects()).toEqual(updated);
  const current = await read();
  const removed = structuredClone(current.pages[0].actions);
  removed.find((a: any) => a.id === speech.id).visualCues = speech.visualCues.filter((c: any) => c.id !== cue.id);
  await replace(current, removed);
  expect(await effects()).toHaveLength(updated.length - 1);
  expect((await read()).media).toEqual(before.media);
});
