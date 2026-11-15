import { readFile, writeFile } from 'node:fs/promises';
import { buildPptx, toBase64, makeStoredZip, zipEntryNames, zipEntryText } from './lib/ooxml';
import { expect, test } from './lib/l0';

test('generated speech motions survive native PPTX save and reopen as editable grouped animations', async ({
  page,
}) => {
  test.setTimeout(240_000);
  const probe = process.env.DECKWISE_MOTION_PROBE;
  test.skip(!probe, 'Requires the isolated motion probe built from the paired Desktop source');
  const probeSource = await readFile(probe!, 'utf8');
  await page.route('**/.deckwise/motion-probe.mjs', (route) =>
    route.fulfill({ body: probeSource, contentType: 'text/javascript' }),
  );
  const original = buildPptx('Hello friends');
  const entries = await Promise.all(
    zipEntryNames(original).map(async (name) => ({ name, data: (await zipEntryText(original, name))! })),
  );
  const slide = entries.find((e) => e.name === 'ppt/slides/slide1.xml')!;
  slide.data = slide.data.replace('<p:ph type="title"/>', '');
  const shape = slide.data.match(/<p:sp>[\s\S]*?<\/p:sp>/)![0];
  slide.data = slide.data.replace(
    '</p:spTree>',
    shape.replace('id="2"', 'id="3"').replace('name="Title 1"', 'name="Friend 2"') + '</p:spTree>',
  );
  await page.goto('/embed-demo.html');
  await page.evaluate(() => {
    document.querySelector('#status')!.textContent = 'loading';
    (document.querySelector('#editorFrame') as HTMLIFrameElement).src = '/editor?embed=1&deckwise=1&locale=zh-CN';
  });
  await expect(page.locator('#status')).toHaveText('ready', { timeout: 60_000 });
  const open = async (b64: string) => {
    await page.evaluate(async (b64) => {
      await post('document:open-buffer', {
        fileName: 'motions.pptx',
        buffer: Uint8Array.from(atob(b64), (c) => c.charCodeAt(0)).buffer,
      });
    }, b64);
    await page.waitForFunction(() => !!window.__ooFrames.readyEditor(), { timeout: 60_000 });
  };
  await open(toBase64(makeStoredZip(entries)));
  const outer = page.frames().find((f) => f.url().includes('/editor?'))!;
  const read = async () =>
    page
      .frames()
      .find((f) => f.name() === 'frameEditor')!
      .evaluate(() => {
        const api = (window as any).editor;
        const slide = api.WordControl.m_oLogicDocument.Slides[0];
        return slide.timing.getAllAnimEffects().map((effect: any) => ({
          target: effect.getObjectId(),
          kind: effect.cTn.presetClass,
          preset: effect.cTn.presetID,
          subtype: effect.cTn.presetSubtype,
          start: effect.cTn.nodeType,
          duration: effect.asc_getDuration(),
        }));
      });
  const result = await outer.evaluate(async () => {
    const probeUrl = '/.deckwise/motion-probe.mjs';
    const probe = await import(/* @vite-ignore */ probeUrl);
    const api = (document.querySelector('iframe')!.contentWindow as any).editor;
    const shapes = api.WordControl.m_oLogicDocument.Slides[0].cSld.spTree;
    const objects = shapes.map((s: any, i: number) => ({
      objectId: `object-${i}`,
      status: 'mapped',
      effectTargetEligible: true,
      sourceObjectName: s.getUniNvProps().cNvPr.name,
    }));
    const effects = ['fadeIn', 'zoomIn', 'slideInLeft', 'slideInRight', 'pulse', 'gentleRotate', 'fadeOut'];
    const actions = effects.map((effect, i) => ({
      id: `speech-${i}`,
      type: 'speech',
      visualCues: objects.map((o: any, j: number) => ({
        id: `cue-${i}-${j}`,
        atMs: 0,
        endMs: 600,
        effect: { type: 'element_animation', elementId: o.objectId, effect },
      })),
    }));
    try {
      return probe.integrateGeneratedNativeMotions(1, actions, objects, probe);
    } catch (error) {
      const e = error as Error;
      throw new Error(
        `${e.message}: ${String(e.cause)}; shapes=${JSON.stringify(shapes.map((s: any) => ({ id: s.Id, name: s.getUniNvProps().cNvPr.name })))}; ` +
          JSON.stringify(
            api.WordControl.m_oLogicDocument.Slides[0].timing?.getAllAnimEffects().map((a: any) => ({
              class: a.cTn.presetClass,
              id: a.cTn.presetID,
              subtype: a.cTn.presetSubtype,
              node: a.cTn.nodeType,
              target: a.getObjectId(),
              duration: a.asc_getDuration(),
              delay: a.asc_getDelay(),
            })),
          ),
      );
    }
  });
  expect(result).toBe(true);
  const before = await read();
  expect(before).toHaveLength(14);
  expect(before.map((e: any) => e.start)).toEqual(Array.from({ length: 7 }, () => [2, 7]).flat());
  expect(before.every((e: any) => e.duration === 600)).toBe(true);
  const saved = await page.evaluate(async () => {
    const result = await post('document:save', {});
    const bytes = new Uint8Array(await result.file.arrayBuffer());
    return btoa(Array.from(bytes, (b) => String.fromCharCode(b)).join(''));
  });
  const bytes = Uint8Array.from(Buffer.from(saved, 'base64'));
  const xml = await zipEntryText(bytes, 'ppt/slides/slide1.xml');
  await writeFile('/tmp/wise-generated-motions-slide.xml', xml!);
  expect(xml).toContain('presetID="10"');
  expect(xml).toContain('presetID="26"');
  await open(saved);
  const reopened = await read();
  expect(reopened.map(({ target: _target, ...rest }: any) => rest)).toEqual(
    before.map(({ target: _target, ...rest }: any) => rest),
  );
  const playback = await outer.evaluate(async () => {
    const probeUrl = '/.deckwise/motion-probe.mjs';
    const probe = await import(/* @vite-ignore */ probeUrl);
    const doc = document.querySelector('iframe')!.contentDocument!;
    const errors: string[] = [];
    let signalReady: () => void = () => {};
    const ready = new Promise<void>((resolve) => {
      signalReady = resolve;
    });
    const show = probe.createTeachingSlideshow({
      doc,
      page: 1,
      timelineStepCount: 7,
      generatedMotions: true,
      onReady: signalReady,
      onExit: () => {},
      onError: (e: Error) => errors.push(e.message),
    });
    try {
      await ready;
      for (let step = 0; step < 7; step++) await show.next(step);
      const api = (document.querySelector('iframe')!.contentWindow as any).editor;
      const ended = api.WordControl.m_oLogicDocument.Slides[0].timing
        .getAllAnimEffects()
        .every((e: any) => e.isAtEnd());
      return { errors, ended };
    } finally {
      show.dispose();
    }
  });
  expect(playback).toEqual({ errors: [], ended: true });
  // Native user edits survive teaching binding and replay; the generated cue is not a template lock.
  const editedSteps = await outer.evaluate(async () => {
    const url = '/.deckwise/motion-probe.mjs';
    await import(/* @vite-ignore */ url);
    const api = (document.querySelector('iframe')!.contentWindow as any).editor;
    const logic = api.WordControl.m_oLogicDocument,
      slide = logic.Slides[0];
    const controller = logic.GetCurrentController();
    controller.resetSelection();
    controller.selectObject(slide.cSld.spTree[0], 0);
    slide.timing.resetSelection();
    api.asc_AddAnimation(1, 1, 0, null, false, false);
    const effect = slide.timing.getAllAnimEffects().at(-1);
    slide.timing.resetSelection();
    effect.select();
    const properties = effect.createDuplicate();
    properties.merge(effect);
    properties.asc_putDuration(1200);
    api.asc_SetAnimationProperties(properties);
    return slide.timing.getAllAnimEffects().length;
  });
  expect(editedSteps).toBe(15);
  const editedSaved = await page.evaluate(async () => {
    const result = await post('document:save', {});
    const bytes = new Uint8Array(await result.file.arrayBuffer());
    return btoa(Array.from(bytes, (b) => String.fromCharCode(b)).join(''));
  });
  await open(editedSaved);
  const editedPlayback = await outer.evaluate(async () => {
    const url = '/.deckwise/motion-probe.mjs';
    const probe = await import(/* @vite-ignore */ url);
    const frame = document.querySelector('iframe')!,
      doc = frame.contentDocument!;
    const slide = (frame.contentWindow as any).editor.WordControl.m_oLogicDocument.Slides[0];
    const objects = slide.cSld.spTree.map((shape: any, i: number) => ({
      objectId: `object-${i}`,
      status: 'mapped',
      effectTargetEligible: true,
      sourceObjectName: shape.getUniNvProps().cNvPr.name,
    }));
    const actions = ['fadeIn', 'zoomIn', 'slideInLeft', 'slideInRight', 'pulse', 'gentleRotate', 'fadeOut'].map(
      (effect, i) => ({
        id: `speech-${i}`,
        type: 'speech',
        visualCues: objects.map((o: any, j: number) => ({
          id: `cue-${i}-${j}`,
          atMs: 0,
          endMs: 600,
          effect: { type: 'element_animation', elementId: o.objectId, effect },
        })),
      }),
    );
    const bindings = probe.nativeSpeechMotionSteps(1, actions, objects, probe);
    const before = slide.timing.getAllAnimEffects();
    const errors: string[] = [];
    let ready!: () => void;
    const started = new Promise<void>((resolve) => {
      ready = resolve;
    });
    const show = probe.createTeachingSlideshow({
      doc,
      page: 1,
      timelineStepCount: bindings.length,
      generatedMotions: true,
      onReady: ready,
      onExit() {},
      onError: (error: Error) => errors.push(error.message),
    });
    try {
      await started;
      for (const binding of bindings) await show.next(binding.stepIndex);
      return {
        bindings,
        errors,
        ended: before.every((effect: any) => effect.isAtEnd()),
        lastDuration: before.at(-1).asc_getDuration(),
        unchanged: slide.timing.getAllAnimEffects().every((effect: any, i: number) => before[i] === effect),
      };
    } finally {
      show.dispose();
    }
  });
  expect(editedPlayback.bindings).toHaveLength(8);
  expect(editedPlayback.errors).toEqual([]);
  expect(editedPlayback.ended).toBe(true);
  expect(editedPlayback.lastDuration).toBe(1200);
  expect(editedPlayback.unchanged).toBe(true);
});
