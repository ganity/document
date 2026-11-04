import { build } from 'vite';
import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { buildPptx, toBase64, makeStoredZip, zipEntryNames, zipEntryText } from './lib/ooxml';
import { test, expect } from './lib/l0';

test('speech schedule plays chosen native effects concurrently and restores timing and history on exit', async ({
  page,
}) => {
  test.setTimeout(120000);
  const directory = resolve(process.env.E2E_PORT ? `dist-e2e-${process.env.E2E_PORT}` : 'dist', '.deckwise');
  await build({
    configFile: false,
    logLevel: 'silent',
    build: {
      outDir: directory,
      emptyOutDir: false,
      lib: {
        entry: resolve('test/e2e/lib/speech-animation-probe.ts'),
        formats: ['es'],
        fileName: () => 'animation-probe.mjs',
      },
    },
  });
  const bytes = buildPptx('Speech animation');
  const entries = await Promise.all(
    zipEntryNames(bytes).map(async (name) => ({ name, data: (await zipEntryText(bytes, name))! })),
  );
  const source = entries.find((entry) => entry.name === 'ppt/slides/slide1.xml')!;
  const shape = source.data.match(/<p:sp>[\s\S]*?<\/p:sp>/)![0];
  source.data = source.data.replace(
    '</p:spTree>',
    shape.replace('id="2"', 'id="3"').replace('name="Title 1"', 'name="Second"') + '</p:spTree>',
  );
  source.data = source.data.replace(
    '</p:sld>',
    `${await readFile(resolve('test/e2e/fixtures/speech-native-fade.xml'), 'utf8')}</p:sld>`,
  );
  await page.goto('/embed-demo.html');
  await page.evaluate(() => {
    document.querySelector('#status')!.textContent = 'loading';
    (document.querySelector('#editorFrame') as HTMLIFrameElement).src = '/editor?embed=1&deckwise=1&locale=zh-CN';
  });
  await expect(page.locator('#status')).toHaveText('ready', { timeout: 60000 });
  await page.evaluate(
    async (b64) => {
      await post('document:open-buffer', {
        fileName: 'schedule.pptx',
        buffer: Uint8Array.from(atob(b64), (c) => c.charCodeAt(0)).buffer,
      });
    },
    toBase64(makeStoredZip(entries)),
  );
  await page.waitForFunction(() => !!window.__ooFrames.readyEditor(), { timeout: 60000 });
  const outer = page.frames().find((frame) => frame.url().includes('/editor?'))!;
  const result = await outer.evaluate(async () => {
    const url = '/.deckwise/animation-probe.mjs';
    const probe = await import(/* @vite-ignore */ url);
    const frame = document.querySelector('iframe')!,
      doc = frame.contentDocument!;
    const win = frame.contentWindow as Window & { editor?: unknown; AscCommon?: unknown };
    const api = Reflect.get(win, 'editor') as unknown as {
        WordControl: {
          m_oLogicDocument: {
            Slides: Array<{
              cSld: { spTree: Array<{ getUniNvProps(): { cNvPr: { name: string } } }> };
              timing: unknown;
            }>;
          };
        };
      },
      logic = api.WordControl.m_oLogicDocument,
      slide = logic.Slides[0];
    const objects = slide.cSld.spTree.map((shape: { getUniNvProps(): { cNvPr: { name: string } } }, index: number) => ({
      objectId: `object-${index}`,
      status: 'mapped',
      effectTargetEligible: true,
      sourceObjectName: shape.getUniNvProps().cNvPr.name,
    }));
    const saved = slide.timing,
      history = (Reflect.get(win, 'AscCommon') as { History: { Index: number } }).History.Index;
    const objectChoices = probe.teachingObjectChoices(1, objects);
    const thumbnails = objectChoices.map((choice: { thumbnail(): string }) => choice.thumbnail());
    objectChoices[1].locate();
    const selected = Reflect.get(logic, 'GetCurrentController').call(logic).selectedObjects;
    const located = selected.length === 1 && selected[0] === slide.cSld.spTree[1];
    const actions = [
      {
        id: 'speech',
        type: 'speech',
        visualCues: [
          {
            id: 'in',
            atMs: 0,
            endMs: 600,
            trigger: { version: 1, kind: 'start' },
            effect: { type: 'element_animation', elementId: objects[0].objectId, effect: 'fadeIn' },
          },
          {
            id: 'parallel',
            nativeAnimation: objects[1].objectId,
            atMs: 0,
            endMs: 600,
            trigger: { version: 1, kind: 'start' },
            effect: { type: 'element_animation', elementId: objects[1].objectId, effect: 'fadeIn' },
          },
          {
            id: 'out',
            atMs: 1000,
            endMs: 1600,
            trigger: { version: 1, kind: 'end' },
            effect: { type: 'element_animation', elementId: objects[0].objectId, effect: 'fadeOut' },
          },
        ],
      },
    ];
    const schedule = probe.prepareSpeechNativeSchedule(1, actions, objects, probe);
    let ready!: () => void;
    const started = new Promise<void>((resolve) => {
      ready = resolve;
    });
    const errors: string[] = [];
    const show = probe.createTeachingSlideshow({
      doc,
      page: 1,
      schedule,
      timelineStepCount: 3,
      onReady: ready,
      onExit() {},
      onError: (error: Error) => {
        errors.push(error.message);
      },
    });
    let ended = false;
    try {
      await started;
      await Promise.all([show.next(0), show.next(1)]);
      await show.next(2);
      ended = schedule.scheduled.every((entry: { effect: { isAtEnd(): boolean } }) => entry.effect.isAtEnd());
    } finally {
      show.dispose();
    }
    return {
      ended,
      errors,
      thumbnails: thumbnails.every((src: unknown) => typeof src === 'string' && src.startsWith('data:image/')),
      located,
      restored: slide.timing === saved,
      unchangedHistory: (Reflect.get(win, 'AscCommon') as { History: { Index: number } }).History.Index === history,
    };
  });
  expect(result).toEqual({
    ended: true,
    errors: [],
    thumbnails: true,
    located: true,
    restored: true,
    unchangedHistory: true,
  });
});
