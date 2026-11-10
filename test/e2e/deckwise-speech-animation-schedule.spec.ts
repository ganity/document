import { build } from 'vite';
import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { pathToFileURL } from 'node:url';
import { buildPptx, toBase64, makeStoredZip, zipEntryNames, zipEntryText } from './lib/ooxml';
import { test, expect } from './lib/l0';

for (const mode of ['legacy', 'generated', 'linked', 'emphasis'] as const)
  test(`speech schedule (${mode}) retains group animations, autoplays unlinked effects and restores timing`, async ({
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
    const second = shape
      .replace('id="2"', 'id="3"')
      .replace('name="Title 1"', 'name="Second"')
      .replace('<p:ph type="title"/>', '');
    const group =
      '<p:grpSp><p:nvGrpSpPr><p:cNvPr id="5" name="Second group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
      '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="9144000" cy="6858000"/>' +
      '<a:chOff x="0" y="0"/><a:chExt cx="9144000" cy="6858000"/></a:xfrm></p:grpSpPr>' +
      second +
      '</p:grpSp>';
    source.data = source.data.replace(
      '</p:spTree>',
      group +
        shape
          .replace('id="2"', 'id="4"')
          .replace('name="Title 1"', 'name="Unlinked"')
          .replace('<p:ph type="title"/>', '') +
        '</p:spTree>',
    );
    source.data = source.data.replace(
      '</p:sld>',
      `${(await readFile(resolve('test/e2e/fixtures/speech-native-fade.xml'), 'utf8'))
        .replace('spid="3"', 'spid="5"')
        .replace(
          /\n              <\/p:childTnLst>\n            <\/p:cTn>/,
          `\n<p:par><p:cTn id="6" fill="hold"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst><p:childTnLst>
      <p:par><p:cTn id="7" presetID="10" presetClass="entr" presetSubtype="0" fill="hold" nodeType="clickEffect">
      <p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst><p:animEffect transition="in" filter="fade"><p:cBhvr>
      <p:cTn id="8" dur="500"/><p:tgtEl><p:spTgt spid="4"/></p:tgtEl></p:cBhvr></p:animEffect></p:childTnLst>
      </p:cTn></p:par></p:childTnLst></p:cTn></p:par>
              </p:childTnLst>\n            </p:cTn>`,
        )}</p:sld>`,
    );
    const authoringPath = process.env.DECKWISE_MOTION_AUTHORING;
    let pptx: Uint8Array;
    if (authoringPath) {
      const { syncSpeechAnimations } = await import(pathToFileURL(authoringPath).href);
      const persisted = await syncSpeechAnimations({ source: makeStoredZip(entries), pages: [{
      objects: [{ id: 'title', objectName: 'Title 1' }, { id: 'group', objectName: 'Second group' }],
      actions: [{ type: 'speech', visualCues: [
        { atMs: 0, endMs: 600, effect: { type: 'element_animation', elementId: 'title', effect: 'fadeIn' } },
        { atMs: 0, endMs: 500, effect: { type: 'element_animation', elementId: 'group', effect: 'fadeIn' } },
        { atMs: 1000, endMs: 1600, effect: { type: 'element_animation', elementId: 'title', effect: mode === 'emphasis' ? 'pulse' : 'fadeOut' } },
      ] }],
      }] });
      expect(persisted.changedPages).toEqual([true]);
      pptx = persisted.pptx;
    } else {
      // Captured native compiler output keeps the standalone WiseOffice suite provider/repository independent.
      const fixtures = JSON.parse(await readFile(resolve('test/e2e/fixtures/speech-native-persisted.json'), 'utf8'));
      const timing = fixtures[mode === 'emphasis' ? 'pulse' : 'fadeOut'] as string;
      const marker = 'nodeType="mainSeq"><p:childTnLst>';
      const effects = timing.slice(timing.indexOf(marker) + marker.length,
        timing.lastIndexOf('</p:childTnLst></p:cTn><p:prevCondLst>')).replace(/\bid="(\d+)"/g, (_, id) => `id="${Number(id) + 100}"`);
      source.data = source.data.replace(/\n              <\/p:childTnLst>\n            <\/p:cTn>/,
        `${effects}\n              </p:childTnLst>\n            </p:cTn>`);
      pptx = makeStoredZip(entries);
    }
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
      toBase64(pptx),
    );
    await page.waitForFunction(() => !!window.__ooFrames.readyEditor(), { timeout: 60000 });
    const outer = page.frames().find((frame) => frame.url().includes('/editor?'))!;
    const result = await outer.evaluate(async (mode) => {
      const url = '/.deckwise/animation-probe.mjs';
      const probe = await import(/* @vite-ignore */ url);
      const frame = document.querySelector('iframe')!,
        doc = frame.contentDocument!;
      const win = frame.contentWindow as Window & { editor?: unknown; AscCommon?: unknown };
      const waitReady = async (ready: Promise<void>, details: () => unknown) => {
        let timer = 0;
        try {
          await Promise.race([
            ready,
            new Promise((_, reject) => {
              timer = window.setTimeout(
                () => reject(new Error(`Native playback stalled: ${JSON.stringify(details())}`)),
                15000,
              );
            }),
          ]);
        } finally {
          clearTimeout(timer);
        }
      };
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
      const objects = slide.cSld.spTree
        .flatMap((shape) => Reflect.get(shape, 'spTree') ?? [shape])
        .map((shape: { getUniNvProps(): { cNvPr: { name: string } } }, index: number) => ({
          objectId: `object-${index}`,
          status: 'mapped',
          effectTargetEligible: true,
          sourceObjectName: shape.getUniNvProps().cNvPr.name,
        }));
      const saved = slide.timing,
        history = (Reflect.get(win, 'AscCommon') as { History: { Index: number } }).History.Index;
      const objectChoices = probe.teachingObjectChoices(1, objects);
      const thumbnails = objectChoices.map((choice: { thumbnail(): string }) => choice.thumbnail());
      objectChoices[0].locate();
      const selected = Reflect.get(logic, 'GetCurrentController').call(logic).selectedObjects;
      const located = selected.length === 1 && selected[0] === slide.cSld.spTree[0];
      const control = api.WordControl as typeof api.WordControl & { m_oOverlay: { HtmlElement: HTMLCanvasElement } };
      const overlay = control.m_oOverlay.HtmlElement;
      if (!overlay.getContext('2d')!.getImageData(0, 0, overlay.width, overlay.height).data.some((value, i) => i % 4 === 3 && value > 0))
        throw new Error('Locate did not draw a selection outline');
      objectChoices[1].locate();
      const controller = Reflect.get(logic, 'GetCurrentController').call(logic);
      const group = slide.cSld.spTree[1] as typeof slide.cSld.spTree[1] & { selectedObjects: unknown[]; spTree: unknown[] };
      if (controller.selection.groupSelection !== group || group.selectedObjects[0] !== group.spTree[0])
        throw new Error('Locate did not select the exact object inside its group');
      if (mode === 'legacy') {
        let ready!: () => void;
        const started = new Promise<void>((resolve) => {
          ready = resolve;
        });
        const automatic = probe.createTeachingSlideshow({
          doc,
          page: 1,
          onReady: ready,
          onExit() {},
          onError(error: Error) {
            throw error;
          },
        });
        try {
          await waitReady(started, () => ({
            effects: Reflect.get(saved as object, 'getAllAnimEffects')
              .call(saved)
              .map((effect: { isIdle(): boolean; isAtEnd(): boolean }) => ({
                idle: effect.isIdle(),
                ended: effect.isAtEnd(),
              })),
            text: doc.querySelector('[data-teaching-slideshow]')?.textContent,
          }));
          const deadline = performance.now() + 15000;
          while (!Reflect.get(saved as object, 'getAllAnimEffects')
            .call(saved).every((effect: { isAtEnd(): boolean }) => effect.isAtEnd()) && performance.now() < deadline)
            await new Promise((resolve) => setTimeout(resolve, 20));
          if (
            !Reflect.get(saved as object, 'getAllAnimEffects')
              .call(saved)
              .every((effect: { isAtEnd(): boolean }) => effect.isAtEnd())
          )
            throw new Error('Unlinked native animations required a manual click');
        } finally {
          automatic.dispose();
        }
      }
      const actions = [
        {
          id: 'speech',
          type: 'speech',
          visualCues: [
            {
              id: 'in',
              atMs: 0,
              endMs: 600,
              ...(mode === 'legacy' ? {} : { trigger: { version: 1, kind: 'start' } }),
              effect: { type: 'element_animation', elementId: objects[0].objectId, effect: 'fadeIn' },
            },
            {
              id: 'parallel',
              atMs: 0,
              endMs: 500,
              ...(mode === 'linked' ? { nativeAnimation: objects[1].objectId } : {}),
              ...(mode === 'legacy' ? {} : { trigger: { version: 1, kind: 'start' } }),
              effect: { type: 'element_animation', elementId: objects[1].objectId, effect: 'fadeIn' },
            },
            {
              id: 'out',
              atMs: 1000,
              endMs: 1600,
              ...(mode === 'legacy' ? {} : { trigger: { version: 1, kind: 'end' } }),
              effect: { type: 'element_animation', elementId: objects[0].objectId, effect: mode === 'emphasis' ? 'pulse' : 'fadeOut' },
            },
          ],
        },
      ];
      const schedule = probe.prepareSpeechNativeSchedule(1, actions, objects, probe);
      if (!schedule) throw new Error('Generated speech cues were not scheduled');
      if (schedule.unboundCount !== 1)
        throw new Error(`Expected one unlinked page-entry effect, got ${schedule.unboundCount}`);
      if (schedule.scheduled[1].effect.asc_getDuration() !== 500)
        throw new Error('Generated cue replaced the saved native duration');
      if (schedule.scheduled[1].effect.getObjectId() !== Reflect.get(slide.cSld.spTree[1], 'Id'))
        throw new Error(
          `Speech did not retain the native group animation: ${JSON.stringify({
            scheduled: schedule.scheduled.map((entry: { effect: { getObjectId(): string } }) =>
              entry.effect.getObjectId(),
            ),
            shapes: slide.cSld.spTree.map((shape) => ({
              id: Reflect.get(shape, 'Id'),
              name: shape.getUniNvProps().cNvPr.name,
              children: Reflect.get(shape, 'spTree')?.map((child: { Id: string }) => child.Id),
            })),
          })}`,
        );
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
        await waitReady(started, () => ({
          errors,
          effects: schedule.timing.getAllAnimEffects().map((effect: { isIdle(): boolean; isAtEnd(): boolean }) => ({
            idle: effect.isIdle(),
            ended: effect.isAtEnd(),
          })),
          text: doc.querySelector('[data-teaching-slideshow]')?.textContent,
        }));
        if (!schedule.scheduled.every((entry: { effect: { isIdle(): boolean } }) => entry.effect.isIdle()))
          throw new Error('Speech-linked animation ran before its speech trigger');
        const player = Reflect.get(slide, 'getAnimationPlayer').call(slide);
        const target = schedule.scheduled[0].effect.getObjectId();
        if (player.animationDrawer.isDrawingVisible(target))
          throw new Error('Speech entrance target is visible before its speech trigger');
        const pictures = () => Array.from(doc.querySelectorAll<HTMLCanvasElement>('[data-deckwise-teaching-canvas] canvas'))
          .map(canvas => canvas.toDataURL()).join('|');
        const before = pictures();
        const playing = Promise.all([show.next(0), show.next(1)]);
        await new Promise(resolve => setTimeout(resolve, 200));
        const during = pictures();
        await playing;
        const after = pictures();
        if (!before || before === during || during === after)
          throw new Error('Native fade did not change rendered canvas pixels over time');
        if (!player.animationDrawer.isDrawingVisible(target))
          throw new Error('Speech entrance target did not become visible');
        await show.next(2);
        const visibility = player.animationDrawer.getSandwich(target)?.getAttributesMap(target)['style.visibility'];
        if (mode !== 'emphasis' && visibility !== 'hidden')
          throw new Error('Speech exit did not hide its target');
        if (mode === 'emphasis' && !player.animationDrawer.isDrawingVisible(target))
          throw new Error('Speech emphasis hid the already revealed target');
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
    }, mode);
    expect(result).toEqual({
      ended: true,
      errors: [],
      thumbnails: true,
      located: true,
      restored: true,
      unchangedHistory: true,
    });
  });
