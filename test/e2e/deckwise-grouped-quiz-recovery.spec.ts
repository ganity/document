import { readFile, writeFile } from 'node:fs/promises';
import { test, expect } from './lib/l0';
import { toBase64 } from './lib/ooxml';

test('an exported grouped Quiz accepts a picture and preserves teaching after save and reopen', async ({
  page,
}, info) => {
  test.setTimeout(180000);
  const fixture = process.env.DECKWISE_GROUPED_QUIZ_FIXTURE;
  test.skip(
    !fixture || !process.env.DECKWISE_AUTHORING_URL,
    'Requires a real PPT-master grouped Quiz and authoring module',
  );
  const binding = { projectId: 'grouped-quiz', documentId: 'lesson', sessionId: 'isolated' };
  await page.goto('/embed-demo.html');
  await page.evaluate(() => {
    document.querySelector('#status')!.textContent = 'loading';
    document.querySelector('#log')!.setAttribute('style', 'display:none');
    (document.querySelector('#editorFrame') as HTMLIFrameElement).src = '/editor?embed=1&deckwise=1&locale=zh-CN';
  });
  await expect(page.locator('#status')).toHaveText('ready', { timeout: 60000 });
  const open = async (bytes: Uint8Array) => {
    await page.evaluate(
      async ({ data, binding }) =>
        post('document:open-buffer', {
          fileName: 'grouped-quiz.pptx',
          buffer: Uint8Array.from(atob(data), (c) => c.charCodeAt(0)).buffer,
          editorBinding: binding,
          trackPageLineage: true,
        }),
      { data: toBase64(bytes), binding },
    );
    await page.waitForFunction(() => !!window.__ooFrames.readyEditor(), { timeout: 60000 });
  };
  await open(await readFile(fixture!));
  const context = await page.evaluate(async (binding) => {
    const route = async (operation: string) =>
      (
        await post('document:native-teaching', {
          binding,
          operation: 'route',
          routeOperation: operation,
          page: 1,
          args: {},
        })
      ).routed.value;
    return { assessment: await route('assessment-context'), teaching: (await route('open')).project };
  }, binding);
  const sceneIndex = context.assessment.classroom.scenes.findIndex((s: any) => s.content.type === 'quiz');
  expect(sceneIndex).toBeGreaterThanOrEqual(0);
  const originalScene = context.assessment.classroom.scenes[sceneIndex];
  const scene = structuredClone(originalScene);
  scene.content.questions[0].options[0].image = {
    src: 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/l9sAAAAASUVORK5CYII=',
    mimeType: 'image/png',
  };
  context.assessment.classroom.scenes[sceneIndex] = scene;
  const nativeBefore = await page
    .frames()
    .find((f) => f.name() === 'frameEditor')!
    .evaluate((n) => {
      const slide = (window as any).editor.WordControl.m_oLogicDocument.Slides[n];
      return slide.cSld.spTree.map((s: any) => ({ name: s.getObjectName(), grouped: !!s.spTree }));
    }, sceneIndex);
  expect(nativeBefore.some((s: { grouped: boolean; name: string }) => s.grouped && s.name === 'assessment')).toBe(true);
  const authoring = await import(process.env.DECKWISE_AUTHORING_URL!);
  const previous = authoring.assessmentPptxScene(originalScene, 'sha256:' + 'a'.repeat(64), context.assessment.canvas);
  const { projected, assets } = await authoring.prepareAssessmentPage(
    scene,
    'sha256:' + 'a'.repeat(64),
    context.assessment.canvas,
    async (src: string) => Buffer.from(src.split(',')[1], 'base64'),
  );
  await page.evaluate(
    async ({ binding, assessment, previous, projected, assets, sceneIndex, sceneId }) => {
      await post('document:native-teaching', {
        binding,
        operation: 'route',
        routeOperation: 'assessment-write',
        page: sceneIndex + 1,
        args: {
          projectId: assessment.projectId,
          expectedRevision: assessment.revision,
          expectedSourceRevision: assessment.sourceRevision,
          classroom: assessment.classroom,
          changedPageIds: [sceneId],
          projections: [
            {
              page: sceneIndex + 1,
              previous,
              next: projected,
              assets: Object.fromEntries(
                Object.entries(assets).map(([id, bytes]) => [id, new Uint8Array(bytes as number[])]),
              ),
            },
          ],
        },
      });
    },
    {
      binding,
      assessment: context.assessment,
      previous,
      projected,
      sceneIndex,
      sceneId: scene.id,
      assets: Object.fromEntries(Object.entries(assets).map(([id, bytes]) => [id, Array.from(bytes as Uint8Array)])),
    },
  );
  const saved = await page.evaluate(async () =>
    Array.from(new Uint8Array(await (await post('document:save')).file.arrayBuffer())),
  );
  await writeFile(info.outputPath('recovered-grouped-quiz.pptx'), new Uint8Array(saved));
  await open(new Uint8Array(saved));
  const reopened = await page.evaluate(async (binding) => {
    const route = async (operation: string) =>
      (
        await post('document:native-teaching', {
          binding,
          operation: 'route',
          routeOperation: operation,
          page: 1,
          args: {},
        })
      ).routed.value;
    return { assessment: await route('assessment-context'), teaching: (await route('open')).project };
  }, binding);
  expect(reopened.teaching.pages.map((p: any) => p.actions)).toEqual(context.teaching.pages.map((p: any) => p.actions));
  expect(reopened.assessment.classroom.scenes[sceneIndex].content.questions).toEqual(scene.content.questions);
  const nativeAfter = await page
    .frames()
    .find((f) => f.name() === 'frameEditor')!
    .evaluate((n) => {
      const shapes = (window as any).editor.WordControl.m_oLogicDocument.Slides[n].cSld.spTree;
      return shapes.map((s: any) => ({
        name: s.getObjectName(),
        image: s.getObjectType() === (window as any).AscDFH.historyitem_type_ImageShape,
      }));
    }, sceneIndex);
  expect(
    nativeAfter.some((s: { image: boolean; name: string }) => s.image && s.name.endsWith('_q0_option0_image')),
  ).toBe(true);
});
