import { readFile, writeFile } from 'node:fs/promises';
import { test, expect } from './lib/l0';
import { toBase64 } from './lib/ooxml';

test('accepted script and a new independent Quiz survive one native Undo/Redo, save and reopen', async ({
  page,
}, testInfo) => {
  test.setTimeout(240000);
  const fixture = process.env.DECKWISE_NATIVE_GENERATED_FIXTURE;
  const authoringUrl = process.env.DECKWISE_AUTHORING_URL;
  test.skip(!fixture || !authoringUrl, 'Requires copied native lesson and assessment authoring module');
  const original = await readFile(fixture + '/native.pptx');
  const binding = { projectId: 'materials-fixture', documentId: 'lesson', sessionId: 'isolated' };
  await page.goto('/embed-demo.html');
  await page.evaluate(() => {
    document.querySelector('#status')!.textContent = 'loading';
    document.querySelector('#log')!.setAttribute('style', 'display:none');
    (document.querySelector('#editorFrame') as HTMLIFrameElement).src = '/editor?embed=1&deckwise=1&locale=zh-CN';
  });
  await expect(page.locator('#status')).toHaveText('ready', { timeout: 60000 });
  await page.evaluate(
    async ({ data, binding }) =>
      post('document:open-buffer', {
        fileName: 'materials.pptx',
        buffer: Uint8Array.from(atob(data), (c) => c.charCodeAt(0)).buffer,
        editorBinding: binding,
        trackPageLineage: true,
      }),
    { data: toBase64(original), binding },
  );
  await page.waitForFunction(() => !!window.__ooFrames.readyEditor(), { timeout: 60000 });
  const accepted = await page.evaluate(async (binding) => {
    const route = async (routeOperation: string, args: unknown = {}, page = 1) =>
      (await post('document:native-teaching', { binding, operation: 'route', routeOperation, args, page })).routed
        .value;
    let view = await route('open');
    const enrichment = {
      schemaVersion: 'deckwise-teaching-enrichment/1',
      originalPages: view.project.pages.map((p: any, i: number) => ({ pageId: p.pageId, pageNumber: i + 1 })),
      script: {
        sha256: 'a'.repeat(64),
        policy: 'preserve',
        silentPageIds: [],
        segments: view.project.pages.map((p: any, i: number) => ({
          id: `source-${i}`,
          text: `Exact speech ${i + 1}. 1.25 kg. Watch the video!`,
          sourceLine: i + 1,
          sourcePageNumber: i + 1,
          pageId: p.pageId,
          actionId: `accepted-${i}`,
        })),
      },
      plan: {
        sha256: 'b'.repeat(64),
        units: [{ id: 'review', title: 'Round shapes', evidence: 'A circle is round.', kind: 'quiz', pageIds: [] }],
      },
    };
    for (let i = 0; i < view.project.pages.length; i++) {
      const p = view.project.pages[i],
        segment = enrichment.script.segments[i];
      view = await route(
        'write',
        {
          viewId: view.project.projectId + '.' + p.pageId,
          expectedTeachingRevision: view.project.teachingRevision,
          actions: [
            { id: segment.actionId, type: 'speech', text: segment.text },
            ...p.actions.filter((a: any) => a.type !== 'speech'),
          ],
          enrichment,
        },
        i + 1,
      );
    }
    return { view, enrichment, count: view.project.pages.length };
  }, binding);
  const { assessmentPptxScene } = await import(authoringUrl!);
  const quiz = {
    type: 'quiz',
    interactionMode: 'classroom_instant_feedback',
    questions: [
      {
        id: 'plan-question',
        type: 'single',
        question: 'Which shape is round?',
        options: [
          { value: 'A', label: 'circle' },
          { value: 'B', label: 'square' },
        ],
        answer: ['A'],
        explanation: 'A circle is round.',
      },
    ],
  };
  const pageId = 'plan-activity';
  const projection = assessmentPptxScene(
    {
      id: pageId,
      stageId: accepted.view.project.projectId,
      title: 'Round shapes',
      type: 'quiz',
      order: 1,
      content: quiz,
      actions: [],
    },
    'sha256:' + 'a'.repeat(64),
    { width: 1000, height: 562.5 },
  );
  const inserted = await page.evaluate(
    async ({ binding, accepted, quiz, projection, pageId }) => {
      const enrichment = accepted.enrichment;
      Object.assign(enrichment.plan.units[0], { pageIds: [pageId], activityId: 'plan-question' });
      const view = (
        await post('document:native-teaching', { binding, operation: 'route', routeOperation: 'open', page: 1 })
      ).routed.value;
      return (
        await post('document:native-teaching', {
          binding,
          operation: 'route',
          routeOperation: 'assessment-insert',
          page: 1,
          args: {
            viewId: view.viewId,
            expectedTeachingRevision: view.project.teachingRevision,
            pageId,
            quiz,
            projection,
            enrichment,
            actions: [
              { id: 'plan-reading', type: 'speech', text: 'Which shape is round? A. circle B. square' },
              {
                id: 'plan-question',
                type: 'interaction_checkpoint',
                checkpointId: 'plan-question',
                checkpointKind: 'chapter_summary',
                chapterId: pageId,
                chapterTitle: 'Round shapes',
                question: 'Which shape is round?',
                response: { kind: 'single', options: ['circle', 'square'], answerIndices: [0] },
                referenceAnswer: { summary: 'A circle is round.', keyPoints: ['circle'] },
                context: { chapterSummary: 'Shapes', sourceSceneIds: [view.pageId] },
              },
            ],
          },
        })
      ).routed.value;
    },
    { binding, accepted, quiz, projection, pageId },
  );
  expect(inserted.project.pages.length).toBe(accepted.count + 1);
  expect(inserted.project.pages[1].pageId).toBe(pageId);
  const native = page.frames().find((f) => f.name() === 'frameEditor')!;
  expect(
    await native.evaluate(() =>
      (window as any).editor.WordControl.m_oLogicDocument.Slides[1].cSld.spTree
        .map((s: any) => s.getDocContent?.()?.GetText?.({}) ?? '')
        .join('\n'),
    ),
  ).toContain('Which shape is round?');
  for (const [operation, count] of [
    ['undo', accepted.count],
    ['redo', accepted.count + 1],
  ] as const) {
    const value = await page.evaluate(
      async ({ binding, operation }) =>
        (
          await post('document:native-teaching', {
            binding,
            operation: 'route',
            routeOperation: operation,
            args: {},
            page: 1,
          })
        ).routed.value,
      { binding, operation },
    );
    expect(value.project.pages.length).toBe(count);
  }
  const bytes = await page.evaluate(async () =>
    Array.from(new Uint8Array(await (await post('document:save')).file.arrayBuffer())),
  );
  await writeFile(testInfo.outputPath('materials-native.pptx'), new Uint8Array(bytes));
  await writeFile(testInfo.outputPath('materials-enrichment.json'), JSON.stringify(inserted.enrichment, null, 2));
  await page.evaluate(
    async ({ bytes, binding }) =>
      post('document:open-buffer', {
        fileName: 'materials-reopened.pptx',
        buffer: new Uint8Array(bytes).buffer,
        editorBinding: binding,
        trackPageLineage: true,
      }),
    { bytes, binding },
  );
  await page.waitForFunction(() => !!window.__ooFrames.readyEditor(), { timeout: 60000 });
  const reopened = await page.evaluate(
    async (binding) =>
      (
        await post('document:native-teaching', {
          binding,
          operation: 'route',
          routeOperation: 'open',
          args: {},
          page: 1,
        })
      ).routed.value,
    binding,
  );
  expect(reopened.enrichment).toEqual(inserted.enrichment);
  for (const segment of reopened.enrichment.script.segments) {
    expect(
      reopened.project.pages
        .find((p: any) => p.pageId === segment.pageId)
        .actions.find((a: any) => a.id === segment.actionId).text,
    ).toBe(segment.text);
  }
  const assessments = await page.evaluate(
    async (binding) =>
      (
        await post('document:native-teaching', {
          binding,
          operation: 'route',
          routeOperation: 'assessment-context',
          args: {},
          page: 2,
        })
      ).routed.value,
    binding,
  );
  expect(assessments.classroom.scenes.find((s: any) => s.id === pageId).content.questions[0].answer).toEqual(['A']);
});
