import { readFile, writeFile } from 'node:fs/promises';
import { test, expect } from './lib/l0';
import { toBase64 } from './lib/ooxml';

test('native narration survives real Office navigation and refreshes edited content before saving', async ({ page }, testInfo) => {
  test.setTimeout(240000);
  const fixture = process.env.DECKWISE_NATIVE_GENERATED_FIXTURE;
  test.skip(!fixture || !process.env.DECKWISE_NATIVE_NARRATION_URL, 'Requires native fixture and Desktop narration module');
  const { generateNativeNarration } = await import(process.env.DECKWISE_NATIVE_NARRATION_URL!);
  const binding = { projectId: 'recovery', documentId: 'lesson', sessionId: 'isolated' };
  await page.goto('/embed-demo.html');
  await page.evaluate(() => {
    (window as any).log = () => {};
    document.querySelector('#log')!.setAttribute('style', 'display:none');
    (document.querySelector('#editorFrame') as HTMLIFrameElement).src = '/editor?embed=1&deckwise=1&locale=zh-CN';
  });
  await expect(page.locator('#status')).toHaveText('ready', { timeout: 60000 });
  const bytes = await readFile(fixture + '/native.pptx');
  await page.evaluate(async ({ data, binding }) => post('document:open-buffer', {
    fileName: 'recovery.pptx', buffer: Uint8Array.from(atob(data), c => c.charCodeAt(0)).buffer, editorBinding: binding, trackPageLineage: true,
  }), { data: toBase64(bytes), binding });
  await page.waitForFunction(() => !!window.__ooFrames.readyEditor(), { timeout: 60000 });
  const native = async (operation: string, args = {}, n = 1) => {
    return page.evaluate(async ({ binding, operation, args, n }) => {
      const editor = await post('document:get-context', binding);
      const result = (await post('document:native-teaching', { binding, operation: 'route', routeOperation: operation, page: n, args })).routed;
      return operation === 'narration-context' ? { value: { ...result.value, editor } } : result;
    }, { binding, operation, args, n });
  };
  let speeches = 0;
  const run = (mutate: () => Promise<void>) => generateNativeNarration({ teaching: { native }, sessionId: 'isolated', selected: [1],
    instruction: 'Explain the visible question', model: { modelKey: 'fixture' }, canDispatch: () => true, assertModel: async () => {},
    complete: async ({ messages }: any) => {
      const data = JSON.parse(messages[0].content);
      if (data.outline) return { text: JSON.stringify({ goal: 'Question', pages: data.requestedPageNumbers.map((pageNumber: number) => ({ pageNumber, purpose: 'Explain', focus: 'Question', bridge: 'Continue' })) }) };
      speeches++;
      await mutate();
      return { text: JSON.stringify({ blocks: [{ type: 'speech', text: 'Read the visible question carefully.' }] }) };
    } });
  const apiFrame = page.frames().find(f => f.name() === 'frameEditor')!;
  const navigation = await run(async () => {
    await apiFrame.evaluate(() => { (window as any).editor.goToPage(1); });
  });
  expect(navigation.saved).toBe(true); expect(speeches).toBe(1);
  let edited = false;
  const contentEdit = await run(async () => {
    if (edited) return; edited = true;
    await apiFrame.evaluate(() => {
      const api = (window as any).editor;
      api.goToPage(0);
      const visit = (shapes: any[]): any[] => shapes.flatMap(s => [s, ...visit(s.spTree ?? [])]);
      const shape = visit(api.WordControl.m_oLogicDocument.Slides[0].cSld.spTree).find((s: any) => s.spPr?.xfrm && typeof s.spPr.xfrm.setOffX === 'function');
      if (!shape) throw new Error('Fixture text shape unavailable');
      const logic = api.WordControl.m_oLogicDocument;
      logic.StartAction((window as any).AscDFH.historydescription_Presentation_ApplyShapeProps);
      shape.spPr.xfrm.setOffX(shape.spPr.xfrm.offX + 3);
      logic.Recalculate();
      logic.FinalizeAction();
    });
  });
  await writeFile(testInfo.outputPath('content-edit-result.json'), JSON.stringify(contentEdit, null, 2));

  expect(contentEdit.saved).toBe(true);
  expect(contentEdit.diagnostics.some((d: any) => /Re-read/.test(d.message))).toBe(true);
  expect(speeches).toBe(3);
  const before = (await native('open'))!.value;
  const context = await page.evaluate(async binding => post('document:get-context', binding), binding);
  await expect(page.evaluate(async ({ binding, before, context }) => post('document:native-teaching', {
    binding, operation: 'route', routeOperation: 'write', page: 1,
    args: { viewId: before.viewId, expectedTeachingRevision: before.project.teachingRevision,
      expectedEditGeneration: context.editGeneration - 1, actions: [] },
  }), { binding, before, context })).rejects.toThrow(/generation changed/);
  expect((await native('open'))!.value.project.teachingRevision).toBe(before.project.teachingRevision);
});
