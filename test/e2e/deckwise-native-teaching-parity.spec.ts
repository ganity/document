import { readFile } from 'node:fs/promises';
import { test, expect } from './lib/l0';
import { toBase64 } from './lib/ooxml';

test('native visual evidence and assessment pictures survive real Office save and reopen', async ({ page }) => {
  test.setTimeout(180000);
  const fixture = process.env.DECKWISE_NATIVE_GENERATED_FIXTURE;
  test.skip(!fixture || !process.env.DECKWISE_AUTHORING_URL, 'Requires generated native lesson fixture and authoring module');
  const original = await readFile(fixture + '/native.pptx');
  const binding = { projectId: 'parity', documentId: 'lesson', sessionId: 'isolated' };
  await page.goto('/embed-demo.html');
  await page.evaluate(() => {
    document.querySelector('#log')!.setAttribute('style', 'display:none');
    (document.querySelector('#editorFrame') as HTMLIFrameElement).src = '/editor?embed=1&deckwise=1&locale=zh-CN';
  });
  await expect(page.locator('#status')).toHaveText('ready', { timeout: 60000 });
  await page.evaluate(async ({ data, binding }) => post('document:open-buffer', {
    fileName: 'parity.pptx', buffer: Uint8Array.from(atob(data), c => c.charCodeAt(0)).buffer, editorBinding: binding, trackPageLineage: true,
  }), { data: toBase64(original), binding });
  await page.waitForFunction(() => !!window.__ooFrames.readyEditor(), { timeout: 60000 });
  const result = await page.evaluate(async binding => {
    const route = async (operation: string, args = {}, page = 1) => (await post('document:native-teaching', { binding, operation: 'route', routeOperation: operation, page, args })).routed.value;
    const before = await route('open');
    const visual = await route('narration-visual');
    const after = await route('open');
    if (before.project.teachingRevision !== after.project.teachingRevision) throw new Error('Rendering mutated teaching');
    const assessment = await route('assessment-context');
    const scene = assessment.classroom.scenes.find((s: any) => s.content.type === 'quiz');
    if (!scene) throw new Error('Generated assessment missing');
    const previous = structuredClone(scene);
    const q = scene.content.questions[0];
    q.question = 'Edited image question'; q.answer = [q.options[0].value]; q.referenceAnswer = 'Saved reference';
    q.options[0].image = { src: 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/l9sAAAAASUVORK5CYII=', mimeType: 'image/png' };
    return { visual: visual.dataUrl.slice(0, 22), sceneId: scene.id, question: q, scene, previous, assessment };
  }, binding);
  const authoring = await import(process.env.DECKWISE_AUTHORING_URL!);
  const { parseMaicClassroom } = await import(new URL('./node.mjs', process.env.DECKWISE_AUTHORING_URL!).href);
  expect(() => parseMaicClassroom(result.assessment.classroom)).not.toThrow();
  const canvas = result.assessment.canvas;
  const { projected, assets } = await authoring.prepareAssessmentPage(result.scene, 'sha256:' + 'a'.repeat(64), canvas, async (src: string) => Buffer.from(src.split(',')[1], 'base64'));
  const revision = await page.evaluate(async ({ binding, result, previous, projected, assets }) => {
    const written = (await post('document:native-teaching', { binding, operation: 'route', routeOperation: 'assessment-write', page: 1,
      args: { projectId: result.assessment.projectId, expectedRevision: result.assessment.revision, expectedSourceRevision: result.assessment.sourceRevision,
        classroom: result.assessment.classroom, changedPageIds: [result.sceneId], projections: [{ page: 1, previous, next: projected, assets: Object.fromEntries(Object.entries(assets).map(([id, bytes]) => [id, new Uint8Array(bytes as number[])])) }] } })).routed.value;
    return written.revision;
  }, { binding, result, previous: authoring.assessmentPptxScene(result.previous, 'sha256:' + 'a'.repeat(64), canvas), projected, assets: Object.fromEntries(Object.entries(assets).map(([id, bytes]) => [id, Array.from(bytes as Uint8Array)])) });
  expect(result.visual).toBe('data:image/png;base64,');
  let native = page.frames().find(f => f.name() === 'frameEditor')!;
  const visibleQuestion = () => native.evaluate(() => (window as any).editor.WordControl.m_oLogicDocument.Slides[0].cSld.spTree.map((s: any) => s.getDocContent?.()?.GetText?.({}) ?? '').join('\n'));
  const imageCount = () => native.evaluate(() => (window as any).editor.WordControl.m_oLogicDocument.Slides[0].cSld.spTree.filter((s: any) => s.isImage?.()).length);
  const editedImages = await imageCount();
  expect(editedImages).toBeGreaterThan(0);
  expect(await visibleQuestion()).toContain('Edited image question');
  await native.evaluate(() => (window as any).editor.Undo());
  expect(await visibleQuestion()).toContain('Choose the greeting');
  expect(await visibleQuestion()).not.toContain('Edited image question');
  expect(await imageCount()).toBe(editedImages - 1);
  const undone = await page.evaluate(async binding => (await post('document:native-teaching', { binding, operation: 'route', routeOperation: 'assessment-context', page: 1 })).routed.value.revision, binding);
  expect(undone).not.toBe(revision);
  await native.evaluate(() => (window as any).editor.Redo());
  expect(await visibleQuestion()).toContain('Edited image question');
  const practiceBefore = await native.evaluate(() => { const shape = (window as any).editor.WordControl.m_oLogicDocument.Slides[1].cSld.spTree.find((s: any) => s.getDocContent?.()?.GetText?.({})?.includes('Greet your friend')); return shape.getDocContent().Content[0].Content[0].Pr.FontSize; });
  const practiceAssessment = await page.evaluate(async binding => {
    const route = async (operation: string, args = {}, page = 1) => (await post('document:native-teaching', { binding, operation: 'route', routeOperation: operation, page, args })).routed.value;
    const assessment = await route('assessment-context');
    const practice = assessment.classroom.scenes.find((s: any) => s.content.type === 'slide' && s.actions.some((a: any) => a.type === 'interaction_checkpoint'));
    if (!practice?.metadata?.nativeQuestionElementId) throw new Error('Practice question has no native text binding');
    const action = practice.actions.find((a: any) => a.type === 'interaction_checkpoint'); action.question = 'Say hello to your friend.';
    return route('assessment-write', { projectId: assessment.projectId, expectedRevision: assessment.revision, expectedSourceRevision: assessment.sourceRevision, classroom: assessment.classroom, changedPageIds: [practice.id] });
  }, binding);
  expect(() => parseMaicClassroom(practiceAssessment.classroom)).not.toThrow();
  const practiceText = await native.evaluate(() => (window as any).editor.WordControl.m_oLogicDocument.Slides[1].cSld.spTree.map((s: any) => s.getDocContent?.()?.GetText?.({}) ?? '').join('\n'));
  expect(practiceText).toContain('Say hello to your friend.');
  const practiceAfter = await native.evaluate(() => { const shape = (window as any).editor.WordControl.m_oLogicDocument.Slides[1].cSld.spTree.find((s: any) => s.getDocContent?.()?.GetText?.({})?.includes('Say hello to your friend.')); return shape.getDocContent().Content[0].Content[0].Pr.FontSize; });
  expect(practiceAfter).toBe(practiceBefore);
  const bytes = await page.evaluate(async () => Array.from(new Uint8Array(await (await post('document:save')).file.arrayBuffer())));
  await page.evaluate(async ({ bytes, binding }) => post('document:open-buffer', {
    fileName: 'parity-reopened.pptx', buffer: new Uint8Array(bytes).buffer, editorBinding: binding, trackPageLineage: true,
  }), { bytes, binding });
  await page.waitForFunction(() => !!window.__ooFrames.readyEditor(), { timeout: 60000 });
  const reopened = await page.evaluate(async ({ binding, sceneId }) => {
    const assessment = (await post('document:native-teaching', { binding, operation: 'route', routeOperation: 'assessment-context', page: 1 })).routed.value;
    return assessment.classroom.scenes.find((s: any) => s.id === sceneId).content.questions[0];
  }, { binding, sceneId: result.sceneId });
  expect(reopened).toEqual(result.question);
  native = page.frames().find(f => f.name() === 'frameEditor')!;
  expect(await visibleQuestion()).toContain('Edited image question');
  expect(await imageCount()).toBe(editedImages);
  const reopenedPractice = await native.evaluate(() => (window as any).editor.WordControl.m_oLogicDocument.Slides[1].cSld.spTree.map((s: any) => s.getDocContent?.()?.GetText?.({}) ?? '').join('\n'));
  expect(reopenedPractice).toContain('Say hello to your friend.');
});
