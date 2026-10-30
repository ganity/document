import { mkdir, writeFile } from 'node:fs/promises';
import { join } from 'node:path';
import { buildPptx, toBase64, zipEntryText } from './lib/ooxml';
import { expect, test } from './lib/l0';

// Real SDK export and acknowledgement; the test parent stands in for Electron's disk commit.
test('native PPT edits automatically export with object lineage and renew the saved baseline', async ({ page }) => {
  test.setTimeout(180_000);
  await page.goto('/embed-demo.html');
  await page.evaluate(() => {
    document.querySelector('#status')!.textContent = 'loading';
    (document.querySelector('#editorFrame') as HTMLIFrameElement).src = '/editor?embed=1&deckwise=1';
  });
  await expect(page.locator('#status')).toHaveText('ready', { timeout: 60_000 });
  await page.evaluate(async b64 => {
    const bytes = Uint8Array.from(atob(b64), c => c.charCodeAt(0));
    await post('document:open-buffer', { fileName: 'autosave.pptx', buffer: bytes.buffer,
      editorBinding: { projectId: 'probe', documentId: 'native-autosave', sessionId: 'session' }, trackPageLineage: true });
    const frame = (document.querySelector('#editorFrame') as HTMLIFrameElement).contentWindow!;
    (window as any).automaticSaves = [];
    window.addEventListener('message', async event => {
      if (event.source !== frame || event.origin !== location.origin) return;
      if (event.data?.type === 'document:save-requested') {
        const saved = await post('document:save', { requireHostCommit: true });
        const file = saved.file as File;
        const bytes = new Uint8Array(await file.arrayBuffer());
        const sha = Array.from(new Uint8Array(await crypto.subtle.digest('SHA-256', bytes)), b => b.toString(16).padStart(2, '0')).join('');
        if (sha !== saved.pageLineage?.outputSha256) throw new Error('Export lineage hash mismatch');
        (window as any).automaticSaves.push({ ...saved.pageLineage, bytes: Array.from(bytes) });
        frame.postMessage({ type: 'document:commit-save', id: saved.pageLineage.saveRequestId, payload: { success: true } }, location.origin);
      }
    });
  }, toBase64(buildPptx('Automatic save')));
  await page.waitForFunction(() => !!window.__ooFrames.readyEditor(), { timeout: 60_000 });
  const native = page.frames().find(frame => frame.name() === 'frameEditor')!;
  const move = () => native.evaluate(() => {
    const w = window as any, api = w.editor, logic = api.WordControl.m_oLogicDocument;
    const shape = logic.Slides[0].cSld.spTree[0];
    logic.StartAction(w.AscDFH.historydescription_Presentation_ApplyShapeProps);
    shape.spPr.xfrm.setOffX(shape.spPr.xfrm.offX + 10);
    logic.Recalculate();
    logic.FinalizeAction();
    api.CheckChangedDocument();
    document.dispatchEvent(new Event('pointerup'));
  });
  await move();
  // Verify the idle deadline follows a second edit while the SDK is already dirty.
  await page.waitForTimeout(1000);
  await move();
  await page.waitForTimeout(1000);
  expect(await page.evaluate(() => (window as any).automaticSaves.length)).toBe(0);
  await expect.poll(() => page.evaluate(() => (window as any).automaticSaves.length), { timeout: 60_000 }).toBe(1);
  await expect.poll(() => page.evaluate(async () => (await post('document:get-state')).dirty)).toBe(false);
  const first = await page.evaluate(() => (window as any).automaticSaves[0]);
  expect(first.objects[0][0]).toMatchObject({ sourceName: 'Title 1', name: 'Title 1' });
  expect(first.objects[0][0].sourcePageToken).toBe(first.openedPageTokens[0]);
  const xml = await zipEntryText(Uint8Array.from(first.bytes), 'ppt/slides/slide1.xml');
  expect(xml).toContain('name="Title 1"');
  await move();
  await expect.poll(() => page.evaluate(() => (window as any).automaticSaves.length), { timeout: 60_000 }).toBe(2);
  const second = await page.evaluate(() => (window as any).automaticSaves[1]);
  expect(second.sourceSha256).toBe(first.outputSha256);
  expect(second.objects[0][0].sourceName).toBe('Title 1');
  if (process.env.DECKWISE_AUTOSAVE_FIXTURE) {
    const directory = process.env.DECKWISE_AUTOSAVE_FIXTURE;
    await mkdir(directory, { recursive: true });
    await writeFile(join(directory, 'original.pptx'), buildPptx('Automatic save'));
    await writeFile(join(directory, 'first.pptx'), Uint8Array.from(first.bytes));
    await writeFile(join(directory, 'second.pptx'), Uint8Array.from(second.bytes));
    await writeFile(join(directory, 'lineage.json'), JSON.stringify([first, second]));
  }

});
