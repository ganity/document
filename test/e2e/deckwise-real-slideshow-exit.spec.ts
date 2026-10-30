import { readFile } from 'node:fs/promises';
import { toBase64 } from './lib/ooxml';
import { expect, test } from './lib/l0';

test('native slideshow exit on supplied courseware emits no SDK runtime error', async ({ page }) => {
  test.skip(!process.env.DECKWISE_EXIT_PPTX, 'Opt-in read-only courseware input');
  test.setTimeout(180_000);
  const bytes = await readFile(process.env.DECKWISE_EXIT_PPTX!);
  await page.goto('/embed-demo.html');
  await page.evaluate(() => {
    document.querySelector('#status')!.textContent = 'loading';
    (document.querySelector('#editorFrame') as HTMLIFrameElement).src = '/editor?embed=1&deckwise=1';
  });
  await expect(page.locator('#status')).toHaveText('ready', { timeout: 60_000 });
  await page.evaluate(async (b64) => {
    const bytes = Uint8Array.from(atob(b64), (c) => c.charCodeAt(0));
    await post('document:open-buffer', { fileName: 'readonly-exit-probe.pptx', buffer: bytes.buffer });
  }, toBase64(bytes));
  await page.waitForFunction(() => !!window.__ooFrames.readyEditor(), { timeout: 60_000 });
  const frame = page.frames().find((frame) => frame.name() === 'frameEditor')!;
  await frame.evaluate(() => {
    const w = window as any;
    w.exitErrors = [];
    w.editor.asc_registerCallback('asc_onError', (...sdk: any[]) => w.exitErrors.push({ sdk }));
    window.addEventListener('error', (event) =>
      w.exitErrors.push({ message: event.message, stack: event.error?.stack }),
    );
    window.addEventListener('unhandledrejection', (event) =>
      w.exitErrors.push({ rejection: String(event.reason), stack: event.reason?.stack }),
    );
  });
  for (const pageNumber of [0, 3, 4]) {
    await frame.evaluate((pageNumber) => {
      const host = document.createElement('div');
      host.id = 'native-exit-probe';
      host.style.cssText = 'position:fixed;inset:0;z-index:10000;background:black';
      document.body.append(host);
      (window as any).editor.StartDemonstration(host.id, pageNumber);
    }, pageNumber);
    await page.waitForTimeout(1000);
    await frame.evaluate(() => {
      (window as any).editor.EndDemonstration();
      document.getElementById('native-exit-probe')!.remove();
    });
    await page.waitForTimeout(1800);
    expect(await frame.evaluate(() => (window as any).exitErrors)).toEqual([]);
    expect(await frame.evaluate(() => (window as any).editor.isSlideShow())).toBe(false);
  }
});
