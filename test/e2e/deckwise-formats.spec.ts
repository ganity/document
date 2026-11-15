import { buildDocx, buildPptx, toBase64 } from './lib/ooxml';
import { buildXlsx } from './actions/fixtures';
import { expect, test } from './lib/l0';

const DOCS = [
  { ext: 'docx', bytes: buildDocx('DeckWise preview') },
  { ext: 'xlsx', bytes: buildXlsx() },
  { ext: 'pptx', bytes: buildPptx('DeckWise preview') },
  { ext: 'csv', bytes: new TextEncoder().encode('name,score\nalice,90\nbob,85') },
];

test.describe('DeckWise formats (real editor)', () => {
  test.describe.configure({ timeout: 120_000 });

  for (const doc of DOCS) {
    test(`${doc.ext}: previews readonly and saves in its own format`, async ({ page }) => {
      await page.goto('/embed-demo.html');
      await expect(page.locator('#status')).toHaveText('ready', { timeout: 60_000 });
      await page.evaluate(() => {
        document.getElementById('status')!.textContent = 'loading';
        (document.getElementById('editorFrame') as HTMLIFrameElement).src = './editor?embed=1&deckwise=1';
      });
      await expect(page.locator('#status')).toHaveText('ready', { timeout: 60_000 });

      const opened = await page.evaluate(
        async ({ ext, base64 }) => {
          const bytes = Uint8Array.from(atob(base64), (char) => char.charCodeAt(0));
          return post('document:open-buffer', {
            fileName: `preview.${ext}`,
            buffer: bytes.buffer,
            readonly: true,
          });
        },
        { ext: doc.ext, base64: toBase64(doc.bytes) },
      );
      expect(opened.readonly).toBe(true);
      await expect
        .poll(() => page.evaluate(() => Boolean(window.__ooFrames.readyEditor())), { timeout: 60_000 })
        .toBe(true);

      const result = await page.evaluate(async () => {
        let readonlyError = '';
        try {
          await post('document:save', {});
        } catch (error) {
          readonlyError = String(error);
        }
        await post('document:set-readonly', { readonly: false });
        const saved = await post('document:save', {});
        const bytes = new Uint8Array(await saved.file.arrayBuffer());
        return {
          readonlyError,
          name: saved.file.name,
          size: bytes.length,
          magic: Array.from(bytes.slice(0, 2)),
          text: new TextDecoder().decode(bytes),
        };
      });
      expect(result.readonlyError).toContain('readonly');
      expect(result.name).toBe(`preview.${doc.ext}`);
      expect(result.size).toBeGreaterThan(0);
      if (doc.ext === 'csv') expect(result.text.trim()).toBe(new TextDecoder().decode(doc.bytes));
      else expect(result.magic).toEqual([0x50, 0x4b]);
    });
  }
});
