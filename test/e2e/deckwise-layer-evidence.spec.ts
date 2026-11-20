import { readFile, writeFile } from 'node:fs/promises';
import { test, expect } from './lib/l0';

test('native layer names, descriptions and pixel bindings survive actual Office save and reopen', async ({
  page,
}, info) => {
  test.setTimeout(180000);
  const fixture = process.env.DECKWISE_LAYER_FIXTURE;
  test.skip(!fixture, 'Requires a generated native Scene with four named foregrounds on page one');
  const binding = { projectId: 'layer-fixture', documentId: 'water', sessionId: 'isolated' };
  await page.goto('/embed-demo.html');
  await page.locator('#log').evaluate((node) => {
    node.style.display = 'none';
  });
  await page.evaluate(() => {
    document.querySelector('#status')!.textContent = 'loading';
    window.addEventListener('message', (event) => {
      if (event.data?.type === 'document:teaching-request') {
        (document.querySelector('#editorFrame') as HTMLIFrameElement).contentWindow!.postMessage(
          {
            type: 'document:teaching-response',
            id: event.data.id,
            payload: { result: [] },
          },
          location.origin,
        );
      }
    });
    (document.querySelector('#editorFrame') as HTMLIFrameElement).src = '/editor?embed=1&deckwise=1&locale=en-US';
  });
  await expect(page.locator('#status')).toHaveText('ready', { timeout: 60000 });
  const open = async (bytes: number[]) => {
    await page.evaluate(
      async ({ bytes, binding }) => {
        await post('document:open-buffer', {
          fileName: 'layers.pptx',
          buffer: new Uint8Array(bytes).buffer,
          editorBinding: binding,
          trackPageLineage: true,
        });
      },
      { bytes, binding },
    );
    await page.waitForFunction(() => !!window.__ooFrames.readyEditor(), { timeout: 60000 });
  };
  const capture = () =>
    page.evaluate(async (binding) => {
      const context = (
        await post('document:native-teaching', {
          binding,
          operation: 'route',
          routeOperation: 'narration-context',
          page: 1,
        })
      ).routed.value;
      const visual = (
        await post('document:native-teaching', {
          binding,
          operation: 'route',
          routeOperation: 'narration-visual',
          page: 1,
        })
      ).routed.value;
      return {
        objects: context.evidence.objects,
        layers: visual.layers.map((l: any) => ({ objectId: l.objectId, length: l.dataUrl.length })),
      };
    }, binding);
  await open([...(await readFile(fixture!))]);
  const before = await capture();
  await info.attach('before', { body: JSON.stringify(before), contentType: 'application/json' });
  expect(before.objects.filter((o: any) => o.layerMetadata)).toHaveLength(4);
  expect(before.layers).toHaveLength(4);
  for (const layer of before.layers) {
    expect(layer.length).toBeGreaterThan(100);
    expect(before.objects.some((o: any) => o.objectId === layer.objectId && o.layerMetadata.description)).toBe(true);
  }
  const saved = await page.evaluate(async () => [
    ...new Uint8Array(await (await post('document:save')).file.arrayBuffer()),
  ]);
  await writeFile(info.outputPath('layers-saved.pptx'), new Uint8Array(saved));
  await open(saved);
  const after = await capture();
  expect(after).toEqual(before);
  await page.screenshot({ path: info.outputPath('layers-reopened.png') });
});
