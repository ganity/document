import { buildPptx, makeStoredZip, toBase64, zipEntryNames, zipEntryText } from './lib/ooxml';
import { test, expect } from './lib/l0';

// Capability evidence, not authority activation: retain the actual converter output.
test('reports teaching metadata preservation through the bundled PPTX converter', async ({ page }, info) => {
  test.setTimeout(120000);
  const namespace = 'urn:deckwise:teaching:capability:1';
  const objectId = '771101f9-7cbe-4e64-ae51-083b680b105a';
  const slideId = '30c68027-5278-495c-a957-d79b825b5830';
  const original = buildPptx('Teaching round trip');
  const entries = await Promise.all(
    zipEntryNames(original).map(async (name) => ({
      name,
      data: (await zipEntryText(original, name))!,
    })),
  );
  const slide = entries.find((entry) => entry.name === 'ppt/slides/slide1.xml')!;
  slide.data = slide.data.replace(
    /<p:cNvPr id="2" name="Title 1"\s*\/>/,
    `<p:cNvPr id="2" name="Title 1"><a:extLst><a:ext uri="${namespace}"><dw:object xmlns:dw="${namespace}" id="${objectId}"/></a:ext></a:extLst></p:cNvPr>`,
  );
  slide.data = slide.data.replace(
    '</p:sld>',
    `<p:extLst><p:ext uri="${namespace}"><dw:page xmlns:dw="${namespace}" id="${slideId}"/></p:ext></p:extLst></p:sld>`,
  );
  expect(slide.data).toContain(objectId);
  const relationships = entries.find((entry) => entry.name === 'ppt/_rels/presentation.xml.rels')!;
  relationships.data = relationships.data.replace(
    '</Relationships>',
    '<Relationship Id="rIdTeaching" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml" Target="../customXml/item1.xml"/></Relationships>',
  );
  entries.push({
    name: 'customXml/item1.xml',
    data: `<dw:teaching xmlns:dw="${namespace}"><dw:speech>Portable narration</dw:speech></dw:teaching>`,
  });
  await page.goto('/embed-demo.html');
  await page.evaluate(() => {
    document.querySelector('#status')!.textContent = 'loading';
    (document.querySelector('#editorFrame') as HTMLIFrameElement).src = '/editor?embed=1&deckwise=1';
  });
  await expect(page.locator('#status')).toHaveText('ready', { timeout: 60000 });
  await page.evaluate(
    async (base64) => {
      await post('document:open-buffer', {
        fileName: 'teaching-capability.pptx',
        editorBinding: { projectId: 'capability', documentId: 'fixture', sessionId: 'isolated' },
        trackPageLineage: true,
        buffer: Uint8Array.from(atob(base64), (c) => c.charCodeAt(0)).buffer,
      });
    },
    toBase64(makeStoredZip(entries)),
  );
  await page.waitForFunction(() => !!window.__ooFrames.readyEditor(), { timeout: 60000 });
  const native = page.frames().find((frame) => frame.name() === 'frameEditor')!;
  const nativeIds = await native.evaluate(() => {
    const api = Reflect.get(window, 'editor') as unknown as {
      WordControl: {
        m_oLogicDocument: {
          Slides: Array<{ cSld: { spTree: Array<{ getUniNvProps(): { cNvPr: { id: number } } }> } }>;
        };
      };
    };
    const slides = api.WordControl.m_oLogicDocument.Slides;
    return slides.map((slide: { cSld: { spTree: Array<{ getUniNvProps(): { cNvPr: { id: number } } }> } }) =>
      slide.cSld.spTree.map((shape) => shape.getUniNvProps().cNvPr.id),
    );
  });
  const historyProbe = await native.evaluate(() => {
    const api = Reflect.get(window, 'editor') as unknown as {
      WordControl: {
        m_oLogicDocument: {
          StartAction(description: number): void;
          FinalizeAction(): void;
          Slides: Array<{
            cSld: { spTree: Array<{ getUniNvProps(): { cNvPr: { name: string; setName(name: string): void } } }> };
          }>;
        };
      };
      Undo(): void;
    };
    const builder = Reflect.get(window, 'AscBuilder');
    const Constructor = builder?.ApiCustomXmlParts;
    if (typeof Constructor !== 'function') return { available: false, undoRestoresParts: false };
    const parts = new Constructor(api.WordControl.m_oLogicDocument);
    const count = parts.GetCount();
    const shape = api.WordControl.m_oLogicDocument.Slides[0].cSld.spTree[0].getUniNvProps().cNvPr;
    const name = shape.name;
    const history = Reflect.get(window, 'AscCommon').History;
    const index = history.Index;
    const logic = api.WordControl.m_oLogicDocument;
    const description = Reflect.get(window, 'AscDFH').historydescription_Document_ApiBuilder;
    logic.StartAction(description);
    const added = parts.Add('<dw:probe xmlns:dw="urn:deckwise:history-probe"/>');
    shape.setName('Native and teaching transaction');
    logic.FinalizeAction();
    const recorded = history.Index !== index;
    api.Undo();
    const undoRestoresParts = recorded && parts.GetCount() === count;
    const undoRestoresShape = shape.name === name;
    const redoApi = api as typeof api & { Redo(): void };
    redoApi.Redo();
    const redoRestoresBoth = parts.GetCount() === count + 1 && shape.name === 'Native and teaching transaction';
    api.Undo();
    if (!undoRestoresParts) added.Delete();
    return { available: true, recorded, undoRestoresParts, undoRestoresShape, redoRestoresBoth };
  });
  expect(historyProbe).toMatchObject({
    available: true,
    recorded: true,
    undoRestoresParts: true,
    undoRestoresShape: true,
    redoRestoresBoth: true,
  });
  const exported = await page.evaluate(async () => {
    const saved = await post('document:save', { requireHostCommit: true });
    const frame = (document.querySelector('#editorFrame') as HTMLIFrameElement).contentWindow!;
    await new Promise<void>((resolve) => {
      const receive = (event: MessageEvent) => {
        if (
          event.source !== frame ||
          event.origin !== location.origin ||
          event.data?.type !== 'document:save-committed' ||
          event.data.id !== saved.pageLineage.saveRequestId
        )
          return;
        window.removeEventListener('message', receive);
        resolve();
      };
      window.addEventListener('message', receive);
      frame.postMessage(
        { type: 'document:commit-save', id: saved.pageLineage.saveRequestId, payload: { success: false } },
        location.origin,
      );
    });
    // No disk writer participates in a capability probe. Reopen from the captured bytes.
    return { bytes: Array.from(new Uint8Array(await saved.file.arrayBuffer())), lineage: saved.pageLineage };
  });
  const output = new Uint8Array(exported.bytes);
  const xml = await Promise.all(
    zipEntryNames(output)
      .filter((name) => name.endsWith('.xml'))
      .map((name) => zipEntryText(output, name)),
  );
  const serialized = xml.join('\n');
  const report = {
    nativeIds,
    lineage: exported.lineage,
    exportedShapeIds: (await zipEntryText(output, 'ppt/slides/slide1.xml'))?.match(/<p:cNvPr[^>]*>/g),
    customXmlPreserved: serialized.includes('Portable narration') && serialized.includes(namespace),
    objectIdentityPreserved: serialized.includes(objectId),
    slideIdentityPreserved: serialized.includes(slideId),
    nativeHistoryProbe: historyProbe,
    nativeHistoryQualified: false,
    activationQualified: false,
  };
  expect(exported.lineage?.protocol).toBe('deckwise-office-page-lineage/2');
  expect(exported.lineage.objects[0][0].sourceShapeId).toBe('2');
  expect(exported.lineage.objects[0][0].shapeId).toBe(String(nativeIds[0][0]));
  expect(await zipEntryText(output, 'ppt/slides/slide1.xml')).toContain('Teaching round trip');
  await info.attach('capability-report', { body: JSON.stringify(report, null, 2), contentType: 'application/json' });
  console.log('Teaching capability:', JSON.stringify(report));
  await info.attach('exported-presentation', {
    body: Buffer.from(output),
    contentType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
  });
  await page.evaluate(async (bytes) => {
    await post('document:open-buffer', {
      fileName: 'reopened.pptx',
      buffer: new Uint8Array(bytes).buffer,
      editorBinding: { projectId: 'capability', documentId: 'fixture', sessionId: 'reopened' },
      trackPageLineage: true,
    });
  }, exported.bytes);
  await page.waitForFunction(() => !!window.__ooFrames.readyEditor(), { timeout: 60000 });
  const reopenedIds = await page
    .frames()
    .find((frame) => frame.name() === 'frameEditor')!
    .evaluate(() => {
      const api = Reflect.get(window, 'editor') as unknown as {
        WordControl: {
          m_oLogicDocument: {
            Slides: Array<{ cSld: { spTree: Array<{ getUniNvProps(): { cNvPr: { id: number } } }> } }>;
          };
        };
      };
      return api.WordControl.m_oLogicDocument.Slides.map(
        (slide: { cSld: { spTree: Array<{ getUniNvProps(): { cNvPr: { id: number } } }> } }) =>
          slide.cSld.spTree.map((shape) => shape.getUniNvProps().cNvPr.id),
      );
    });
  expect(reopenedIds[0][0]).not.toBe(nativeIds[0][0]);
  const reopened = await page.evaluate(async () => {
    const saved = await post('document:save', { requireHostCommit: true });
    return saved.pageLineage;
  });
  expect(reopened.protocol).toBe('deckwise-office-page-lineage/2');
  expect(reopened.objects[0][0].sourceShapeId).toBe(String(nativeIds[0][0]));
});
