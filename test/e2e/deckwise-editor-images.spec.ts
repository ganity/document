import { expect, test } from './lib/l0';
import { buildPptx, toBase64, zipEntryText } from './lib/ooxml';

test('DeckWise generates an art image, retains its parameters on reopen and replaces the selected picture', async ({
  page,
}) => {
  test.setTimeout(150_000);
  await page.addInitScript(() => {
    if (window.parent === window) return;
    const listeners = new Set<(data: unknown) => void>();
    window.addEventListener('message', (event) => {
      if (event.source === parent && event.origin === location.origin) listeners.forEach((fn) => fn(event.data));
    });
    window.deckwiseEditorHost = {
      onCommand(fn) {
        listeners.add(fn);
        return () => {
          listeners.delete(fn);
        };
      },
      sendEvent(event) {
        if (event.type === 'message') parent.postMessage(event.message, location.origin);
        if (event.type === 'saved')
          parent.postMessage(
            {
              type: 'document:saved',
              id: event.id,
              payload: { file: new File([event.bytes!], event.fileName!) },
            },
            location.origin,
          );
      },
    };
  });
  await page.goto('/embed-demo.html');
  await page.evaluate(() => {
    const frame = document.querySelector<HTMLIFrameElement>('#editorFrame')!;
    const canvas = document.createElement('canvas');
    canvas.width = 128;
    canvas.height = 64;
    const ctx = canvas.getContext('2d')!;
    ctx.fillStyle = '#000';
    ctx.fillRect(0, 0, 128, 64);
    ctx.fillStyle = '#ffa526';
    ctx.fillRect(16, 16, 96, 32);
    const image = canvas.toDataURL();
    ctx.fillStyle = '#ff00ff';
    ctx.fillRect(0, 0, 128, 64);
    ctx.fillStyle = '#ffffff';
    ctx.fillRect(16, 16, 96, 32);
    const foreground = canvas.toDataURL();
    Object.assign(window, { imageCalls: [] });
    window.addEventListener('message', (event) => {
      if (event.source !== frame.contentWindow) return;
      if (event.data?.type === 'document:ready' && frame.contentWindow?.location.pathname === '/module.html')
        document.body.dataset.moduleReady = 'true';
      if (event.data?.type !== 'document:teaching-request') return;
      const { operation, args } = event.data.payload;
      if (operation.startsWith('editor-image:')) (window as any).imageCalls.push({ operation, args });
      const result =
        operation === 'editor-image:models'
          ? [{ model: 'test-image', supportsReferenceImage: true }, { model: 'text-only' }]
          : operation === 'editor-image:quote'
            ? { id: 'test-quote', maxPoints: 3, expiresAt: new Date(Date.now() + 60_000).toISOString() }
            : operation === 'editor-image:submit'
              ? { state: 'ready', image: (window as any).imageCalls.at(-2)?.args.prompt.includes('#FF00FF') ? foreground : image }
              : [];
      frame.contentWindow!.postMessage(
        { type: 'document:teaching-response', id: event.data.id, payload: { result } },
        location.origin,
      );
    });
    document.querySelector('#status')!.textContent = 'loading module';
    frame.src = '/module.html?deckwise=1';
  });
  await expect(page.locator('#status')).toHaveText('ready', { timeout: 30_000 });
  await expect(page.locator('body')).toHaveAttribute('data-module-ready', 'true', { timeout: 30000 });
  const binding = { projectId: 'image-project', documentId: 'image-deck', sessionId: 'image-session' };
  const open = async (data: string) => {
    await page.evaluate(
      async ({ data, binding }) => {
        await post('document:open-buffer', {
          fileName: 'art-text.pptx',
          buffer: Uint8Array.from(atob(data), (c) => c.charCodeAt(0)).buffer,
          editorBinding: binding,
        });
      },
      { data, binding },
    );
    await page.waitForFunction(() => !!window.__ooFrames.readyEditor(), { timeout: 60_000 });
    await page.evaluate((binding) => {
      document.querySelector<HTMLIFrameElement>('#editorFrame')!.contentWindow!.postMessage(
        {
          type: 'document:teaching-enable',
          payload: { binding, editorImages: true },
        },
        location.origin,
      );
    }, binding);
  };
  await open(toBase64(buildPptx('Title')));
  const editor = page.frameLocator('#editorFrame').frameLocator('#app iframe');
  const tools = editor.locator('[data-editor-image-tools]');
  const generatePreview = async () => {
    await tools.getByRole('button', { name: '生成预览', exact: true }).click();
    await expect(tools.getByRole('img', { name: '生成效果预览' })).toBeVisible();
    await expect(tools.getByRole('button', { name: '确认生成' })).toHaveCount(0);
    await expect(tools.getByRole('button', { name: '关闭', exact: true })).toHaveCount(0);
    await expect(tools.getByRole('status')).not.toContainText('积分');
  };
  const imageTab = editor.getByRole('button', { name: 'AI 图片工具', exact: true });
  await expect(editor.locator('.deckwise-teaching-sidebar [data-editor-image-tools]')).toHaveCount(0);
  await expect(imageTab).toBeVisible();
  await imageTab.click();
  await expect(editor.locator('.deckwise-teaching-sidebar')).toBeHidden();
  await tools.getByRole('button', { name: '插入 AI 图片艺术字' }).click();
  await expect(tools.locator('.image-presets button')).toHaveCount(14);
  await expect.poll(() => tools.locator('.image-presets img').first().evaluate((image: HTMLImageElement) => image.naturalWidth)).toBeGreaterThan(0);
  await tools.getByRole('button', { name: '绘本冒险字' }).click();
  await tools.getByLabel('艺术字文案').fill('快乐出发');
  await generatePreview();
  await tools.getByRole('button', { name: '应用到本页' }).click();
  await expect(tools.getByRole('status')).toHaveText('艺术字已插入');
  const save = () =>
    page.evaluate(async () => Array.from(new Uint8Array(await (await post('document:save')).file.arrayBuffer())));
  const bytes = Uint8Array.from(await save());
  const xml = await zipEntryText(bytes, 'ppt/slides/slide1.xml');
  expect(xml).toContain('deckwise-art-text/1');
  expect(xml).toContain('快乐出发');
  await open(toBase64(bytes));

  const selectImage = () =>
    page.evaluate(() => {
      const editor = (window.__ooFrames.readyEditor() as any).Asc.editor;
      const slide = editor.WordControl.m_oLogicDocument.Slides[0];
      const controller = editor.WordControl.m_oLogicDocument.GetCurrentController();
      const image = slide.cSld.spTree.find((shape: any) => shape.isImage?.());
      controller.resetSelection();
      controller.selectObject(image, 0);
      editor.WordControl.m_oLogicDocument.Document_UpdateInterfaceState();
      return image.Id as string;
    });
  const id = await selectImage();
  await imageTab.click();
  await expect(tools.getByLabel('艺术字文案')).toHaveValue('快乐出发');
  await expect(tools.getByRole('button', { name: '绘本冒险字' })).toHaveAttribute('aria-pressed', 'true');
  await tools.getByLabel('艺术字文案').fill('再次出发');
  await generatePreview();
  await tools.getByRole('button', { name: '应用到本页' }).click();
  await expect(tools.getByRole('status')).toHaveText('图片已替换');
  expect(await selectImage()).toBe(id);
  const editedXml = await zipEntryText(Uint8Array.from(await save()), 'ppt/slides/slide1.xml');
  expect(editedXml).toContain('再次出发');
  await page.evaluate(() => {
    const editor = (window.__ooFrames.readyEditor() as any).Asc.editor;
    const logic = editor.WordControl.m_oLogicDocument;
    const controller = logic.GetCurrentController();
    const image = controller.selectedObjects[0];
    image.setDescription('Ordinary picture');
    controller.resetSelection();
    logic.Document_UpdateInterfaceState();
  });
  await expect(tools.getByRole('button', { name: '重新生成图片' })).toBeHidden();
  expect(await selectImage()).toBe(id);
  await expect(tools.getByRole('button', { name: '重新生成图片' })).toHaveAttribute('aria-pressed', 'true');
  const reference = tools.getByRole('checkbox', { name: '参考当前图片' });
  await expect(reference).toBeChecked();
  await generatePreview();
  const latestQuote = () => page.evaluate(() =>
    (window as any).imageCalls.findLast((call: { operation: string }) => call.operation === 'editor-image:quote').args,
  );
  expect((await latestQuote()).reference).toMatch(/^data:image\/png;base64,/);
  await reference.uncheck();
  await tools.getByLabel('生图模型').selectOption('text-only');
  await generatePreview();
  expect(await latestQuote()).toMatchObject({ model: 'text-only' });
  expect(await latestQuote()).not.toHaveProperty('reference');
  await tools.getByRole('button', { name: '应用到本页' }).click();
  await expect(tools.getByRole('status')).toHaveText('图片已替换');
  expect(await selectImage()).toBe(id);

  const pictureState = () => page.evaluate(() => {
    const editor = (window.__ooFrames.readyEditor() as any).Asc.editor;
    return editor.WordControl.m_oLogicDocument.Slides[0].cSld.spTree
      .filter((shape: any) => shape.isImage?.())
      .map((shape: any) => ({
        id: shape.Id,
        rasterId: shape.blipFill.RasterImageId,
        description: shape.getCNvProps().descr,
      }));
  });
  const original = await pictureState();
  await tools.getByRole('button', { name: '插入 AI 图片艺术字' }).click();
  await expect(tools.getByLabel('艺术字文案')).toHaveValue('');
  await tools.getByRole('button', { name: '绘本冒险字' }).click();
  await tools.getByLabel('艺术字文案').fill('新的标题');
  await generatePreview();
  expect(await latestQuote()).not.toHaveProperty('reference');
  await tools.getByRole('button', { name: '应用到本页' }).click();
  await expect(tools.getByRole('status')).toHaveText('艺术字已插入');
  const pictures = await pictureState();
  expect(pictures).toHaveLength(original.length + 1);
  expect(pictures.slice(0, -1)).toEqual(original);
  expect(pictures.at(-1)?.description).toContain('新的标题');
  const submitted = await page.evaluate(() =>
    (window as any).imageCalls.filter((call: { operation: string }) => call.operation === 'editor-image:submit').length,
  );
  expect(submitted).toBe(5);
});
