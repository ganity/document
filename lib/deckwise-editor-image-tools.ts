import { captureDocumentToolTarget } from './agent-plugin/document-tool-action';
import { centeredSlideImageBox, prepareSlideImageInsertion } from './agent-plugin/slide-image-insertion';
import { validateSlidePng } from './agent-plugin/slide-image-png';
import { prepareSlideImageReplacement, requireSlideImage } from './agent-plugin/slide-image-replacement';
import { artStyles, artTextPrompt, type ArtStyle } from './deckwise-art-text-styles';
import { prepareEditorImage, readPageStyleReference, readSelectedImageReference, refinementReference } from './deckwise-editor-image-pixels';
import { markDocumentDirty } from './unsaved-guard';

const redrawPrompt = '参考当前图片，重新生成一张教学图片，保留原图的主体、内容、构图、画风和已有文字，提升清晰度，仅保留主体和已有文字、描边，背景透明，不添加无关内容。';
const standalonePrompt = 'Create exactly one standalone foreground asset. Match the current course style. Prioritize the requested edits. Preserve subject details, outlines, white areas and any existing text that the user did not ask to change. Render replacement text exactly. Use a perfectly flat solid chroma magenta #FF00FF background edge-to-edge for removal to transparent PNG. Do not use this key color in the subject. No unrelated objects, scenery, extra text, logo, watermark, gradient or detached shadow.';
const freshPrompt = '重新生成一张教学图片，仅保留主体和已有文字、描边，背景透明，不添加无关内容。';
type ArtText = { schema: 'deckwise-art-text/1'; text: string; style: ArtStyle; instruction: string; model: string };
type Mode = 'redraw' | 'insert' | 'edit';
type Model = { model: string; supportsReferenceImage?: boolean };
type Offer = { id: string; expiresAt: string };
type Status = { state: string; image?: string; errorCode?: string };

export const editorImageStyles = `
  [data-editor-image-tools] {
    position:absolute; inset:0 40px 0 0; z-index:20; overflow:auto; padding:14px;
    background:var(--background-normal,#fff); color:var(--text-normal,#222);
    font:13px/1.5 Arial,sans-serif;
    --image-line:var(--border-regular-control,#d0d5db);
    --image-accent:var(--text-accent,#267956);
  }
  [data-editor-image-tools] [hidden], [data-editor-image-tools][hidden] { display:none!important; }
  [data-editor-image-tools] h2 { margin:0 0 12px; font-size:16px; font-weight:600; color:inherit; }
  [data-editor-image-tools] label { display:block; margin:12px 0; font-weight:500; }
  [data-editor-image-tools] fieldset { border:0; display:block; min-width:0; margin:0; padding:0; }
  [data-editor-image-tools] textarea, [data-editor-image-tools] select {
    display:block; box-sizing:border-box; width:100%; margin-top:4px; padding:6px 8px;
    border:1px solid var(--image-line); border-radius:5px; font:inherit; color:inherit;
    background:var(--background-normal,#fff);
  }
  [data-editor-image-tools] textarea { resize:vertical; }
  [data-editor-image-tools] button {
    border:1px solid var(--image-line); border-radius:5px; margin:0 4px 4px 0; padding:5px 8px;
    background:transparent; color:inherit; font:inherit; cursor:pointer;
  }
  [data-editor-image-tools] button:disabled { opacity:.5; cursor:default; }
  [data-editor-image-tools] button:hover:not(:disabled) { background:var(--highlight-button-hover,#8882); }
  [data-editor-image-tools] button[aria-pressed="true"], [data-editor-image-tools] .image-primary {
    border-color:var(--image-accent); color:var(--image-accent);
  }
  [data-editor-image-tools] .image-actions { border-bottom:1px solid var(--image-line); padding-bottom:8px; }
  [data-editor-image-tools] .image-reference { display:flex; align-items:center; gap:6px; }
  [data-editor-image-tools] .image-reference input { margin:0; }
  [data-editor-image-tools] .image-presets {
    display:grid; grid-template-columns:repeat(2,minmax(0,1fr)); gap:5px;
    max-height:190px; overflow:auto; margin:4px 0 8px; padding:2px;
  }
  [data-editor-image-tools] .image-presets button { min-width:0; margin:0; padding:0 4px 4px; }
  [data-editor-image-tools] .image-presets img { display:block; width:100%; height:48px; object-fit:contain; }
  [data-editor-image-tools] .image-output {
    display:block; width:100%; max-height:190px; object-fit:contain; margin:8px 0;
    background:repeating-conic-gradient(#8883 0% 25%,transparent 0% 50%) 0 0 / 12px 12px;
  }
  [data-editor-image-tools] :focus-visible { outline:2px solid var(--image-accent); outline-offset:1px; }
  [data-editor-image-tools] [role=status] { margin:6px 0; overflow-wrap:anywhere; }
`;

function artTextFromDescription(description?: string): ArtText | undefined {
  if (!description || description.length > 4000) return;
  try {
    const data = JSON.parse(description) as ArtText;
    const style = ['ocean', 'space', 'fabric', 'blocks'].includes(data.style) ? 'kg-' + data.style : data.style;
    if (
      data.schema === 'deckwise-art-text/1' &&
      typeof data.text === 'string' &&
      Array.from(data.text).length <= 80 &&
      Object.hasOwn(artStyles, style) &&
      typeof data.model === 'string' &&
      typeof data.instruction === 'string' &&
      data.instruction.length <= 600
    )
      return { ...data, style: style as ArtStyle };
  } catch {
    // An ordinary picture can have a plain-text description.
  }
}

export function createDeckwiseImageTools(
  doc: Document,
  request: (operation: string, args?: Record<string, unknown>) => Promise<any>,
) {
  const create = <K extends keyof HTMLElementTagNameMap>(tag: K, text?: string) => {
    const node = doc.createElement(tag);
    if (text) node.textContent = text;
    return node;
  };
  const button = (text: string, action: () => void) => {
    const node = create('button', text);
    node.type = 'button';
    node.addEventListener('click', action);
    return node;
  };
  const label = (text: string, control: HTMLElement) => {
    const node = create('label', text);
    node.append(control);
    return node;
  };

  const section = create('aside');
  section.dataset.editorImageTools = '';
  section.setAttribute('aria-label', 'AI 图片工具');
  const design = create('div');
  design.className = 'image-design';
  const form = create('div');
  const text = create('textarea');
  text.maxLength = 160;
  text.rows = 2;
  text.placeholder = '例如：一起去探索！';
  const styleButtons = new Map<ArtStyle, HTMLButtonElement>();
  const presets = create('div');
  presets.className = 'image-presets';
  presets.setAttribute('aria-label', '推荐风格列表');
  for (const [id, style] of Object.entries(artStyles) as [ArtStyle, (typeof artStyles)[ArtStyle]][]) {
    const option = button(style.name, () => {
      selectedStyle = id;
      syncModels();
      reset();
      if (id === 'custom') prompt.focus();
    });
    option.title = style.description;
    if (id.startsWith('kg-')) {
      const sample = create('img');
      sample.alt = '';
      sample.loading = 'lazy';
      sample.src = new URL(`./ai-art-text/styles/${id}-transparent.webp`, location.href).href;
      option.prepend(sample);
      presets.append(option);
    } else {
      design.append(option);
      if (id === 'course') design.append(presets);
    }
    styleButtons.set(id, option);
  }
  const prompt = create('textarea');
  prompt.rows = 4;
  const promptLabel = create('span');
  const instruction = label('', prompt);
  instruction.prepend(promptLabel);
  const models = create('select');
  const referenceCurrent = create('input');
  referenceCurrent.type = 'checkbox';
  const referenceLabel = label('', referenceCurrent);
  const referenceText = create('span');
  referenceLabel.append(referenceText);
  referenceLabel.className = 'image-reference';
  const refinement = create('textarea');
  refinement.rows = 2;
  refinement.maxLength = 600;
  refinement.placeholder = '说明希望修改的颜色、材质或装饰';
  const results = create('div');
  results.className = 'image-versions';
  results.setAttribute('aria-label', '生成候选');
  const fields = create('fieldset');
  const preview = create('img');
  preview.alt = '生成效果预览';
  preview.className = 'image-output';
  preview.hidden = true;
  const notice = create('p');
  notice.setAttribute('role', 'status');
  const words = label('艺术字文案', text);
  const title = create('h2', 'AI 图片工具');
  let selectedStyle: ArtStyle = 'course';
  let choices: Model[] = [];
  let mode: Mode;
  let target: ReturnType<typeof captureDocumentToolTarget>;
  let aspect: number | undefined;
  let job: { id: string; art?: ArtText } | undefined;
  type Candidate = { image: Blob; url: string; art?: ArtText };
  let selected: Candidate | undefined;
  let candidates: Candidate[] = [];
  let polling: ReturnType<typeof setTimeout> | undefined;
  let busy = false;
  let disposed = false;
  let lastPicture: ReturnType<typeof requireSlideImage>['image'] | undefined;
  let lastPictureMode: 'redraw' | 'edit' | undefined;
  function clearCandidates() {
    for (const { url } of candidates) URL.revokeObjectURL(url);
    candidates = [];
    selectCandidate(undefined);
  }
  function selectCandidate(candidate: Candidate | undefined) {
    selected = candidate;
    if (candidate) preview.src = candidate.url;
    else preview.removeAttribute('src');
    preview.hidden = !candidate;
    results.replaceChildren(
      ...candidates.map((item, index) => {
        const pick = button(`版本 ${candidates.length - index}`, () => {
          reset();
          selectCandidate(item);
          sync();
        });
        pick.setAttribute('aria-pressed', String(item === selected));
        return pick;
      }),
    );
  }
  function reset(clear = false) {
    clearTimeout(polling);
    job = undefined;
    if (clear) clearCandidates();
    notice.textContent = '';
    sync();
  }
  function needsReference() {
    return mode === 'redraw' ? referenceCurrent.checked : selectedStyle === 'course' || (mode === 'edit' && referenceCurrent.checked);
  }
  function syncModels(preferred = models.value) {
    const options = choices
      .filter((model) => (!needsReference() && !(mode !== 'redraw' && refinement.value.trim() && selected)) || model.supportsReferenceImage)
      .map((model) => new Option(model.model, model.model));
    models.replaceChildren(...options);
    if (options.some((option) => option.value === preferred)) models.value = preferred;
  }
  function current() {
    if (disposed || !target?.isCurrent(true)) throw new Error('原页面或选中对象已变化，请重新选择');
  }
  function sync() {
    const pending = busy || !!job;
    fields.disabled = pending;
    redraw.disabled = insert.disabled = edit.disabled = pending;
    redraw.setAttribute('aria-pressed', String(!form.hidden && mode === 'redraw'));
    insert.setAttribute('aria-pressed', String(!form.hidden && mode === 'insert'));
    edit.setAttribute('aria-pressed', String(!form.hidden && mode === 'edit'));
    generate.disabled = pending || !models.value;
    refresh.hidden = !job;
    refresh.disabled = busy;
    apply.hidden = !selected;
    apply.disabled = pending || (!!selected?.art && text.value.trim() !== selected.art.text);
    refinementBox.hidden = mode === 'redraw' || !selected?.art;
    referenceLabel.hidden = mode !== 'redraw' && (mode !== 'edit' || selectedStyle === 'course');
    for (const [id, option] of styleButtons) option.setAttribute('aria-pressed', String(id === selectedStyle));
    for (const pick of results.querySelectorAll('button')) pick.disabled = pending;
  }
  async function run(action: () => Promise<void>) {
    if (busy || disposed) return;
    busy = true;
    sync();
    try {
      await action();
    } catch (error) {
      if (!target?.isCurrent(true)) {
        reset(true);
        form.hidden = true;
      }
      notice.textContent = error instanceof Error ? error.message : String(error);
    } finally {
      busy = false;
      sync();
      updateSelection();
    }
  }
  const redraw = button('重新生成图片', () => begin('redraw'));
  const insert = button('插入 AI 图片艺术字', () => begin('insert'));
  const edit = button('改艺术字文案 / 风格', () => begin('edit'));

  function begin(next: Mode) {
    void run(async () => {
      const picture = next === 'insert' ? undefined : requireSlideImage();
      const saved = next === 'edit' ? artTextFromDescription(picture?.image.getCNvProps?.()?.descr) : undefined;
      if (next === 'edit' && !saved) throw new Error('请选择已生成的 AI 图片艺术字');
      if (next === 'insert') centeredSlideImageBox(1, 1);
      aspect = picture ? Number(picture.image.spPr!.xfrm!.extX) / Number(picture.image.spPr!.xfrm!.extY) : undefined;
      target = captureDocumentToolTarget();
      current();
      mode = next;
      choices = [];
      models.replaceChildren();
      reset(true);
      form.hidden = false;
      words.hidden = design.hidden = mode === 'redraw';
      text.value = saved?.text ?? '';
      selectedStyle = saved?.style ?? 'course';
      referenceCurrent.checked = mode === 'redraw';
      referenceText.textContent = mode === 'redraw' ? '参考当前图片' : '沿用原图片的配色与字形风格';
      promptLabel.textContent = mode === 'redraw' ? '生图提示词' : '补充要求';
      prompt.maxLength = mode === 'redraw' ? 4000 : 600;
      prompt.value = saved?.instruction ?? (mode === 'redraw' ? redrawPrompt : '');
      refinement.value = '';
      sync();
      notice.textContent = '正在读取 DeckWise 生图模型…';
      choices = await request('editor-image:models');
      current();
      syncModels(saved?.model);
      notice.textContent = '';
    });
  }

  const generate = button(
    '生成预览',
    () =>
      void run(async () => {
        current();
        const refine = mode === 'redraw' ? '' : refinement.value.trim();
        const previous = selected?.image;
        reset(mode === 'redraw');
        const wording = text.value.trim();
        const art: ArtText | undefined =
          mode === 'redraw'
            ? undefined
            : {
                schema: 'deckwise-art-text/1',
                text: wording,
                style: selectedStyle,
                instruction: prompt.value.trim(),
                model: models.value,
              };
        if (art && (!art.text || Array.from(art.text).length > 80 || art.text.split('\n').length > 4))
          throw new Error('请输入 1–80 字、最多 4 行的艺术字文案');
        if (art?.style === 'custom' && !art.instruction) throw new Error('请输入自定义风格要求');
        const description = prompt.value.trim();
        if (!art && !description) throw new Error('请输入图片修改要求');
        notice.textContent = '正在生成图片…';
        const reference = refine && previous
          ? await refinementReference(previous)
          : art?.style === 'course'
            ? readPageStyleReference()
            : needsReference() ? await readSelectedImageReference() : undefined;
        const input = art
          ? artTextPrompt(art.text, art.style, art.instruction, !!reference, previous ? refine : '')
          : description + '\n' + standalonePrompt;
        if (input.length > 4000) throw new Error(`生成要求过长，请缩短至 ${4000 - standalonePrompt.length - 1} 字以内`);
        current();
        const offer: Offer = await request('editor-image:quote', {
          model: models.value,
          prompt: input,
          ...(reference ? { reference } : {}),
        });
        current();
        if (Date.parse(offer.expiresAt) <= Date.now()) throw new Error('生成请求已过期，请重试');
        job = { id: offer.id, art };
        await update('submit');
      }),
  );

  async function update(operation: 'submit' | 'status') {
    current();
    if (!job) return;
    const result: Status = await request(`editor-image:${operation}`, { id: job.id });
    current();
    if (result.state === 'ready' && result.image) {
      const { art } = job;
      job = undefined;
      const image = await prepareEditorImage(result.image, !!art, aspect);
      current();
      refinement.value = '';
      const candidate = { image, url: URL.createObjectURL(image), art };
      candidates.unshift(candidate);
      for (const { url } of candidates.splice(6)) URL.revokeObjectURL(url);
      selectCandidate(candidate);
      notice.textContent = art ? '请核对文案和笔画。' : '预览已生成';
    } else if (['failed', 'cancelled', 'archive_failed'].includes(result.state)) {
      job = undefined;
      notice.textContent = `生成未完成${result.errorCode ? ' · ' + result.errorCode : ''}`;
    } else {
      notice.textContent = result.state === 'submission_unknown' ? '正在确认提交状态…' : '正在生成图片…';
      clearTimeout(polling);
      polling = setTimeout(() => void run(() => update('status')), 3000);
    }
  }
  const refresh = button('刷新生成状态', () => void run(() => update('status')));
  const apply = button(
    '应用到本页',
    () =>
      void run(async () => {
        current();
        if (!selected || (selected.art && text.value.trim() !== selected.art.text)) return;
        const bytes = new Uint8Array(await selected.image.arrayBuffer());
        current();
        const dimensions = validateSlidePng(bytes);
        const metadata = selected.art ? JSON.stringify(selected.art) : undefined;
        const prepared =
          mode === 'insert'
            ? await prepareSlideImageInsertion(
                bytes,
                centeredSlideImageBox(dimensions.width, dimensions.height),
                undefined,
                { description: metadata, allowSelectedObjects: true },
              )
            : await prepareSlideImageReplacement(bytes, undefined, metadata);
        current();
        prepared.apply();
        markDocumentDirty();
        form.hidden = true;
        reset(true);
        updateSelection('applied');
        notice.textContent = mode === 'insert' ? '艺术字已插入' : '图片已替换';
      }),
  );
  const refinementBox = label('基于当前版本继续修改', refinement);
  for (const field of [text, prompt, models, referenceCurrent, refinement]) {
    field.addEventListener('input', () => {
      if (job) return;
      if (field === referenceCurrent && mode === 'redraw') {
        if (prompt.value === redrawPrompt || prompt.value === freshPrompt)
          prompt.value = referenceCurrent.checked ? redrawPrompt : freshPrompt;
      }
      if (field === referenceCurrent || field === refinement) syncModels();
      reset(mode === 'redraw');
    });
  }
  fields.append(words, design, instruction, label('生图模型', models), referenceLabel, refinementBox);
  const actions = create('div');
  actions.className = 'image-actions';
  actions.append(redraw, insert, edit);
  generate.className = 'image-primary';
  form.append(fields, generate, refresh, preview, results, apply);
  form.hidden = true;
  section.append(title, actions, form, notice);
  function updateSelection(reason: 'selection' | 'opened' | 'applied' = 'selection') {
    if (disposed || (section.hidden && reason !== 'applied') || (busy && reason !== 'applied')) return;
    const staleForm = !form.hidden && target && !target.isCurrent(true);
    if (staleForm) {
      form.hidden = true;
      reset(true);
    }
    let picture: typeof lastPicture;
    try {
      picture = requireSlideImage().image;
    } catch {
      // The selection is not a supported picture.
    }
    const art = artTextFromDescription(picture?.getCNvProps?.()?.descr);
    const matchingMode = picture ? (art ? 'edit' : 'redraw') : undefined;
    const selectionChanged = picture !== lastPicture || matchingMode !== lastPictureMode;
    lastPicture = picture;
    lastPictureMode = matchingMode;
    redraw.hidden = matchingMode !== 'redraw';
    edit.hidden = matchingMode !== 'edit';
    if (
      reason !== 'applied' &&
      matchingMode &&
      (reason === 'opened' || selectionChanged || staleForm) &&
      (form.hidden || mode !== matchingMode)
    )
      begin(matchingMode);
  }
  return {
    element: section,
    updateSelection,
    dispose() {
      disposed = true;
      clearTimeout(polling);
      clearCandidates();
      section.remove();
    },
  };
}
