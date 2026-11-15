import { requireEditorContext, getEditorContext } from './editor-bridge';
import { preloadSlidePng, validateSlidePng } from './slide-image-png';
import { readSlideShapeText } from './slide-text-read';
import { getReadonlyMode } from '../onlyoffice/readonly';

type Projection = { page: number; previous: any; next: any; assets: Record<string, Uint8Array> };
/** Prepare images before opening the native teaching transaction. */
export async function prepareQuizProjection({ page, previous, next, assets }: Projection) {
  const native = requireEditorContext() as any,
    { api, presentationBuilder: builder, AscCommon: common } = native;
  const logic = api.WordControl?.m_oLogicDocument,
    slide = logic?.Slides?.[page - 1],
    history = common?.History;
  if (
    getReadonlyMode() ||
    !slide ||
    !builder ||
    !history ||
    api.canEdit?.() !== true ||
    logic.CanEdit?.() !== true ||
    logic.IsMasterMode?.() !== false ||
    api.isSlideShow?.() ||
    common.CollaborativeEditing?.Get_GlobalLock() ||
    history._getLongPointIndex?.() !== -1 ||
    logic.IsSelectionLocked(common.changestype_Drawing_Props)
  )
    throw new Error('Quiz page is not writable');
  if (
    !previous ||
    !next ||
    previous.slideId !== next.slideId ||
    !Array.isArray(previous.elements) ||
    !Array.isArray(next.elements) ||
    next.elements.length > 64 ||
    previous.elements.length > 64
  )
    throw new Error('Invalid Quiz projection');
  const originals: any[] = [...slide.cSld.spTree],
    byName = new Map<string, any>(),
    owners = new Map<any, any>();
  const trees: Array<{ tree: any[]; shapes: any[]; parent?: any }> = [];
  const collect = (tree: any[], owner?: any, parent?: any) => {
    trees.push({ tree, shapes: [...tree], parent });
    for (const shape of tree) {
      const name = shape.getObjectName();
      if (byName.has(name)) throw new Error('Duplicate native Quiz object name');
      byName.set(name, shape);
      owners.set(shape, owner ?? shape);
      if (shape.spTree) collect(shape.spTree, owner ?? shape, shape);
    }
  };
  collect(slide.cSld.spTree);
  // The persisted logical page ID can differ from the original assessment template prefix.
  const anchor = previous.elements.find((e: any) => e.type === 'native_text' && e.id.endsWith('_q0_question'));
  const anchorText = (text: string | undefined) => text?.replace(/[\r\n]/g, '').trim();
  const anchors = [...byName.values()].filter(
    (s) =>
      s.getObjectName().endsWith('_q0_question') && anchorText(readSlideShapeText(s)) === anchorText(anchor?.content),
  );
  const prefix = anchors.length === 1 ? anchors[0].getObjectName().replace(/_q0_question$/, '') : previous.slideId;
  const actualName = (id: string) => prefix + id.slice(previous.slideId.length);
  for (const element of previous.elements)
    if (!byName.has(actualName(element.id))) throw new Error('Quiz object binding changed; refresh before saving');
  const bound = previous.elements.map((e: any) => byName.get(actualName(e.id)));
  const group = bound.map((shape: any) => owners.get(shape)).find((owner: any) => owner.spTree);
  // Rebuild an exported assessment group only when every leaf belongs to this Quiz.
  if (
    group &&
    (bound.some((shape: any) => owners.get(shape) !== group || shape.spTree) ||
      [...owners].some(([shape, owner]) => owner === group && shape !== group && !bound.includes(shape)))
  )
    throw new Error('Quiz object binding includes unrelated group content; refresh before saving');
  const changed = next.elements.filter(
    (e: any) =>
      group ||
      JSON.stringify(e) !== JSON.stringify(previous.elements.find((p: any) => p.id === e.id)) ||
      e.type === 'image_layer',
  );
  const replace = new Set([
    ...changed.map((e: any) => actualName(e.id)),
    ...previous.elements
      .filter((e: any) => !next.elements.some((n: any) => n.id === e.id))
      .map((e: any) => actualName(e.id)),
  ]);
  const before = history.Index,
    point = history.Points?.[before];
  const assertCurrent = () => {
    if (
      getReadonlyMode() ||
      getEditorContext()?.api !== api ||
      logic.Slides[page - 1] !== slide ||
      history.Index !== before ||
      history.Points?.[before] !== point ||
      slide.cSld.spTree !== trees[0].tree ||
      trees.some(
        ({ tree, shapes, parent }) =>
          (parent && parent.spTree !== tree) || tree.length !== shapes.length || shapes.some((s, i) => tree[i] !== s),
      )
    )
      throw new Error('Quiz page changed during image preparation');
  };
  const scaleX = logic.GetWidthMM() / next.canvas.width,
    scaleY = logic.GetHeightMM() / next.canvas.height;
  if (![scaleX, scaleY].every((v) => Number.isFinite(v) && v > 0)) throw new Error('Invalid Quiz canvas');
  const images = new Map<string, Awaited<ReturnType<typeof preloadSlidePng>>>();
  for (const element of changed) {
    if (
      !['native_text', 'native_shape', 'image_layer'].includes(element.type) ||
      typeof element.id !== 'string' ||
      !element.id.startsWith(`${next.slideId}_`) ||
      !Array.isArray(element.bboxPx) ||
      element.bboxPx.length !== 4 ||
      !element.bboxPx.every(Number.isFinite)
    )
      throw new Error('Invalid live Quiz element');
    if (element.type === 'image_layer') {
      const bytes = assets[element.assetId];
      images.set(element.id, await preloadSlidePng(bytes, { api, common }, validateSlidePng(bytes)));
      assertCurrent();
    }
  }
  assertCurrent();
  const created = new Map<string, any>();
  return {
    assertCurrent,
    apply() {
      const currentSlide = builder.GetPresentation().GetSlideByIndex(page - 1);
      const rgb = (value: string) =>
        builder.CreateRGBColor(
          ...value
            .replace('#', '')
            .match(/../g)!
            .map((v) => parseInt(v, 16)),
        );
      for (const element of changed) {
        const [x, y, w, h] = element.bboxPx;
        let drawing: any;
        if (element.type === 'image_layer') {
          const image = images.get(element.id)!;
          image.ready();
          drawing = logic
            .GetCurrentController()
            .createImage(image.full, x * scaleX, y * scaleY, w * scaleX, h * scaleY);
        } else {
          const shape = builder.CreateShape(
            element.type === 'native_shape' ? 'roundRect' : 'rect',
            w * scaleX * 36000,
            h * scaleY * 36000,
            element.type === 'native_shape' ? builder.CreateSolidFill(rgb(element.fill)) : builder.CreateNoFill(),
            builder.CreateStroke(0, builder.CreateNoFill()),
          );
          shape.SetPosition(x * scaleX * 36000, y * scaleY * 36000);
          drawing = shape.Drawing;
          if (element.type === 'native_text') {
            shape.SetVerticalTextAlign('top');
            shape.SetPaddings(0, 0, 0, 0);
            const content = shape.GetDocContent();
            content.RemoveAllElements();
            element.content.split(/\r?\n/u).forEach((line: string, i: number) => {
              const paragraph = i ? builder.CreateParagraph() : content.GetElement(0);
              paragraph.AddText(line);
              paragraph.SetFontSize(element.style.fontSizePt * 2);
              paragraph.SetBold(element.style.fontWeight >= 600);
              paragraph.SetColor(
                ...element.style.color
                  .slice(1)
                  .match(/../g)
                  .map((v: string) => parseInt(v, 16)),
              );
              paragraph.SetFontFamily(element.style.fontFace);
              if (i) content.Push(paragraph);
            });
          }
          if (!currentSlide.AddObject(shape)) throw new Error('Quiz object insertion failed');
          slide.removeFromSpTreeByPos(slide.cSld.spTree.indexOf(drawing));
        }
        drawing.setParent(slide);
        drawing.setName(actualName(element.id));
        created.set(element.id, drawing);
      }
      for (let i = originals.length - 1; i >= 0; i--)
        if (originals[i] === group || replace.has(originals[i].getObjectName())) slide.removeFromSpTreeByPos(i);
      for (const element of next.elements) {
        const drawing = created.get(element.id);
        if (!drawing) continue;
        const following = next.elements
          .slice(next.elements.indexOf(element) + 1)
          .map((e: any) => byName.get(actualName(e.id)))
          .find((s: any) => slide.cSld.spTree.includes(s));
        slide.addToSpTreeToPos(
          group
            ? originals.indexOf(group) + next.elements.indexOf(element)
            : following
              ? slide.cSld.spTree.indexOf(following)
              : slide.cSld.spTree.length,
          drawing,
        );
      }
      logic.Recalculate({ Drawings: { All: true, Map: {} } });
      if (
        next.elements.some(
          (e: any) =>
            created.has(e.id) &&
            (!slide.cSld.spTree.includes(created.get(e.id)) ||
              (e.type === 'native_text' &&
                readSlideShapeText(created.get(e.id))?.replace(/\r\n?/g, '\n').trimEnd() !==
                  e.content.replace(/\r\n?/g, '\n').trimEnd())),
        )
      )
        throw new Error('Native Quiz update readback mismatch');
      logic.Document_UpdateInterfaceState();
      logic.DrawingDocument.UpdateThumbnailsAttack();
    },
    restore() {
      for (let i = slide.cSld.spTree.length - 1; i >= 0; i--) slide.removeFromSpTreeByPos(i);
      originals.forEach((s, i) => {
        s.setParent(slide);
        slide.addToSpTreeToPos(i, s);
      });
      logic.Recalculate({ Drawings: { All: true, Map: {} } });
    },
  };
}
