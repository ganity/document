import { requireEditorContext } from './editor-bridge';
import { getReadonlyMode } from '../onlyoffice/readonly';
import { readSlideShapeText } from './slide-text-read';

type TextElement = {
  id: string;
  type: 'native_text';
  bboxPx: number[];
  content: string;
  style: { fontSizePt: number; fontWeight: number; color: string; fontFace: string };
};
type ShapeElement = { id: string; type: 'native_shape'; bboxPx: number[]; fill: string };
export type AssessmentProjection = {
  slideId: string;
  canvas: { width: number; height: number };
  elements: Array<TextElement | ShapeElement>;
};

/** Publish a detached assessment with the carrier in the caller's one native history action. */
export function prepareAssessmentInsertion(after: number, projection: AssessmentProjection) {
  const { api, presentationBuilder: builder, AscCommon: common } = requireEditorContext() as any;
  const logic = api.WordControl?.m_oLogicDocument;
  const before = [...(logic?.Slides ?? [])];
  if (
    getReadonlyMode() ||
    api.canEdit?.() !== true ||
    logic?.CanEdit?.() !== true ||
    logic.IsMasterMode?.() !== false ||
    api.isSlideShow?.() ||
    common?.CollaborativeEditing?.Get_GlobalLock() ||
    !builder ||
    typeof logic.insertSlide !== 'function' ||
    typeof logic.removeSlide !== 'function' ||
    after < 0 ||
    after >= before.length ||
    common?.History?._getLongPointIndex?.() !== -1
  )
    throw new Error('Native assessment insertion is unavailable');
  if (
    !projection ||
    !projection.canvas ||
    !Array.isArray(projection.elements) ||
    !projection.elements.length ||
    projection.elements.length > 64 ||
    ![projection.canvas.width, projection.canvas.height].every((n) => Number.isFinite(n) && n > 0) ||
    projection.elements.some(
      (e) =>
        !['native_text', 'native_shape'].includes(e.type) ||
        !e.id.startsWith(projection.slideId + '_') ||
        !Array.isArray(e.bboxPx) ||
        e.bboxPx.length !== 4 ||
        !e.bboxPx.every(Number.isFinite) ||
        (e.type === 'native_text' &&
          (typeof e.content !== 'string' ||
            e.content.length > 8000 ||
            !e.style ||
            !Number.isFinite(e.style.fontSizePt) ||
            e.style.fontSizePt <= 0 ||
            !/^#[a-f0-9]{6}$/i.test(e.style.color))) ||
        (e.type === 'native_shape' && !/^#[a-f0-9]{6}$/i.test(e.fill)),
    )
  )
    throw new Error('Invalid native assessment projection');
  const assertCurrent = () => {
    if (
      getReadonlyMode() ||
      logic.Slides.length !== before.length ||
      before.some((slide, i) => logic.Slides[i] !== slide)
    )
      throw new Error('Assessment insertion target changed');
  };
  let inserted: any;
  return {
    assertCurrent,
    apply() {
      assertCurrent();
      const page = builder.CreateSlide();
      page.RemoveAllObjects();
      const sx = logic.GetWidthMM() / projection.canvas.width,
        sy = logic.GetHeightMM() / projection.canvas.height;
      const rgb = (value: string) =>
        value
          .slice(1)
          .match(/../g)!
          .map((v) => parseInt(v, 16));
      for (const element of projection.elements) {
        const [x, y, w, h] = element.bboxPx;
        const shape = builder.CreateShape(
          element.type === 'native_shape' ? 'roundRect' : 'rect',
          w * sx * 36000,
          h * sy * 36000,
          element.type === 'native_shape'
            ? builder.CreateSolidFill(builder.CreateRGBColor(...rgb(element.fill)))
            : builder.CreateNoFill(),
          builder.CreateStroke(0, builder.CreateNoFill()),
        );
        shape.SetPosition(x * sx * 36000, y * sy * 36000);
        if (element.type === 'native_text') {
          shape.SetVerticalTextAlign('top');
          shape.SetPaddings(0, 0, 0, 0);
          const content = shape.GetDocContent();
          content.RemoveAllElements();
          element.content.split(/\r?\n/u).forEach((line, i) => {
            const paragraph = i ? builder.CreateParagraph() : content.GetElement(0);
            paragraph.AddText(line);
            paragraph.SetFontSize(element.style.fontSizePt * 2);
            paragraph.SetBold(element.style.fontWeight >= 600);
            paragraph.SetColor(...rgb(element.style.color));
            paragraph.SetFontFamily(element.style.fontFace);
            if (i) content.Push(paragraph);
          });
        }
        shape.Drawing.setName(element.id);
        if (!page.AddObject(shape)) throw new Error('Assessment object insertion failed');
        if (
          element.type === 'native_text' &&
          readSlideShapeText(shape.Drawing)?.trimEnd() !== element.content.trimEnd()
        )
          throw new Error('Assessment text readback differs');
      }
      inserted = page.Slide;
      logic.insertSlide(after + 1, inserted);
      if (logic.Slides[after + 1] !== inserted) throw new Error('Assessment page insertion could not be verified');
      logic.Recalculate({ Drawings: { All: true, Map: {} } });
      logic.Document_UpdateInterfaceState();
      logic.DrawingDocument.UpdateThumbnailsAttack();
      return inserted;
    },
    restore() {
      const index = logic.Slides.indexOf(inserted);
      if (index >= 0) logic.removeSlide(index);
      logic.Recalculate({ Drawings: { All: true, Map: {} } });
    },
  };
}
