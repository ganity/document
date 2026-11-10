import { requireEditorContext } from './editor-bridge';
import { readPresentationShapeText } from './slide-text-read';

/** Replace only the explicitly bound question; native history retains the original runs. */
export function prepareAssessmentText(page: number, drawing: any, before: string, after: string) {
  const { api, presentationBuilder: builder } = requireEditorContext() as any;
  const logic = api.WordControl.m_oLogicDocument, slide = logic.Slides[page - 1];
  const shape = builder?.GetPresentation().GetSlideByIndex(page - 1).GetAllShapes().find((s: any) => s.Drawing === drawing);
  const body = drawing.txBody?.createDuplicate?.();
  if (!shape || !body || readPresentationShapeText(drawing).trim() !== before.trim()) throw new Error('Question text binding changed');
  const paragraph = shape.GetDocContent().GetElement(0);
  const paragraphPr = paragraph.Paragraph.Pr.Copy();
  const textPr = paragraph.GetElement(0).GetTextPr();
  return { apply() {
    const content = shape.GetDocContent(); content.RemoveAllElements();
    after.split(/\r?\n/).forEach((line, index) => {
      const paragraph = index ? builder.CreateParagraph() : content.GetElement(0);
      paragraph.Paragraph.Set_Pr(paragraphPr.Copy());
      paragraph.AddText(line).SetTextPr(textPr); if (index) content.Push(paragraph);
    });
    if (readPresentationShapeText(drawing).trim() !== after.trim()) throw new Error('Question text readback mismatch');
    logic.Recalculate({ Drawings: { All: true, Map: {} } });
  }, restore() { drawing.setTxBody(body); logic.Recalculate({ Drawings: { All: true, Map: {} } }); } };
}
