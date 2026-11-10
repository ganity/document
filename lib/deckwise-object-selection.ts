import { requireEditorContext } from './agent-plugin/editor-bridge';
import { captureDocumentToolTarget } from './agent-plugin/document-tool-action';
import { readSlideShapeText, type SlideTextShape } from './agent-plugin/slide-text-read';

interface Shape extends SlideTextShape { group?: unknown; nvSpPr?: unknown; nvPicPr?: unknown;
  spPr?: { xfrm?: { offX?: number; offY?: number; extX?: number; extY?: number; rot?: number } }; }
interface SelectionApi {
  getCurrentPage?(): number;
  WordControl?: { m_oLogicDocument?: { Slides?: Array<{ cSld?: { spTree?: Shape[] } }>;
    GetWidthMM(): number; GetHeightMM(): number; GetCurrentController?(): { selectedObjects?: Shape[] } } };
}

/** Office regenerates cNvPr IDs on import. Return bounded fingerprints, never runtime IDs. */
export function readDeckWiseSelectedObjects() {
  try {
    const target = captureDocumentToolTarget();
    const api = requireEditorContext().api as unknown as SelectionApi;
    const logic = api.WordControl?.m_oLogicDocument, page = api.getCurrentPage?.();
    const shapes = page === undefined ? undefined : logic?.Slides?.[page]?.cSld?.spTree;
    const selected = logic?.GetCurrentController?.().selectedObjects;
    if (!logic || !shapes || !selected?.length || selected.length > 8 || new Set(selected).size !== selected.length || target.context.kind !== 'slide') return [];
    const width = logic.GetWidthMM(), height = logic.GetHeightMM();
    if (![width,height].every(n => Number.isFinite(n) && n > 0)) return [];
    const values = selected.map(shape => {
      if (shape.group || !shapes.includes(shape) || (!shape.nvSpPr && !shape.nvPicPr)) throw Error('Unsupported selection');
      const x = shape.spPr?.xfrm;
      if (!x || ![x.offX,x.offY,x.extX,x.extY].every(Number.isFinite) || x.extX! <= 0 || x.extY! <= 0) throw Error('Missing explicit geometry');
      const radians = x.rot ?? 0;
      if (!Number.isFinite(radians)) throw Error('Invalid rotation');
      const w = Math.abs(x.extX! * Math.cos(radians)) + Math.abs(x.extY! * Math.sin(radians));
      const h = Math.abs(x.extX! * Math.sin(radians)) + Math.abs(x.extY! * Math.cos(radians));
      const left = Math.max(0,x.offX! + (x.extX! - w)/2), top = Math.max(0,x.offY! + (x.extY! - h)/2);
      const right = Math.min(width,x.offX! + (x.extX! + w)/2), bottom = Math.min(height,x.offY! + (x.extY! + h)/2);
      if (right <= left || bottom <= top) throw Error('Outside page');
      const text = shape.nvPicPr ? '' : readSlideShapeText(shape);
      if (text === undefined) throw Error('Unverifiable text');
      const normalized = text.replace(/\s+/gu,'').slice(0,120);
      return { kind: shape.nvPicPr ? 'image' : normalized ? 'text' : 'shape', text: normalized,
        rect: { x: left/width, y: top/height, width: (right-left)/width, height: (bottom-top)/height } };
    });
    return target.isCurrent(false) ? values : [];
  } catch { return []; }
}
