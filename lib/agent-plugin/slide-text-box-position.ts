import type { AgentTool } from '@ranuts/agent-core/types';
import { requireEditorContext } from './editor-bridge';
import { readSlideShapeText } from './slide-text-read';
import { requireSlideTextSelection, type SlideTextApi } from './slide-text-selection';

interface PositionShape {
  x: number;
  y: number;
  extX: number;
  extY: number;
  rot: number;
  flipH?: boolean;
  flipV?: boolean;
  group?: unknown;
  isPlaceholder?(): boolean;
  getObjectType?(): number;
  spPr?: { geometry?: { preset?: string } };
}
type PositionApi = SlideTextApi & {
  ImgApply?(props: { Position?: { X: number; Y: number }; Width?: number; Height?: number }): void;
  WordControl?: { m_oLogicDocument?: { GetWidthMM?(): number; GetHeightMM?(): number } };
};
export function requireSlideTextBox() {
  const context = requireEditorContext();
  const api = context.api as PositionApi;
  const selected = requireSlideTextSelection(api);
  const shape = selected.shape as typeof selected.shape & PositionShape;
  const logic = api.WordControl?.m_oLogicDocument;
  const width = logic?.GetWidthMM?.();
  const height = logic?.GetHeightMM?.();
  const shapeType = Reflect.get(context.AscDFH ?? {}, 'historyitem_type_Shape');
  const imageType = Reflect.get(context.AscDFH ?? {}, 'historyitem_type_ImageShape');
  if (
    typeof api.ImgApply !== 'function' ||
    !Number.isSafeInteger(shapeType) ||
    shape.getObjectType?.() !== shapeType ||
    shape.isPlaceholder?.() !== false ||
    shape.group ||
    shape.rot !== 0 ||
    shape.flipH ||
    shape.flipV ||
    shape.spPr?.geometry?.preset !== 'rect' ||
    selected.slide.cSld.spTree.some((item) => {
      const peer = item as PositionShape;
      return (
        peer.getObjectType?.() !== shapeType &&
        (!Number.isSafeInteger(imageType) ||
          peer.getObjectType?.() !== imageType ||
          peer.group ||
          peer.isPlaceholder?.() !== false)
      );
    }) ||
    ![shape.x, shape.y, shape.extX, shape.extY, width, height].every(
      (value) => typeof value === 'number' && Number.isFinite(value),
    ) ||
    shape.extX <= 0 ||
    shape.extY <= 0 ||
    width! <= 0 ||
    height! <= 0 ||
    shape.x < 0 ||
    shape.y < 0 ||
    shape.x + shape.extX > width! ||
    shape.y + shape.extY > height!
  )
    throw new Error(
      'This command requires a supported unrotated, ungrouped, on-page rectangular text box on a slide containing only shapes and ordinary images',
    );
  return { ...selected, shape, width: width!, height: height! };
}
export function canSetSlideTextBoxPosition(): boolean {
  try {
    requireSlideTextBox();
    return true;
  } catch {
    return false;
  }
}
export const setSlideTextBoxPositionTool: AgentTool<
  { xMm: number; yMm: number },
  { xMm: number; yMm: number; page: number; verified: true }
> = {
  name: 'set_slide_text_box_position',
  description:
    'PPT only. Move the entire rectangular text box containing the current nonempty text selection to explicit top-left millimeter coordinates. Ordinary peer images are preserved; no connectors, groups, resize, rotated or placeholder text boxes, or off-page positions. Does not select text or save.',
  inputSchema: {
    type: 'object',
    properties: {
      xMm: { type: 'number', minimum: 0, maximum: 1000 },
      yMm: { type: 'number', minimum: 0, maximum: 1000 },
    },
    required: ['xMm', 'yMm'],
    additionalProperties: false,
  },
  readOnlyHint: false,
  execute: async ({ xMm, yMm }, signal) => {
    if (
      ![xMm, yMm].every((value) => typeof value === 'number' && Number.isFinite(value) && value >= 0 && value <= 1000)
    )
      throw new TypeError('Position must use finite millimeter coordinates between 0 and 1000');
    signal?.throwIfAborted();
    const before = requireSlideTextBox();
    const { extX, extY } = before.shape;
    const text = readSlideShapeText(before.shape);
    if (text === undefined || xMm + extX > before.width || yMm + extY > before.height)
      throw new Error('The entire text box must fit on the current slide');
    before.api.ImgApply!({ Position: { X: xMm, Y: yMm } });
    signal?.throwIfAborted();
    const after = requireSlideTextBox();
    if (
      after.api !== before.api ||
      after.slide !== before.slide ||
      after.shape !== before.shape ||
      after.width !== before.width ||
      after.height !== before.height ||
      after.shape.x !== xMm ||
      after.shape.y !== yMm ||
      after.shape.extX !== extX ||
      after.shape.extY !== extY ||
      JSON.stringify(after.selected) !== JSON.stringify(before.selected) ||
      readSlideShapeText(after.shape) !== text
    )
      throw new Error(
        'Text box position could not be verified. Inspect the document and use Undo if needed; do not retry automatically.',
      );
    return { xMm, yMm, page: before.page + 1, verified: true };
  },
};
