import type { AgentTool } from '@ranuts/agent-core/types';
import { requireSlideTextBox } from './slide-text-box-position';
import { readSlideShapeText } from './slide-text-read';

export function canSetSlideTextBoxSize(): boolean {
  try {
    requireSlideTextBox();
    return true;
  } catch {
    return false;
  }
}

export const setSlideTextBoxSizeTool: AgentTool<
  { widthMm: number; heightMm: number },
  { widthMm: number; heightMm: number; page: number; verified: true }
> = {
  name: 'set_slide_text_box_size',
  description:
    'PPT only. Resize the entire rectangular text box containing the current nonempty text selection to explicit millimeter width and height, preserving its top-left position. Only ordinary unrotated, ungrouped, on-page text boxes; no placeholders or connectors. Text may reflow or overflow: review visually. Does not select text or save.',
  inputSchema: {
    type: 'object',
    properties: {
      widthMm: { type: 'number', minimum: 1, maximum: 1000 },
      heightMm: { type: 'number', minimum: 1, maximum: 1000 },
    },
    required: ['widthMm', 'heightMm'],
    additionalProperties: false,
  },
  readOnlyHint: false,
  execute: async ({ widthMm, heightMm }, signal) => {
    if (
      ![widthMm, heightMm].every(
        (value) => typeof value === 'number' && Number.isFinite(value) && value >= 1 && value <= 1000,
      )
    )
      throw new TypeError('Size must use finite millimeter dimensions between 1 and 1000');
    signal?.throwIfAborted();
    const before = requireSlideTextBox();
    const { x, y } = before.shape;
    const text = readSlideShapeText(before.shape);
    if (text === undefined || x + widthMm > before.width || y + heightMm > before.height)
      throw new Error('The entire text box must fit on the current slide');
    before.api.ImgApply!({ Width: widthMm, Height: heightMm });
    signal?.throwIfAborted();
    const after = requireSlideTextBox();
    if (
      after.api !== before.api ||
      after.slide !== before.slide ||
      after.shape !== before.shape ||
      after.width !== before.width ||
      after.height !== before.height ||
      after.shape.x !== x ||
      after.shape.y !== y ||
      after.shape.extX !== widthMm ||
      after.shape.extY !== heightMm ||
      JSON.stringify(after.selected) !== JSON.stringify(before.selected) ||
      readSlideShapeText(after.shape) !== text
    )
      throw new Error(
        'Text box size could not be verified. Inspect the document and use Undo if needed; do not retry automatically.',
      );
    return { widthMm, heightMm, page: before.page + 1, verified: true };
  },
};
