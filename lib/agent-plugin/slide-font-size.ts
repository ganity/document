import type { AgentTool } from '@ranuts/agent-core/types';
import { requireEditorApi } from './editor-bridge';
import { readSlideShapeText } from './slide-text-read';
import { requireSlideTextSelection, type SlideTextApi } from './slide-text-selection';

interface FontApi extends SlideTextApi {
  put_TextPrFontSize?(size: number): void;
  get_TextProps?(): { get_TextPr?(): { get_FontSize?(): number } };
}
function selection() {
  const api = requireEditorApi() as FontApi;
  const selected = requireSlideTextSelection(api);
  if (
    typeof api.put_TextPrFontSize !== 'function' ||
    typeof api.get_TextProps?.().get_TextPr?.().get_FontSize !== 'function'
  )
    throw new Error('Font size requires native setter and readback');
  return selected;
}

export function canSetSlideFontSize(): boolean {
  try {
    selection();
    return true;
  } catch {
    return false;
  }
}

export const setSlideFontSizeTool: AgentTool<
  { fontSizePt: number },
  { fontSizePt: number; page: number; verified: true }
> = {
  name: 'set_slide_font_size',
  description:
    'PPT only. Set 1-96 pt on the current nonempty plain-text selection in one text object. Verifies native size/text; does not select text or save.',
  inputSchema: {
    type: 'object',
    properties: { fontSizePt: { type: 'number', minimum: 1, maximum: 96, multipleOf: 0.5 } },
    required: ['fontSizePt'],
    additionalProperties: false,
  },
  readOnlyHint: false,
  execute: async ({ fontSizePt }, signal) => {
    if (!Number.isFinite(fontSizePt) || fontSizePt < 1 || fontSizePt > 96 || !Number.isInteger(fontSizePt * 2))
      throw new TypeError('Font size must be 1-96 pt in half-point increments');
    signal?.throwIfAborted();
    const before = selection();
    const text = readSlideShapeText(before.shape);
    if (text === undefined) throw new Error('Selected text content cannot be verified');
    before.api.put_TextPrFontSize!(fontSizePt);
    signal?.throwIfAborted();
    const after = selection();
    if (
      after.api !== before.api ||
      after.slide !== before.slide ||
      after.shape !== before.shape ||
      JSON.stringify(after.selected) !== JSON.stringify(before.selected) ||
      readSlideShapeText(after.shape) !== text ||
      after.api.get_TextProps!().get_TextPr!().get_FontSize!() !== fontSizePt
    )
      throw new Error(
        'Font size could not be verified. Inspect the document and use Undo if needed; do not retry automatically.',
      );
    return { fontSizePt, page: before.page + 1, verified: true };
  },
};
