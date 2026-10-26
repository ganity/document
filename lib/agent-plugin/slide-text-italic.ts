import type { AgentTool } from '@ranuts/agent-core/types';
import { requireEditorApi } from './editor-bridge';
import { readSlideShapeText } from './slide-text-read';
import { requireSlideTextSelection, type SlideTextApi } from './slide-text-selection';

interface ItalicApi extends SlideTextApi {
  put_TextPrItalic?(enabled: boolean): void;
  get_TextProps?(): { get_TextPr?(): { get_Italic?(): boolean } };
}
function selection() {
  const api = requireEditorApi() as ItalicApi;
  const selected = requireSlideTextSelection(api);
  if (
    typeof api.put_TextPrItalic !== 'function' ||
    typeof api.get_TextProps?.().get_TextPr?.().get_Italic !== 'function'
  )
    throw new Error('Italic formatting requires native setter and readback');
  return selected;
}

export function canSetSlideTextItalic(): boolean {
  try {
    selection();
    return true;
  } catch {
    return false;
  }
}

export const setSlideTextItalicTool: AgentTool<
  { enabled: boolean },
  { enabled: boolean; page: number; verified: true }
> = {
  name: 'set_slide_text_italic',
  description:
    'PPT only. Set or clear italic on the current nonempty plain-text selection in one text object. Verifies native italic/text; does not select text or save.',
  inputSchema: {
    type: 'object',
    properties: { enabled: { type: 'boolean' } },
    required: ['enabled'],
    additionalProperties: false,
  },
  readOnlyHint: false,
  execute: async ({ enabled }, signal) => {
    if (typeof enabled !== 'boolean') throw new TypeError('Italic formatting requires a boolean');
    signal?.throwIfAborted();
    const before = selection();
    const text = readSlideShapeText(before.shape);
    if (text === undefined) throw new Error('Selected text content cannot be verified');
    before.api.put_TextPrItalic!(enabled);
    signal?.throwIfAborted();
    const after = selection();
    if (
      after.api !== before.api ||
      after.slide !== before.slide ||
      after.shape !== before.shape ||
      JSON.stringify(after.selected) !== JSON.stringify(before.selected) ||
      readSlideShapeText(after.shape) !== text ||
      after.api.get_TextProps!().get_TextPr!().get_Italic!() !== enabled
    )
      throw new Error(
        'Italic formatting could not be verified. Inspect the document and use Undo if needed; do not retry automatically.',
      );
    return { enabled, page: before.page + 1, verified: true };
  },
};
