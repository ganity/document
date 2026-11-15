import type { AgentTool } from '@ranuts/agent-core/types';
import { requireEditorContext } from './editor-bridge';
import { readSlideShapeText } from './slide-text-read';
import { requireSlideTextSelection, type SlideTextApi } from './slide-text-selection';

interface FontApi extends SlideTextApi {
  put_TextPrFontName?(fontFamily: string): void;
  get_TextProps?(): { get_TextPr?(): { get_FontFamily?(): { get_Name?(): string } } };
}
interface FontLoader {
  currentInfoLoaded: unknown;
  fontInfos: Array<{ Name: string; indexR: number; indexI: number; indexB: number; indexBI: number }>;
  fontFiles: Array<{ CheckLoaded?(): boolean }>;
}
function selection() {
  const context = requireEditorContext();
  const api = context.api as FontApi;
  const selected = requireSlideTextSelection(api);
  const loader = (context.AscCommon as { g_font_loader?: FontLoader } | undefined)?.g_font_loader;
  if (
    typeof api.put_TextPrFontName !== 'function' ||
    typeof api.get_TextProps?.().get_TextPr?.().get_FontFamily?.()?.get_Name !== 'function' ||
    !loader ||
    loader.currentInfoLoaded ||
    !Array.isArray(loader.fontInfos) ||
    !Array.isArray(loader.fontFiles)
  )
    throw new Error('Font family requires native setter/readback and an idle font loader');
  // Unloaded fonts defer the native edit and can later affect a different selection.
  const fonts = [
    ...new Set(
      loader.fontInfos
        .filter(
          (font) =>
            typeof font.Name === 'string' &&
            font.Name.length > 0 &&
            font.Name.length <= 128 &&
            font.Name.trim() === font.Name &&
            // eslint-disable-next-line no-control-regex -- Native font names must reject ASCII control characters.
            !/[\x00-\x1f\x7f]/.test(font.Name) &&
            [font.indexR, font.indexI, font.indexB, font.indexBI].every(
              (index) => Number.isSafeInteger(index) && index >= 0 && loader.fontFiles[index]?.CheckLoaded?.() === true,
            ),
        )
        .map((font) => font.Name),
    ),
  ].slice(0, 100);
  return { ...selected, api, fonts };
}

export function getSlideFontFamilies(): string[] {
  try {
    return selection().fonts;
  } catch {
    return [];
  }
}

export const setSlideFontFamilyTool: AgentTool<
  { fontFamily: string },
  { fontFamily: string; page: number; verified: true }
> = {
  name: 'set_slide_font_family',
  description:
    'PPT only. Set an already loaded font family from the current capability enum on one nonempty plain-text selection. Verifies native family/text; does not load fonts, select text or save.',
  inputSchema: {
    type: 'object',
    properties: { fontFamily: { type: 'string', minLength: 1, maxLength: 128 } },
    required: ['fontFamily'],
    additionalProperties: false,
  },
  readOnlyHint: false,
  execute: async ({ fontFamily }, signal) => {
    signal?.throwIfAborted();
    const before = selection();
    if (!before.fonts.includes(fontFamily)) throw new Error('Font family is not in the latest loaded-font capability');
    const text = readSlideShapeText(before.shape);
    if (text === undefined) throw new Error('Selected text content cannot be verified');
    before.api.put_TextPrFontName!(fontFamily);
    signal?.throwIfAborted();
    const after = selection();
    if (
      after.api !== before.api ||
      after.slide !== before.slide ||
      after.shape !== before.shape ||
      JSON.stringify(after.selected) !== JSON.stringify(before.selected) ||
      readSlideShapeText(after.shape) !== text ||
      after.api.get_TextProps!().get_TextPr!().get_FontFamily!().get_Name!() !== fontFamily
    )
      throw new Error(
        'Font family could not be verified. Inspect the document and use Undo if needed; do not retry automatically.',
      );
    return { fontFamily, page: before.page + 1, verified: true };
  },
};
