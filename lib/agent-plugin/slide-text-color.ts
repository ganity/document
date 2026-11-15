import type { AgentTool } from '@ranuts/agent-core/types';
import { requireEditorContext, type EditorAsc } from './editor-bridge';
import { readSlideShapeText } from './slide-text-read';
import { requireSlideTextSelection, type SlideTextApi } from './slide-text-selection';

interface NativeColor {
  put_type(type: number): void;
  put_r(value: number): void;
  put_g(value: number): void;
  put_b(value: number): void;
}
type ColorAsc = EditorAsc & {
  asc_CColor?: new () => NativeColor;
  c_oAscColor?: { COLOR_TYPE_SRGB: number };
  c_oAscFill?: { FILL_TYPE_SOLID: number };
};
type ColorApi = SlideTextApi & {
  put_TextColor?(color: NativeColor): void;
  WordControl?: {
    m_oLogicDocument?: {
      GetCalculatedTextPr?(): {
        Unifill?: {
          transparent?: number | null;
          fill?: {
            type?: number;
            color?: {
              Mods?: { Mods: unknown[] } | null;
              color?: { type?: number };
              RGBA?: { R: number; G: number; B: number; A: number };
            };
          };
        };
      };
    };
  };
};
function selection() {
  const context = requireEditorContext();
  const api = context.api as ColorApi;
  const selected = requireSlideTextSelection(api);
  const asc = context.Asc as ColorAsc;
  if (
    typeof api.put_TextColor !== 'function' ||
    typeof api.WordControl?.m_oLogicDocument?.GetCalculatedTextPr !== 'function' ||
    typeof asc.asc_CColor !== 'function' ||
    asc.c_oAscColor?.COLOR_TYPE_SRGB !== 1 ||
    asc.c_oAscFill?.FILL_TYPE_SOLID !== 3
  )
    throw new Error('Text color requires supported native setter, color constructor and readback');
  const color = new asc.asc_CColor();
  if (['put_type', 'put_r', 'put_g', 'put_b'].some((key) => typeof Reflect.get(color, key) !== 'function'))
    throw new Error('Native RGB color construction is unavailable');
  return { ...selected, color };
}
function selectedColor(api: ColorApi): string | undefined {
  // get_Color() reads the legacy Color field; the PPT native command writes Unifill.
  const fill = api.WordControl?.m_oLogicDocument?.GetCalculatedTextPr?.().Unifill;
  const color = fill?.fill?.color;
  const rgba = color?.RGBA;
  if (
    fill?.transparent != null ||
    fill?.fill?.type !== 3 ||
    color?.color?.type !== 1 ||
    (color.Mods?.Mods.length ?? 0) !== 0 ||
    !rgba ||
    rgba.A !== 255
  )
    return undefined;
  const channels = [rgba.R, rgba.G, rgba.B];
  if (channels.some((value) => !Number.isInteger(value) || value < 0 || value > 255)) return undefined;
  return channels
    .map((value) => value.toString(16).padStart(2, '0'))
    .join('')
    .toUpperCase();
}
export function canSetSlideTextColor(): boolean {
  try {
    selection();
    return true;
  } catch {
    return false;
  }
}
export const setSlideTextColorTool: AgentTool<
  { colorHex: string },
  { colorHex: string; page: number; verified: true }
> = {
  name: 'set_slide_text_color',
  description:
    'PPT only. Set opaque RGB color on the nonempty plain-text selection in one current-slide text object. Does not select text or save.',
  inputSchema: {
    type: 'object',
    properties: { colorHex: { type: 'string', pattern: '^[0-9A-F]{6}$' } },
    required: ['colorHex'],
    additionalProperties: false,
  },
  readOnlyHint: false,
  execute: async ({ colorHex }, signal) => {
    if (typeof colorHex !== 'string' || !/^[0-9A-F]{6}$/.test(colorHex))
      throw new TypeError('Text color must be six uppercase RGB hex digits without a prefix');
    signal?.throwIfAborted();
    const before = selection();
    const text = readSlideShapeText(before.shape);
    if (text === undefined) throw new Error('Selected text content cannot be verified');
    before.color.put_type(1);
    before.color.put_r(parseInt(colorHex.slice(0, 2), 16));
    before.color.put_g(parseInt(colorHex.slice(2, 4), 16));
    before.color.put_b(parseInt(colorHex.slice(4, 6), 16));
    before.api.put_TextColor!(before.color);
    signal?.throwIfAborted();
    const after = selection();
    if (
      after.api !== before.api ||
      after.slide !== before.slide ||
      after.shape !== before.shape ||
      JSON.stringify(after.selected) !== JSON.stringify(before.selected) ||
      readSlideShapeText(after.shape) !== text ||
      selectedColor(after.api) !== colorHex
    )
      throw new Error(
        'Text color could not be verified. Inspect the document and use Undo if needed; do not retry automatically.',
      );
    return { colorHex, page: before.page + 1, verified: true };
  },
};
