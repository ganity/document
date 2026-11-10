type FontApi = {
  fonthint_EastAsia?: number;
  fontslot_EastAsia?: number;
  fontslot_HAnsi?: number;
  GetFontSlot?: (codePoint: number, hint: number, language: number, cs: boolean, rtl: boolean) => number;
  GetFontSlotByTextPr?: (codePoint: number, textPr: { RFonts?: { Hint?: number }; Lang?: { Val?: number; EastAsia?: number }; CS?: boolean; RTL?: boolean }) => number;
  __eastAsianQuoteHint?: boolean;
};

const QUOTES = new Set([0x2018, 0x2019, 0x201c, 0x201d]);

/** PowerPoint treats smart quotes in Chinese-language runs as East Asian glyphs. */
export function installEastAsianPunctuationGuard(win: Window): boolean {
  const api = (win as Window & { AscWord?: FontApi }).AscWord;
  const resolve = api?.GetFontSlotByTextPr;
  const getFontSlot = api?.GetFontSlot;
  if (!api || typeof resolve !== 'function' || typeof getFontSlot !== 'function') return false;
  if (api.__eastAsianQuoteHint) return true;

  const original = resolve;
  const originalThis = api;
  api.GetFontSlotByTextPr = function (codePoint, textPr) {
    const language = textPr?.Lang?.Val ?? textPr?.Lang?.EastAsia ?? 0;
    if ((language & 0x3ff) === 0x04 && QUOTES.has(codePoint)) {
      return getFontSlot.call(api, codePoint, api.fonthint_EastAsia ?? 2, language,
        Boolean(textPr.CS), Boolean(textPr.RTL));
    }
    return original.call(originalThis, codePoint, textPr);
  };
  api.__eastAsianQuoteHint = true;
  return true;
}
