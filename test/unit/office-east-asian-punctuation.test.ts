import { describe, expect, it, vi } from 'vitest';
import { installEastAsianPunctuationGuard } from '../../lib/onlyoffice/guards/east-asian-punctuation';

function setup() {
  const calls: unknown[][] = [];
  const original = vi.fn((codePoint, hint, language, cs, rtl) => {
    calls.push([codePoint, hint, language, cs, rtl]);
    return hint === 2 ? 2 : 8;
  });
  const win = { AscWord: {
    fonthint_EastAsia: 2,
    GetFontSlot: original,
    GetFontSlotByTextPr: (codePoint: number, textPr: Record<string, any>) =>
      original(codePoint, textPr.RFonts.Hint, textPr.Lang.EastAsia, textPr.CS, textPr.RTL),
  } } as unknown as Window;
  return { win, calls, original };
}

describe('PowerPoint East Asian punctuation font slots', () => {
  it('uses Chinese East Asian metrics for smart quotes in a Chinese-language run', () => {
    const { win, calls } = setup();
    expect(installEastAsianPunctuationGuard(win)).toBe(true);
    const api = (win as any).AscWord;
    expect(api.GetFontSlotByTextPr(0x201c, { RFonts: { Hint: 0 }, Lang: { Val: 2052, EastAsia: 1033 } })).toBe(2);
    expect(calls[0]).toEqual([0x201c, 2, 2052, false, false]);
  });

  it('preserves English quote and ordinary punctuation metrics', () => {
    const { win, calls } = setup();
    installEastAsianPunctuationGuard(win);
    const api = (win as any).AscWord;
    const textPr = { RFonts: { Hint: 0 }, Lang: { Val: 1033, EastAsia: 2052 } };
    api.GetFontSlotByTextPr(0x201c, textPr);
    api.GetFontSlotByTextPr(0x4e2d, { ...textPr, Lang: { Val: 2052, EastAsia: 2052 } });
    expect(calls.map((call) => call[1])).toEqual([0, 0]);
  });

  it('patches once when the editor iframe is prepared repeatedly', () => {
    const { win } = setup();
    expect(installEastAsianPunctuationGuard(win)).toBe(true);
    const wrapped = (win as any).AscWord.GetFontSlotByTextPr;
    expect(installEastAsianPunctuationGuard(win)).toBe(true);
    expect((win as any).AscWord.GetFontSlotByTextPr).toBe(wrapped);
  });
});
