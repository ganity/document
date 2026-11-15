import { describe, expect, it } from 'vitest';
import { installMixedScriptLineMetricsGuard } from '../../lib/onlyoffice/guards/mixed-script-line-metrics';

function setup() {
  let font = { height: 13.55, ascent: 10.2, descent: -1.67 };
  const measurer = {
    GetHeight: () => font.height,
    GetAscender: () => font.ascent,
    GetDescender: () => font.descent,
  };
  const paragraph = { bFromDocument: false, GetText: () => '李白（701—762）' };
  class Run {
    Paragraph = paragraph;
    math = false;
    offset = 0;
    throwOnMeasure = false;
    Get_CompiledPr() {
      return {};
    }
    IsMathRun() {
      return this.math;
    }
    getYOffset() {
      return this.offset;
    }
    Recalculate_LineMetrics(state: ReturnType<typeof emptyState>) {
      const height = measurer.GetHeight();
      if (this.throwOnMeasure) throw new Error('measurement failed');
      const descent = Math.abs(measurer.GetDescender());
      const ascent = measurer.GetAscender();
      state.LineTextAscent = Math.max(state.LineTextAscent, height - descent);
      state.LineTextAscent2 = Math.max(state.LineTextAscent2, ascent);
      state.LineTextDescent = Math.max(state.LineTextDescent, descent);
      state.LineAscent = Math.max(state.LineAscent, height - descent);
      state.LineDescent = Math.max(state.LineDescent, descent);
    }
  }
  const win = { AscWord: { ParaRun: Run }, AscCommon: { g_oTextMeasurer: measurer }, Asc: { linerule_Auto: 1 } };
  const call = (run: Run, state: ReturnType<typeof emptyState>, rule = 1, line = 0) =>
    (run.Recalculate_LineMetrics as Function)(state, { Spacing: { LineRule: rule } }, line, 0);
  return {
    win: win as unknown as Window,
    Run,
    measurer,
    paragraph,
    call,
    font: (next: typeof font) => {
      font = next;
    },
  };
}

function emptyState() {
  return { LineTextAscent: 0, LineTextAscent2: 0, LineTextDescent: 0, LineAscent: 0, LineDescent: 0 };
}

describe('PowerPoint mixed-script line metrics', () => {
  it('keeps one complete font cell for same-size Chinese and Latin faces', () => {
    const s = setup();
    installMixedScriptLineMetricsGuard(s.win);
    const state = emptyState();
    s.call(new s.Run(), state);
    s.font({ height: 13.55, ascent: 10.57, descent: -2.98 });
    s.call(new s.Run(), state);
    expect(state.LineTextAscent + state.LineTextDescent).toBeCloseTo(13.55);
    expect(state.LineTextAscent).toBeCloseTo(11.88);
    expect(state.LineTextDescent).toBeCloseTo(1.67);
  });

  it('resets the measurement across lines and repeated layout passes', () => {
    const s = setup();
    installMixedScriptLineMetricsGuard(s.win);
    const state = emptyState();
    s.call(new s.Run(), state);
    Object.assign(state, emptyState());
    s.font({ height: 6, ascent: 4, descent: -2 });
    s.call(new s.Run(), state);
    expect(state.LineTextAscent + state.LineTextDescent).toBeCloseTo(6);
    Object.assign(state, emptyState());
    s.font({ height: 8, ascent: 6, descent: -2 });
    s.call(new s.Run(), state, 1, 1);
    expect(state.LineTextAscent + state.LineTextDescent).toBeCloseTo(8);
  });

  it('preserves larger type, inline objects and their required height', () => {
    const s = setup();
    installMixedScriptLineMetricsGuard(s.win);
    const state = emptyState();
    s.call(new s.Run(), state);
    s.font({ height: 20, ascent: 15, descent: -5 });
    state.LineAscent = 30;
    s.call(new s.Run(), state);
    expect(state.LineTextAscent + state.LineTextDescent).toBeCloseTo(20);
    expect(state.LineAscent).toBe(30);
  });

  it.each(['math', 'offset', 'word', 'latin', 'exact'] as const)('preserves the SDK layout for %s', (kind) => {
    const s = setup();
    installMixedScriptLineMetricsGuard(s.win);
    const run = new s.Run();
    if (kind === 'math') run.math = true;
    if (kind === 'offset') run.offset = 2;
    if (kind === 'word') s.paragraph.bFromDocument = true;
    if (kind === 'latin') s.paragraph.GetText = () => 'English only';
    const state = emptyState();
    const rule = kind === 'exact' ? 0 : 1;
    s.call(run, state, rule);
    s.font({ height: 13.55, ascent: 10.57, descent: -2.98 });
    s.call(run, state, rule);
    expect(state.LineTextAscent + state.LineTextDescent).toBeCloseTo(14.86);
  });

  it('restores the shared measurer on failure and patches only once', () => {
    const s = setup();
    const getHeight = s.measurer.GetHeight;
    expect(installMixedScriptLineMetricsGuard(s.win)).toBe(true);
    const patched = s.Run.prototype.Recalculate_LineMetrics;
    expect(installMixedScriptLineMetricsGuard(s.win)).toBe(true);
    expect(s.Run.prototype.Recalculate_LineMetrics).toBe(patched);
    const run = new s.Run();
    run.throwOnMeasure = true;
    expect(() => s.call(run, emptyState())).toThrow('measurement failed');
    expect(s.measurer.GetHeight).toBe(getHeight);
    expect(installMixedScriptLineMetricsGuard({} as Window)).toBe(false);
  });
});
