/**
 * 22. Preserve source line metrics for mixed Latin and East Asian text.
 */
type LineState = {
  LineTextAscent: number;
  LineTextAscent2: number;
  LineTextDescent: number;
  LineAscent: number;
  LineDescent: number;
};

type Run = {
  Paragraph?: { bFromDocument?: boolean; GetText?: () => string };
  Get_CompiledPr?: (copy: boolean) => { CS?: boolean; RTL?: boolean };
  Recalculate_LineMetrics?: (...args: unknown[]) => void;
  IsMathRun?: () => boolean;
  getYOffset?: () => number;
};

type Api = {
  ParaRun?: { prototype: Run };
  __mixedScriptLineMetricsGuard?: boolean;
};

type TextMeasurer = {
  GetHeight?: () => number;
  GetAscender?: () => number;
  GetDescender?: () => number;
};

type OfficeWindow = Window & {
  AscWord?: Api;
  AscCommon?: { g_oTextMeasurer?: TextMeasurer };
  Asc?: { linerule_Auto?: number };
};

type LineMetrics = {
  paragraph: Run['Paragraph'];
  line: unknown;
  range: unknown;
  maxHeight: number;
  ascent: number;
  descent: number;
  baseline: number;
  hasChineseText: boolean;
  unsupported: boolean;
  raw: LineState | null;
};

const STATE_KEYS = ['LineTextAscent', 'LineTextAscent2', 'LineTextDescent', 'LineAscent', 'LineDescent'] as const;

/** Match PowerPoint's one-line font box when its Latin and East Asian faces are mixed. */
export function installMixedScriptLineMetricsGuard(win: Window): boolean {
  const scope = win as OfficeWindow;
  const api = scope.AscWord;
  const runPrototype = api?.ParaRun?.prototype;
  const original = runPrototype?.Recalculate_LineMetrics;
  const measurer = scope.AscCommon?.g_oTextMeasurer;
  if (
    !api ||
    !runPrototype ||
    typeof original !== 'function' ||
    !measurer ||
    typeof measurer.GetHeight !== 'function' ||
    typeof measurer.GetAscender !== 'function' ||
    typeof measurer.GetDescender !== 'function' ||
    scope.Asc?.linerule_Auto === undefined
  ) {
    return false;
  }
  if (api.__mixedScriptLineMetricsGuard) return true;

  const byState = new WeakMap<object, LineMetrics>();
  runPrototype.Recalculate_LineMetrics = function (this: Run, ...args: unknown[]) {
    const [state, paragraphProperties, line, range] = args as [
      LineState,
      { Spacing?: { LineRule?: number } },
      unknown,
      unknown,
    ];
    if (
      this.Paragraph?.bFromDocument !== false ||
      paragraphProperties?.Spacing?.LineRule !== scope.Asc?.linerule_Auto
    ) {
      return original.apply(this, args);
    }

    let metrics = byState.get(state);
    if (
      !metrics ||
      metrics.paragraph !== this.Paragraph ||
      metrics.line !== line ||
      metrics.range !== range ||
      state.LineTextAscent === 0
    ) {
      metrics = {
        paragraph: this.Paragraph,
        line,
        range,
        maxHeight: 0,
        ascent: 0,
        descent: 0,
        baseline: 0,
        hasChineseText: /[\u3400-\u9fff]/u.test(this.Paragraph?.GetText?.() ?? ''),
        unsupported: false,
        raw: null,
      };
      byState.set(state, metrics);
    }

    const textProperties = this.Get_CompiledPr?.(false);
    if (this.IsMathRun?.() || this.getYOffset?.() !== 0 || textProperties?.CS || textProperties?.RTL) {
      metrics.unsupported = true;
    }

    if (metrics.raw) {
      for (const key of STATE_KEYS) state[key] = Math.max(state[key], metrics.raw[key]);
    }
    const getHeight = measurer.GetHeight!;
    measurer.GetHeight = function () {
      const height = getHeight.call(measurer);
      const descent = Math.abs(measurer.GetDescender!());
      const ascent = Math.min(measurer.GetAscender!(), height - descent);
      const baseline = height - descent;
      // Choose one complete font cell. Combining the largest ascender, descender
      // and leading from different fonts invents a cell larger than either font.
      if (
        height > metrics!.maxHeight + 0.001 ||
        (Math.abs(height - metrics!.maxHeight) < 0.001 && baseline > metrics!.baseline)
      ) {
        metrics!.maxHeight = height;
        metrics!.ascent = ascent;
        metrics!.descent = descent;
        metrics!.baseline = baseline;
      }
      return height;
    };

    try {
      original.apply(this, args);
    } finally {
      measurer.GetHeight = getHeight;
    }

    metrics.raw = Object.fromEntries(STATE_KEYS.map((key) => [key, state[key]])) as LineState;
    if (
      metrics.hasChineseText &&
      !metrics.unsupported &&
      metrics.maxHeight > 0 &&
      state.LineTextAscent + state.LineTextDescent > metrics.maxHeight
    ) {
      if (state.LineAscent <= state.LineTextAscent + 0.001) state.LineAscent = metrics.baseline;
      if (state.LineDescent <= state.LineTextDescent + 0.001) state.LineDescent = metrics.descent;
      state.LineTextAscent = metrics.baseline;
      state.LineTextAscent2 = metrics.ascent;
      state.LineTextDescent = metrics.descent;
    }
  };

  api.__mixedScriptLineMetricsGuard = true;
  return true;
}
