type Matrix = {
  sx: number; shx: number; shy: number; sy: number; tx: number; ty: number;
  CreateDublicate(): Matrix;
};
type Graphics = {
  m_oTransform?: Matrix;
  m_bIntegerGrid?: boolean;
  transform3?: (matrix: Matrix) => void;
  SetIntegerGrid?: (value: boolean) => void;
};
type Shape = { txBody?: { getCompiledBodyPr?: () => { vert?: number } } };
type DrawText = (this: Shape, graphics: Graphics, ...args: unknown[]) => unknown;
type DrawGrapheme = (id: number, graphics: Graphics, x: number, y: number, size: number, scale?: number) => unknown;
type VerticalWindow = Window & {
  AscFormat?: { CShape?: { prototype: { drawTxBody?: DrawText } }; nVertTTeaVert?: number };
  AscFonts?: {
    DrawGrapheme?: DrawGrapheme;
    GetGraphemeCodePoints?: (id: number) => number[];
    GetGraphemeFontId?: (id: number) => number;
    GetFontNameByFontId?: (id: number) => string;
    GetFontStyleByFontId?: (id: number) => number;
    GetGraphemeWidth?: (id: number) => number;
  };
  AscCommon?: { g_oTextMeasurer?: {
    SetFontInternal(name: string, size: number, style: number): void;
    GetAscender(): number; GetDescender(): number;
  } };
  __eastAsianVerticalGuard?: boolean;
};

/** Counter-rotate a glyph cell within the SDK's clockwise vertical text transform. */
export function uprightVerticalMatrix(matrix: Matrix, x: number, y: number): Matrix {
  const result = matrix.CreateDublicate();
  result.sx = -matrix.shx; result.shy = -matrix.sy;
  result.shx = matrix.sx; result.sy = matrix.shy;
  result.tx = matrix.tx + matrix.sx * (x - y) + matrix.shx * (y + x);
  result.ty = matrix.ty + matrix.shy * (x - y) + matrix.sy * (y + x);
  return result;
}

/** Keep Han glyphs upright for source eaVert, without changing text or hit-test cells. */
export function installEastAsianVerticalGuard(win: Window): boolean {
  const scope = win as VerticalWindow;
  const format = scope.AscFormat;
  const fonts = scope.AscFonts;
  const prototype = format?.CShape?.prototype;
  const drawText = prototype?.drawTxBody;
  const drawGlyph = fonts?.DrawGrapheme;
  const measurer = scope.AscCommon?.g_oTextMeasurer;
  if (!prototype || !fonts || typeof drawText !== 'function' || typeof drawGlyph !== 'function' ||
      typeof fonts.GetGraphemeCodePoints !== 'function' || typeof fonts.GetGraphemeFontId !== 'function' ||
      typeof fonts.GetFontNameByFontId !== 'function' || typeof fonts.GetFontStyleByFontId !== 'function' ||
      typeof fonts.GetGraphemeWidth !== 'function' || !measurer || format?.nVertTTeaVert === undefined) return false;
  if (scope.__eastAsianVerticalGuard) return true;
  const active = new WeakSet<Graphics>();

  prototype.drawTxBody = function (graphics, ...args) {
    const enabled = this.txBody?.getCompiledBodyPr?.()?.vert === format.nVertTTeaVert;
    const alreadyActive = active.has(graphics);
    if (enabled) active.add(graphics);
    try { return drawText.call(this, graphics, ...args); }
    finally { if (enabled && !alreadyActive) active.delete(graphics); }
  };

  fonts.DrawGrapheme = function (id, graphics, x, y, size, scale = 1) {
    const points = fonts.GetGraphemeCodePoints!(id);
    const matrix = graphics.m_oTransform;
    // Punctuation needs vertical OpenType forms; Latin remains sideways. Do not
    // rotate ligatures, combining clusters, or non-canvas export backends here.
    if (!active.has(graphics) || points.length !== 1 ||
        !/[\u3400-\u9fff\uf900-\ufaff\u{20000}-\u{323af}]/u.test(String.fromCodePoint(points[0])) ||
        !matrix?.CreateDublicate || typeof graphics.transform3 !== 'function' ||
        !Number.isFinite(size) || size <= 0 || !Number.isFinite(scale) || scale <= 0) {
      return drawGlyph(id, graphics, x, y, size, scale);
    }
    const fontId = fonts.GetGraphemeFontId!(id);
    measurer.SetFontInternal(fonts.GetFontNameByFontId!(fontId), size, fonts.GetFontStyleByFontId!(fontId));
    const centerX = x + fonts.GetGraphemeWidth!(id) * size * scale / 2;
    const centerY = y - (measurer.GetAscender() + measurer.GetDescender()) * scale / 2;
    const original = matrix.CreateDublicate();
    const grid = graphics.m_bIntegerGrid;
    try {
      graphics.transform3(uprightVerticalMatrix(original, centerX, centerY));
      return drawGlyph(id, graphics, x, y, size, scale);
    } finally {
      graphics.transform3(original);
      if (typeof grid === 'boolean') graphics.SetIntegerGrid?.(grid);
    }
  };
  scope.__eastAsianVerticalGuard = true;
  return true;
}
