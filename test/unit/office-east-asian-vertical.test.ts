import { describe, expect, it, vi } from 'vitest';
import { installEastAsianVerticalGuard, uprightVerticalMatrix } from '../../lib/onlyoffice/guards/east-asian-vertical';

function matrix(sx = 0, shx = -1, shy = 1, sy = 0, tx = 30, ty = 40) {
  return {
    sx,
    shx,
    shy,
    sy,
    tx,
    ty,
    CreateDublicate() {
      return matrix(this.sx, this.shx, this.shy, this.sy, this.tx, this.ty);
    },
  };
}

function setup() {
  const glyph = vi.fn();
  const fonts = {
    DrawGrapheme: glyph,
    GetGraphemeCodePoints: (id: number) => (id === -1 ? [0x4e00, 0x301] : [id]),
    GetGraphemeFontId: () => 1,
    GetFontNameByFontId: () => 'KaiTi',
    GetFontStyleByFontId: () => 0,
    GetGraphemeWidth: () => 25.4 / 72,
  };
  const measurer = { SetFontInternal: vi.fn(), GetAscender: () => 8, GetDescender: () => -2 };
  const graphics = {
    m_oTransform: matrix(),
    m_bIntegerGrid: true,
    transform3(m: ReturnType<typeof matrix>) {
      this.m_oTransform = m;
      this.m_bIntegerGrid = false;
    },
    SetIntegerGrid(grid: boolean) {
      this.m_bIntegerGrid = grid;
    },
  };
  class Shape {
    vert = 0;
    code = 0x674e;
    txBody = { getCompiledBodyPr: () => ({ vert: this.vert }) };
    drawTxBody(g: typeof graphics) {
      fonts.DrawGrapheme(this.code, g, 2, 9, 32, 1);
    }
  }
  const win = {
    AscFormat: { CShape: Shape, nVertTTeaVert: 0 },
    AscFonts: fonts,
    AscCommon: { g_oTextMeasurer: measurer },
  } as unknown as Window;
  return { win, Shape, graphics, fonts, glyph, measurer };
}

describe('upright Han glyphs in source eaVert text', () => {
  it('counter-rotates around the cell center while preserving the existing affine transform', () => {
    const before = matrix();
    const after = uprightVerticalMatrix(before, 2, 3);
    const point = (m: typeof before) => [m.sx * 2 + m.shx * 3 + m.tx, m.shy * 2 + m.sy * 3 + m.ty];
    expect(point(after)).toEqual(point(before));
    expect([after.sx, after.shx, after.shy, after.sy]).toEqual([1, 0, -0, 1]);
    expect(before).toMatchObject({ sx: 0, shx: -1, shy: 1, sy: 0, tx: 30, ty: 40 });
  });

  it('renders the original Han glyph and restores the transform and integer grid', () => {
    const s = setup();
    installEastAsianVerticalGuard(s.win);
    const shape = new s.Shape(),
      transform = s.graphics.m_oTransform.CreateDublicate();
    let painted: ReturnType<typeof matrix> | undefined;
    s.glyph.mockImplementation((_id, g) => {
      painted = g.m_oTransform.CreateDublicate();
    });
    shape.drawTxBody(s.graphics);
    expect(painted?.sx).toBe(1);
    expect(s.glyph).toHaveBeenCalledWith(shape.code, s.graphics, 2, 9, 32, 1);
    expect(s.graphics.m_oTransform).toMatchObject({
      sx: transform.sx,
      shx: transform.shx,
      tx: transform.tx,
      ty: transform.ty,
    });
    expect(s.graphics.m_bIntegerGrid).toBe(true);
    expect(shape.txBody.getCompiledBodyPr().vert).toBe(0);
  });

  it.each([0x41, 0x3002, 0x2014, -1])('preserves Latin, punctuation and combining clusters (%s)', (code) => {
    const s = setup();
    installEastAsianVerticalGuard(s.win);
    const shape = new s.Shape();
    shape.code = code;
    shape.drawTxBody(s.graphics);
    expect(s.measurer.SetFontInternal).not.toHaveBeenCalled();
    expect(s.graphics.m_oTransform.sx).toBe(0);
  });

  it.each([1, 3, 4, 5])('preserves horizontal and other rotation modes (%s)', (vert) => {
    const s = setup();
    installEastAsianVerticalGuard(s.win);
    const shape = new s.Shape();
    shape.vert = vert;
    shape.drawTxBody(s.graphics);
    expect(s.measurer.SetFontInternal).not.toHaveBeenCalled();
  });

  it('cleans up after drawing throws and installs only once', () => {
    const s = setup();
    expect(installEastAsianVerticalGuard(s.win)).toBe(true);
    const patched = s.fonts.DrawGrapheme;
    expect(installEastAsianVerticalGuard(s.win)).toBe(true);
    expect(s.fonts.DrawGrapheme).toBe(patched);
    s.glyph.mockImplementationOnce(() => {
      throw new Error('draw failed');
    });
    expect(() => new s.Shape().drawTxBody(s.graphics)).toThrow('draw failed');
    expect(s.graphics.m_oTransform.sx).toBe(0);
    expect(s.graphics.m_bIntegerGrid).toBe(true);
    s.measurer.SetFontInternal.mockClear();
    (s.fonts.DrawGrapheme as Function)(0x674e, s.graphics, 2, 9, 32);
    expect(s.measurer.SetFontInternal).not.toHaveBeenCalled();
    expect(installEastAsianVerticalGuard({} as Window)).toBe(false);
  });
});
