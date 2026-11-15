import { describe, expect, it, vi } from 'vitest';
import { applyImageColorKey, installImageColorKeyGuard } from '../../lib/onlyoffice/guards/image-color-key';

function setup() {
  class RGB {
    RGBA = { R: 255, G: 255, B: 255 };
  }
  class Effect {
    clrFrom = { color: new RGB(), Mods: null };
    clrTo = { color: new RGB(), Mods: { Mods: [{ name: 'alpha', val: 0 }] } };
    useA = false;
  }
  const image = { complete: true, naturalWidth: 2, naturalHeight: 1 };
  const pixels = new Uint8ClampedArray([255, 255, 255, 255, 0, 128, 0, 255]);
  const context = { drawImage: vi.fn(), getImageData: () => ({ data: pixels }), putImageData: vi.fn() };
  const canvas = { getContext: () => context, width: 0, height: 0 };
  const createElement = vi.fn(() => canvas);
  const paint = vi.fn();
  class Graphics {
    drawImage2 = paint;
  }
  class Shape {
    blipFill = { RasterImageId: 'source.jpg', Effects: [new Effect()] as unknown[] };
    fail = false;
    draw(graphics: Graphics) {
      graphics.drawImage2(image, 1, 2, 3, 4, 255, { l: 0, t: 0, r: 100, b: 96 });
      if (this.fail) throw new Error('draw failed');
    }
  }
  const win = {
    document: { createElement },
    AscFormat: { CImageShape: Shape, CClrChange: Effect, CRGBColor: RGB },
    AscCommon: { getFullImageSrc2: (id: string) => id },
    Asc: { editor: { ImageLoader: { map_image_index: { 'source.jpg': { Image: image } } } } },
  };
  return { win: win as unknown as Window, Shape, Graphics, image, canvas, context, createElement, paint, pixels };
}

describe('OOXML transparent color-key rendering', () => {
  it('changes only alpha of exactly matching pixels, including partial alpha', () => {
    const data = new Uint8ClampedArray([255, 255, 255, 100, 254, 255, 255, 255, 0, 1, 2, 0]);
    applyImageColorKey(data, [255, 255, 255]);
    expect([...data]).toEqual([255, 255, 255, 0, 254, 255, 255, 255, 0, 1, 2, 0]);
  });

  it('uses a cached rendering copy and preserves crop/position and source effects', () => {
    const s = setup();
    installImageColorKeyGuard(s.win);
    const shape = new s.Shape(),
      graphics = new s.Graphics();
    const effects = structuredClone(shape.blipFill);
    shape.draw(graphics);
    shape.draw(graphics);
    expect(s.paint).toHaveBeenLastCalledWith(s.canvas, 1, 2, 3, 4, 255, { l: 0, t: 0, r: 100, b: 96 });
    expect(s.createElement).toHaveBeenCalledTimes(1);
    expect(s.pixels[3]).toBe(0);
    expect(s.pixels[7]).toBe(255);
    expect(shape.blipFill).toMatchObject(effects);
    expect(graphics.drawImage2).toBe(s.paint);
  });

  it.each(['no-effect', 'other-effect', 'multiple', 'opaque', 'theme', 'loading', 'too-large', 'unreadable'])(
    'keeps the original rendering for %s',
    (kind) => {
      const s = setup();
      installImageColorKeyGuard(s.win);
      const shape = new s.Shape();
      const effect = shape.blipFill.Effects[0] as any;
      if (kind === 'no-effect') shape.blipFill.Effects = [];
      if (kind === 'other-effect') shape.blipFill.Effects = [{}];
      if (kind === 'multiple') shape.blipFill.Effects.push({});
      if (kind === 'opaque') effect.clrTo.Mods.Mods[0].val = 100000;
      if (kind === 'theme') effect.clrFrom.color = { RGBA: { R: 255, G: 255, B: 255 } };
      if (kind === 'loading') s.image.complete = false;
      if (kind === 'too-large') s.image.naturalWidth = 20_000_000;
      if (kind === 'unreadable')
        s.context.getImageData = () => {
          throw new Error('tainted');
        };
      shape.draw(new s.Graphics());
      expect(s.paint.mock.calls[0][0]).toBe(s.image);
    },
  );

  it('keeps another use of the same source opaque and restores drawing after errors', () => {
    const s = setup();
    installImageColorKeyGuard(s.win);
    const shape = new s.Shape(),
      graphics = new s.Graphics();
    shape.fail = true;
    expect(() => shape.draw(graphics)).toThrow('draw failed');
    expect(graphics.drawImage2).toBe(s.paint);
    const plain = new s.Shape();
    plain.blipFill.Effects = [];
    plain.draw(graphics);
    expect(s.paint.mock.calls.at(-1)![0]).toBe(s.image);
    const patched = s.Shape.prototype.draw;
    expect(installImageColorKeyGuard(s.win)).toBe(true);
    expect(s.Shape.prototype.draw).toBe(patched);
    expect(installImageColorKeyGuard({} as Window)).toBe(false);
  });
});
