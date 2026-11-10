import { expect, it, vi } from 'vitest';
import { featherEllipseAlpha, installImageSoftEdgeGuard } from '../../lib/onlyoffice/guards/image-soft-edge';

it('feathers the ellipse inward, preserving RGB, center alpha and symmetry', () => {
  const pixels = new Uint8ClampedArray(21 * 21 * 4);
  for (let i = 0; i < pixels.length; i += 4) pixels.set([40, 70, 90, 128], i);
  featherEllipseAlpha(pixels, 21, 21, 21, 21, 3);
  const alpha = (x: number, y: number) => pixels[(y * 21 + x) * 4 + 3];
  expect(alpha(10, 10)).toBe(128);
  expect(alpha(0, 0)).toBe(0);
  expect(alpha(0, 10)).toBeGreaterThan(0);
  expect(alpha(0, 10)).toBeLessThan(alpha(1, 10));
  expect(alpha(0, 10)).toBe(alpha(20, 10));
  expect(alpha(10, 0)).toBe(alpha(10, 20));
  for (let i = 0; i < pixels.length; i += 4) expect([...pixels.slice(i, i + 3)]).toEqual([40, 70, 90]);
});

function setup() {
  const source = { complete: true, naturalWidth: 3, naturalHeight: 3 };
  const pixels = new Uint8ClampedArray(36).fill(255);
  const canvas = { width: 0, height: 0, getContext: () => ({ drawImage() {},
    getImageData: () => ({ data: pixels }), putImageData() {} }) };
  const createElement = vi.fn(() => canvas), paint = vi.fn();
  class Shape {
    extX = 100; extY = 80;
    blipFill: any = { RasterImageId: 'photo', srcRect: null, Effects: [] };
    spPr: any = { geometry: { preset: 'ellipse' }, effectProps: { EffectLst: { softEdge: { rad: 112500 } } } };
    fail = false;
    draw(g: any) { g.drawImage2(source, 1, 2, 3, 4); if (this.fail) throw new Error('paint failed'); }
  }
  const scope = { document: { createElement }, AscFormat: { CImageShape: Shape },
    AscCommon: { getFullImageSrc2: (id: string) => id },
    Asc: { editor: { ImageLoader: { map_image_index: { photo: { Image: source } } } } } } as unknown as Window;
  return { scope, Shape, source, canvas, createElement, paint };
}

it('uses a cached rendering copy without rewriting the source or effect and restores after failure', () => {
  const s = setup(); installImageSoftEdgeGuard(s.scope);
  const shape = new s.Shape(), before = structuredClone(shape), g = { drawImage2: s.paint };
  shape.draw(g); shape.draw(g);
  expect(s.paint).toHaveBeenLastCalledWith(s.canvas, 1, 2, 3, 4);
  expect(s.createElement).toHaveBeenCalledTimes(1);
  expect(shape).toMatchObject(before);
  shape.fail = true;
  expect(() => shape.draw(g)).toThrow('paint failed');
  expect(g.drawImage2).toBe(s.paint);
  const patched = s.Shape.prototype.draw;
  installImageSoftEdgeGuard(s.scope);
  expect(s.Shape.prototype.draw).toBe(patched);
  expect(installImageSoftEdgeGuard({} as Window)).toBe(false);
});

it.each(['rectangle', 'crop', 'effects', 'shadow', 'radius', 'loading'])('preserves unsupported %s rendering', (kind) => {
  const s = setup(); installImageSoftEdgeGuard(s.scope);
  const shape = new s.Shape();
  if (kind === 'rectangle') shape.spPr.geometry.preset = 'rect';
  if (kind === 'crop') shape.blipFill.srcRect = {};
  if (kind === 'effects') shape.blipFill.Effects = [{}];
  if (kind === 'shadow') shape.spPr.effectProps.EffectLst.outerShdw = {};
  if (kind === 'radius') shape.spPr.effectProps.EffectLst.softEdge.rad = -1;
  if (kind === 'loading') s.source.complete = false;
  shape.draw({ drawImage2: s.paint });
  expect(s.paint.mock.calls[0][0]).toBe(s.source);
  expect(s.createElement).not.toHaveBeenCalled();
});
