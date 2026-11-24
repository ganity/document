import { expect, it } from 'vitest';
import { splitBackgroundPixels } from '../../lib/agent-plugin/slide-background-region';

it('moves exactly the original region pixels and restores the original image when composited', () => {
  const width = 12,
    height = 8;
  const original = Uint8ClampedArray.from({ length: width * height * 4 }, (_, i) => (i % 4 === 3 ? 255 : i % 251));
  const saved = original.slice();
  const result = splitBackgroundPixels(original, width, height, [0.25, 0.25, 0.5, 0.5]);
  expect(result.box).toEqual([3, 2, 6, 4]);
  const restored = result.background.slice();
  for (let y = 0; y < 4; y++)
    for (let x = 0; x < 6; x++) {
      const offset = ((y + 2) * width + x + 3) * 4;
      expect([...result.background.subarray(offset, offset + 4)]).toEqual([255, 255, 255, 255]);
      restored.set(result.region.subarray((y * 6 + x) * 4, (y * 6 + x + 1) * 4), offset);
    }
  expect(restored).toEqual(original);
  expect(original).toEqual(saved);
});

for (const box of [
  [0, 0, 1, 1],
  [-0.1, 0, 0.5, 0.5],
  [0.8, 0, 0.5, 0.5],
  [0, 0, 0, 0.5],
  [0, 0, NaN, 0.5],
])
  it(`rejects invalid or whole-slide extraction: ${box}`, () => {
    expect(() => splitBackgroundPixels(new Uint8ClampedArray(12 * 8 * 4).fill(255), 12, 8, box)).toThrow();
  });

it('rejects transparent backgrounds instead of changing their composition', () => {
  expect(() => splitBackgroundPixels(new Uint8ClampedArray(12 * 8 * 4), 12, 8, [0.2, 0.2, 0.5, 0.5])).toThrow();
});
