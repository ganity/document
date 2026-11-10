import { deckwisePageLineage } from '../../lib/deckwise-page-lineage';
import { expect, it, vi } from 'vitest';
import { attentionGeometry, selectedAttentionTargets } from '../../lib/deckwise-teaching-preview';
const object = { objectId: 'logical', sourceShapeId: '42', status: 'mapped', effectTargetEligible: true };
const shape = (id = 42) => ({
  getUniNvProps: () => ({ cNvPr: { id } }),
  extX: 20,
  extY: 10,
  transform: { TransformPointX: (x: number, _y: number) => x + 10, TransformPointY: (_: number, y: number) => y + 5 },
});
it('uses saved identity and current native coordinates at different zooms and offsets', () => {
  for (const zoom of [0.5, 1, 1.5]) {
    const native = shape();
    const result = attentionGeometry([native], [object], ['logical'], 100, 50, (x, y) => ({
      X: 37 + x * zoom,
      Y: 81 + y * zoom,
    }));
    expect(result.targets[0].shape).toBe(native);
    expect(result.targets[0].rect).toEqual({ x: 9.6, y: 9.4, w: 20.8, h: 21.2 });
  }
});
it('bounds all four transformed corners for rotated targets', () => {
  const native = shape();
  native.transform = { TransformPointX: (_, y) => 30 - y, TransformPointY: (x) => 10 + x };
  const result = attentionGeometry([native], [object], ['logical'], 100, 50, (x, y) => ({ X: x, Y: y }));
  expect(result.targets[0].rect).toEqual({ x: 19.6, y: 19.4, w: 10.8, h: 41.2 });
});
it('rejects missing, ambiguous, changed, grouped or non-finite targets without ordinal fallback', () => {
  for (const [tree, objects] of [
    [[], [object]],
    [[shape(1)], [object]],
    [[shape(), shape()], [object]],
    [[shape()], [{ ...object, status: 'missing' }]],
    [[shape()], [{ ...object, sourceShapeId: undefined }]],
    [[{ ...shape(), group: {} }], [object]],
    [[{ ...shape(), extX: NaN }], [object]],
  ] as const) {
    expect(() =>
      attentionGeometry([...tree], [...objects], ['logical'], 100, 50, (x, y) => ({ X: x, Y: y })),
    ).toThrow();
  }
});

it('uses an explicitly mapped unique compiler name after native IDs are rewritten', () => {
  const native = { ...shape(987), getUniNvProps: () => ({ cNvPr: { id: 987, name: 'compiler-layer' } }) };
  const proof = { ...object, sourceObjectName: 'compiler-layer' };
  const read = (tree = [native], objects = [proof]) =>
    attentionGeometry(tree, objects, ['logical'], 100, 50, (x, y) => ({ X: x, Y: y }));
  expect(read().targets[0].shape).toBe(native);
  expect(() => read([native, native])).toThrow();
  expect(() => read([native], [proof, { ...proof, objectId: 'peer' }])).toThrow();
});

it('maps canvas multi-selection through saved object identities, never drawing order', () => {
  const panda = shape(),
    monkey = shape(43);
  const objects = [object, { ...object, objectId: 'monkey', sourceShapeId: '43' }];
  expect(selectedAttentionTargets([panda, monkey], [monkey, panda], objects)).toEqual(['monkey', 'logical']);
  for (const selected of [[], [shape()], [{ ...panda, group: {} }], [panda, panda]])
    expect(() => selectedAttentionTargets([panda, monkey], selected, objects)).toThrow();
  expect(() => selectedAttentionTargets([panda], [panda], [...objects, object])).toThrow();
  expect(() => selectedAttentionTargets([panda], [panda], [{ ...object, effectTargetEligible: false }])).toThrow();
});

it('resolves renamed and same-name shapes by captured package identity and never binds a copy', () => {
  const original = { ...shape(987), getUniNvProps: () => ({ cNvPr: { id: 987, name: 'renamed' } }) };
  const peer = { ...shape(988), getUniNvProps: () => ({ cNvPr: { id: 988, name: 'renamed' } }) };
  const copy = { ...original };
  const capability = vi.spyOn(deckwisePageLineage, 'hasSourceIdentity').mockReturnValue(true);
  const identity = vi.spyOn(deckwisePageLineage, 'sourceShapeId').mockImplementation(s =>
    s === original ? '42' : s === peer ? '43' : undefined);
  try {
    const proofs = [object, { ...object, objectId: 'peer', sourceShapeId: '43' }];
    expect(selectedAttentionTargets([original, peer, copy], [peer, original], proofs)).toEqual(['peer', 'logical']);
    expect(attentionGeometry([copy, peer, original], proofs, ['logical'], 100, 50, (X, Y) => ({ X, Y })).targets[0].shape).toBe(original);
    expect(() => selectedAttentionTargets([copy, peer], [copy], proofs)).toThrow();
    expect(() => attentionGeometry([copy, peer], proofs, ['logical'], 100, 50, (X, Y) => ({ X, Y }))).toThrow();
  } finally { capability.mockRestore(); identity.mockRestore(); }
});
