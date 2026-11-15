import { expect, it, vi } from 'vitest';
import { installNestedSlideAnimations } from '../../lib/onlyoffice/guards/nested-slide-animations';

it('keeps attached nested animation targets during native import without accepting detached or cyclic shapes', () => {
  class Target {
    spid = '26'; subSpId: string | null = null;
    parent = { onRemoveChild: vi.fn() };
    setSpid = vi.fn(); setSubSpId = vi.fn();
    assignConnection(_shapes: Record<string, unknown>) { this.parent.onRemoveChild(); }
  }
  const scope = { AscFormat: { CSpTgt: Target } };
  expect(installNestedSlideAnimations(scope as unknown as Window)).toBe(true);
  expect(installNestedSlideAnimations(scope as unknown as Window)).toBe(true);
  const slide = { cSld: { spTree: [] as unknown[] } };
  const group: { parent: typeof slide; spTree: unknown[]; group?: unknown } = { parent: slide, spTree: [] };
  const child = { Id: 'live-child', group };
  slide.cSld.spTree.push(group); group.spTree.push(child);
  const target = new Target(); target.subSpId = '27';
  target.assignConnection({ '26': child, '27': { Id: 'sub-shape' } });
  expect(target.setSpid).toHaveBeenCalledWith('live-child');
  expect(target.setSubSpId).toHaveBeenCalledWith('sub-shape');
  expect(target.parent.onRemoveChild).not.toHaveBeenCalled();
  group.spTree.length = 0;
  target.assignConnection({ '26': child });
  expect(target.parent.onRemoveChild).toHaveBeenCalledOnce();
  group.spTree.push(child, group); group.group = group;
  target.assignConnection({ '26': child });
  expect(target.parent.onRemoveChild).toHaveBeenCalledTimes(2);
});
