import { describe, expect, it, vi } from 'vitest';
const fixture = vi.hoisted(() => ({ api: {} as Record<string, unknown>, current: true }));
vi.mock('../../lib/agent-plugin/editor-bridge', () => ({ requireEditorContext: () => ({ api: fixture.api }) }));
vi.mock('../../lib/agent-plugin/document-tool-action', () => ({ captureDocumentToolTarget: () => ({ context: { kind: 'slide' }, isCurrent: () => fixture.current }) }));
import { readDeckWiseSelectedObjects } from '../../lib/deckwise-object-selection';
function setup() {
  const shape = { nvPicPr: {}, spPr: { xfrm: { offX: 10, offY: 20, extX: 30, extY: 40 } } };
  const shapes = [shape], selectedObjects = [shape]; fixture.current = true;
  fixture.api = { getCurrentPage: () => 0, WordControl: { m_oLogicDocument: { GetWidthMM:()=>100, GetHeightMM:()=>100, Slides: [{ cSld: { spTree: shapes } }], GetCurrentController: () => ({ selectedObjects }) } } };
  return { shape, shapes, selectedObjects };
}
describe('native selected object fingerprints', () => {
  it('returns normalized current geometry without using regenerated runtime IDs', () => {
    setup(); expect(readDeckWiseSelectedObjects()).toEqual([{kind:'image',text:'',rect:{x:.1,y:.2,width:.3,height:.4}}]);
  });
  it('excludes stale, repeated, grouped, absent and excessive selection', () => {
    setup(); fixture.current = false; expect(readDeckWiseSelectedObjects()).toEqual([]);
    const f=setup();f.selectedObjects.push(f.shape);expect(readDeckWiseSelectedObjects()).toEqual([]);
    const g=setup();Object.assign(g.shape,{group:{}});expect(readDeckWiseSelectedObjects()).toEqual([]);
    const h=setup();h.shapes.length=0;expect(readDeckWiseSelectedObjects()).toEqual([]);
    const j=setup();j.selectedObjects.push(...Array(8).fill(j.shape));expect(readDeckWiseSelectedObjects()).toEqual([]);
  });
});
