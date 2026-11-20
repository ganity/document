import { beforeEach, describe, expect, it, vi } from 'vitest';
import { prepareQuizProjection } from '../../lib/agent-plugin/slide-quiz-projection';

const state = vi.hoisted(() => ({ editor: undefined as any, preload: vi.fn() }));
vi.mock('../../lib/agent-plugin/editor-bridge', () => ({
  requireEditorContext: () => state.editor,
  getEditorContext: () => state.editor,
}));
vi.mock('../../lib/onlyoffice/readonly', () => ({ getReadonlyMode: () => false }));
vi.mock('../../lib/agent-plugin/slide-image-png', () => ({
  validateSlidePng: () => ({ width: 1, height: 1 }),
  preloadSlidePng: state.preload,
}));

function fixture(grouped = true) {
  const text = (id: string, content: string) => ({
    type: 'native_text',
    id: `page-3_${id}`,
    content,
    bboxPx: [10, 10, 100, 30],
    style: { fontSizePt: 16, fontWeight: 400, color: '#000000', fontFace: 'Arial' },
  });
  const previous = {
    slideId: 'page-3',
    canvas: { width: 1280, height: 720 },
    elements: [text('q0_question', 'Which object makes a sound?'), text('q0_option0_label', 'A bell')],
  };
  const shape = (name: string, value = ''): any => ({
    name,
    getObjectName() {
      return this.name;
    },
    setName(n: string) {
      this.name = n;
    },
    setParent(parent: any) {
      this.parent = parent;
    },
    getText: () => value,
  });
  const leaves = previous.elements.map((e) => shape(e.id.replace('page-3', 'assessment-3'), e.content));
  leaves[0].getText = () => 'Which object makes \r\na sound?';
  const group = { ...shape('assessment'), spTree: leaves };
  const before = shape('unrelated-before'),
    after = shape('unrelated-after');
  const originals = [before, ...(grouped ? [group] : leaves), after];
  const slide = {
    cSld: { spTree: [...originals] },
    removeFromSpTreeByPos(i: number) {
      this.cSld.spTree.splice(i, 1);
    },
    addToSpTreeToPos(i: number, s: any) {
      this.cSld.spTree.splice(i, 0, s);
    },
  };
  const paragraph = () => ({
    text: '',
    AddText(t: string) {
      this.text += t;
    },
    SetFontSize() {},
    SetBold() {},
    SetColor() {},
    SetFontFamily() {},
  });
  const builder = {
    GetPresentation: () => ({
      GetSlideByIndex: () => ({
        AddObject(s: any) {
          slide.cSld.spTree.push(s.Drawing);
          return true;
        },
      }),
    }),
    CreateRGBColor() {},
    CreateSolidFill() {},
    CreateNoFill() {},
    CreateStroke() {},
    CreateParagraph: paragraph,
    CreateShape() {
      let paragraphs = [paragraph()];
      const drawing = shape('');
      drawing.getText = () => paragraphs.map((p) => p.text).join('\r\n');
      return {
        Drawing: drawing,
        SetPosition() {},
        SetVerticalTextAlign() {},
        SetPaddings() {},
        GetDocContent: () => ({
          RemoveAllElements() {
            paragraphs = [paragraph()];
          },
          GetElement: () => paragraphs[0],
          Push(p: ReturnType<typeof paragraph>) {
            paragraphs.push(p);
          },
        }),
      };
    },
  };
  state.editor = {
    api: {
      canEdit: () => true,
      isSlideShow: () => false,
      WordControl: {
        m_oLogicDocument: {
          Slides: [slide],
          CanEdit: () => true,
          IsMasterMode: () => false,
          IsSelectionLocked: () => false,
          GetWidthMM: () => 320,
          GetHeightMM: () => 180,
          GetCurrentController: () => ({ createImage: () => shape('') }),
          Recalculate() {},
          Document_UpdateInterfaceState() {},
          DrawingDocument: { UpdateThumbnailsAttack() {} },
        },
      },
    },
    presentationBuilder: builder,
    AscCommon: { History: { Index: 0, Points: [{}], _getLongPointIndex: () => -1 } },
  };
  const next = structuredClone(previous) as any;
  next.elements.push({
    type: 'image_layer',
    id: 'page-3_q0_option0_image',
    assetId: 'image',
    bboxPx: [10, 50, 100, 100],
  });
  return {
    slide,
    group,
    leaves,
    originals,
    previous,
    next,
    input: { page: 1, previous, next, assets: { image: new Uint8Array([1]) } },
  };
}

beforeEach(() => state.preload.mockReset().mockResolvedValue({ full: 'native-image', ready() {} }));

describe('native Quiz projection recovery', () => {
  it.each([true, false])(
    'adopts an image with wrapped text and a different template prefix (grouped=%s)',
    async (grouped) => {
      const f = fixture(grouped);
      const prepared = await prepareQuizProjection(f.input);
      expect(f.slide.cSld.spTree).toEqual(f.originals);
      prepared.apply();
      expect(f.slide.cSld.spTree.map((s) => s.getObjectName())).toEqual(
        grouped
          ? [
              'unrelated-before',
              'assessment-3_q0_question',
              'assessment-3_q0_option0_label',
              'assessment-3_q0_option0_image',
              'unrelated-after',
            ]
          : [
              'unrelated-before',
              'assessment-3_q0_question',
              'assessment-3_q0_option0_label',
              'unrelated-after',
              'assessment-3_q0_option0_image',
            ],
      );
      prepared.restore();
      expect(f.slide.cSld.spTree).toEqual(f.originals);
      expect(f.group.spTree).toEqual(f.leaves);
    },
  );

  it('preserves a group with an extra user object by rejecting adoption', async () => {
    const f = fixture();
    f.group.spTree.push({ getObjectName: () => 'user-note' });
    await expect(prepareQuizProjection(f.input)).rejects.toThrow('binding');
    expect(f.slide.cSld.spTree).toEqual(f.originals);
    expect(state.preload).not.toHaveBeenCalled();
  });

  it('rejects duplicate native identities', async () => {
    const f = fixture();
    f.group.spTree.push(f.leaves[0]);
    await expect(prepareQuizProjection(f.input)).rejects.toThrow('Duplicate');
  });

  it('detects nested edits while preparing an image without modifying the slide', async () => {
    const f = fixture();
    state.preload.mockImplementation(async () => {
      f.group.spTree.reverse();
      return { full: 'native-image', ready() {} };
    });
    await expect(prepareQuizProjection(f.input)).rejects.toThrow('changed during image preparation');
    expect(f.slide.cSld.spTree).toEqual(f.originals);
  });
});

it('refreshes a stale managed hint even when previous and next use the same upgraded template', async () => {
  const f = fixture(false);
  f.previous.elements[1].id = 'page-3_hint';
  f.previous.elements[1].content = 'Write your answer.';
  f.leaves[1].setName('assessment-3_hint');
  f.leaves[1].getText = () => 'Choose the best answer.';
  const prepared = await prepareQuizProjection({ ...f.input, next: structuredClone(f.previous) });
  prepared.apply();
  expect(f.slide.cSld.spTree.find(s => s.getObjectName() === 'assessment-3_hint').getText()).toBe('Write your answer.');
  expect(f.slide.cSld.spTree).toContain(f.originals[0]);
  prepared.restore();
  expect(f.slide.cSld.spTree).toEqual(f.originals);
});
