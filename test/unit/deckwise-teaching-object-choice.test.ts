import { afterEach, beforeEach, expect, it, vi } from 'vitest';
import { LanguageCode, setLanguage } from '@ranuts/shared/i18n';
const state = vi.hoisted(() => ({ api: {} as any }));
vi.mock('../../lib/agent-plugin/editor-bridge', () => ({ getEditorApi: () => state.api }));
import { teachingObjectChoices, appendTeachingObjectChoice } from '../../lib/deckwise-teaching-object-choice';
beforeEach(() => setLanguage(LanguageCode.ZH));
afterEach(() => {
  document.body.replaceChildren();
  setLanguage(LanguageCode.EN);
});
it('uses Chinese semantics, real slide text or visual type instead of English descriptions, and locates the exact mapped object', () => {
  const shapes = ['a', 'b', 'c'].map((name) => ({
    getUniNvProps: () => ({ cNvPr: { name } }),
    isImage: () => true,
    getText: () => (name === 'b' ? 'Hello, Teddy Bear!' : ''),
    getBase64Img: vi.fn(() => 'data:image/png;base64,AA=='),
  }));
  const controller = { resetSelection: vi.fn(), selectObject: vi.fn() };
  state.api = {
    WordControl: {
      m_oLogicDocument: { Slides: [{ cSld: { spTree: shapes } }], GetCurrentController: () => controller },
    },
  };
  const choices = teachingObjectChoices(
    1,
    shapes.map((_, i) => ({
      objectId: String(i),
      sourceObjectName: ['a', 'b', 'c'][i],
      status: 'mapped',
      effectTargetEligible: true,
      role: 'Brown teddy bear waving a paw; the main character',
      aliases: i === 0 ? ['挥手的小熊'] : [],
    })),
  );
  expect(choices.map((choice) => choice.name)).toEqual(['挥手的小熊', '文字「Hello, Teddy Bear!」', '图片 3']);
  const box = document.createElement('input');
  box.type = 'checkbox';
  const load = appendTeachingObjectChoice(document, document.body, box, choices[0], vi.fn());
  expect(shapes[0].getBase64Img).not.toHaveBeenCalled();
  load();
  load();
  expect(shapes[0].getBase64Img).toHaveBeenCalledOnce();
  expect(document.querySelector('img')?.getAttribute('src')).toBe('data:image/png;base64,AA==');
  document.querySelector('button')!.click();
  expect(controller.resetSelection).toHaveBeenCalledOnce();
  expect(controller.selectObject).toHaveBeenCalledWith(shapes[0], 0);
  expect(box.checked).toBe(false);
});
it('keeps unavailable objects selectable but rejects unsafe thumbnail sources and locating stale mappings', () => {
  state.api = {};
  const [choice] = teachingObjectChoices(1, [
    {
      objectId: 'missing',
      sourceObjectName: 'missing',
      status: 'mapped',
      effectTargetEligible: true,
      role: 'Long English description',
    },
  ]);
  expect(choice.name).toBe('图形 1');
  expect(choice.thumbnail()).toBeUndefined();
  expect(() => choice.locate()).toThrow();
});
