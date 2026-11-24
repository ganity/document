import { afterEach, beforeEach, expect, it, vi } from 'vitest';
import { LanguageCode, loadLocale, setLanguage } from '@ranuts/shared/i18n';
import { createTeachingAnimationForm } from '../../lib/deckwise-teaching-animation-form';
import { relocateAnimationTrigger } from '../../lib/deckwise-teaching-animation-bindings';
import type { AnimationCue } from '../../lib/deckwise-teaching-animation-bindings';
beforeEach(async () => {
  await loadLocale(LanguageCode.ZH);
  setLanguage(LanguageCode.ZH);
});
afterEach(() => {
  document.body.replaceChildren();
  setLanguage(LanguageCode.EN);
});
function fixture(cues: AnimationCue[] = []) {
  let clock = 0;
  const save = vi.fn(),
    error = vi.fn();
  const node = createTeachingAnimationForm({
    doc: document,
    action: { id: 'speech', text: '小狗说 Hello', audioAssetId: 'audio', durationMs: 5000, visualCues: cues },
    editable: true,
    objects: [{ objectId: 'dog', name: '小狗' }],
    selection: () => ({ start: 0, end: 2 }),
    native: () => [
      { index: 0, objectId: 'dog', effect: 'slideInLeft', durationMs: 800, supported: true, owner: '第2句' },
    ],
    selectedObjects: () => ['dog'],
    preview: async () => {},
    position: () => clock,
    save,
    error,
  });
  document.body.append(node);
  const click = (text: string) =>
    [...node.querySelectorAll<HTMLButtonElement>('button')]
      .find((b) => !b.closest('[hidden]') && (b.getAttribute('aria-label') === text || b.textContent === text))!
      .click();
  const select = (label: string, value: string) => {
    const n = node.querySelector<HTMLSelectElement>(`[aria-label="${label}"]`)!;
    n.value = value;
    n.dispatchEvent(new Event('change'));
  };
  return {
    node,
    save,
    error,
    click,
    select,
    clock: (ms: number) => {
      clock = ms;
    },
  };
}
it.each(['add', 'link'])(
  'saves all four audio trigger positions through %s and preserves selection after focus changes',
  (kind) => {
    for (const [position, phase, expected] of [
      ['start', 'with', 0],
      ['end', 'with', 5000],
      ['text', 'with', 1200],
      ['text', 'after', 1700],
    ] as const) {
      const f = fixture();
      f.click(kind === 'add' ? '添加动画' : '关联动画');
      f.select('什么时候播放动画', position);
      f.select('执行时机', phase);
      if (position === 'text') {
        f.clock(1200);
        f.click('标记词开始');
        f.clock(1700);
        f.click('标记词结束');
      }
      if (kind === 'link') {
        const c = f.node.querySelector<HTMLInputElement>('input[type=checkbox]')!;
        c.checked = true;
        c.dispatchEvent(new Event('change'));
        expect(f.node.textContent).toContain('将移到此处');
      }
      f.node.querySelector('form')!.dispatchEvent(new Event('submit', { cancelable: true }));
      expect(f.error).not.toHaveBeenCalled();
      expect(f.save).toHaveBeenCalledOnce();
      expect(f.save.mock.calls[0][0][0]).toMatchObject({
        atMs: expected,
        trigger: { kind: position },
        effect: { elementId: 'dog' },
      });
      if (kind === 'link') expect(f.save.mock.calls[0][0][0].nativeAnimation).toBe('dog');
      f.node.remove();
    }
  },
);
it('rejects unmarked selected words, cancels staged checks, and unlinks without editing native effects', () => {
  const cue: AnimationCue = {
    id: 'linked',
    nativeAnimation: 'dog',
    atMs: 0,
    endMs: 800,
    trigger: { version: 1, kind: 'start' },
    effect: { type: 'element_animation', elementId: 'dog', effect: 'fadeIn' },
  };
  const f = fixture([cue]);
  f.click('添加动画');
  f.node.querySelector('form')!.dispatchEvent(new Event('submit', { cancelable: true }));
  expect(f.error).toHaveBeenCalled();
  expect(f.save).not.toHaveBeenCalled();
  f.click('取消');
  f.click('关联动画');
  f.node.querySelector<HTMLInputElement>('input[type=checkbox]')!.checked = true;
  f.click('取消');
  expect(f.save).not.toHaveBeenCalled();
  f.click('取消关联');
  expect(f.save).toHaveBeenCalledWith([], []);
});
it('keeps a semantic range after unique relocation and marks old timing stale', () => {
  expect(
    relocateAnimationTrigger(
      { version: 1, kind: 'text', text: '小狗', start: 0, end: 2, phase: 'after', audioAssetId: 'a' },
      '看，小狗',
    ),
  ).toMatchObject({ start: 2, end: 4, stale: true });
});

it('explains trigger choices and summarizes objects with translated effects and duration units', () => {
  const f = fixture([
    {
      id: 'fade',
      atMs: 0,
      endMs: 600,
      trigger: { version: 1, kind: 'start' },
      effect: { type: 'element_animation', elementId: 'dog', effect: 'fadeIn' },
    },
  ]);
  const row = f.node.querySelector('.teaching-animation-edit')!;
  expect(row.textContent).toContain('开始讲这句时');
  expect(row.textContent).toContain('小狗 · 淡入 · 0.6 秒');
  expect(row.textContent).not.toContain('fadeIn');
  expect(row.getAttribute('title')).toContain('小狗');
  f.click('添加动画');
  expect(f.node.querySelector('.teaching-animation-help')!.textContent).toContain('先在上方讲解词');
  f.select('什么时候播放动画', 'end');
  expect(f.node.querySelector('.teaching-animation-help')!.textContent).toContain('全部说完后');
  expect(f.node.querySelector('form')!.textContent).toContain('持续时间（秒）');
});
