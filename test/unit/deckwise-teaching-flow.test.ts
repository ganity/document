import { afterEach, expect, it } from 'vitest';
import { LanguageCode, setLanguage } from '@ranuts/shared/i18n';
import { speechPauseMode, speechPauseLabel, teachingStepSummary } from '../../lib/deckwise-teaching-flow';

afterEach(() => setLanguage(LanguageCode.EN));

it('uses raw Host milliseconds for short pauses and 3 to 30 second timed waits', () => {
  expect(speechPauseMode()).toBe('none');
  expect(speechPauseMode(0)).toBe('none');
  expect(speechPauseMode(400)).toBe('pause');
  expect(speechPauseMode(2000)).toBe('pause');
  expect(speechPauseMode(3000)).toBe('timed');
  expect(speechPauseMode(30000)).toBe('timed');
  expect(speechPauseLabel(0)).toBe('');
  expect(speechPauseLabel(30000)).toContain('30');
  expect(speechPauseLabel(4500)).toContain('4.5');
});

it('summarizes narration without inserting authored text as markup', () => {
  const text = '<img src=x onerror=alert(1)>';
  expect(teachingStepSummary({ type: 'speech', text })).toBe(text);
  expect(teachingStepSummary({ type: 'speech', text: ' ' })).toBeTruthy();
});

it('resizes the native layout while teaching is open and restores it on close or dispose', async () => {
  const { createTeachingPanelLayout } = await import('../../lib/deckwise-teaching-flow');
  document.body.innerHTML = '<div id="right-menu"><div class="right-panel" style="padding-right:40px"></div></div>';
  const menu = document.getElementById('right-menu')!;
  let width = 260;
  const changes: string[] = [];
  const native = window as any;
  native.PE = {
    getController: (name: string) =>
      name === 'RightMenu'
        ? {
            getView: () => ({
              setInnerWidth: (value: number) => {
                width = value + 40;
              },
            }),
          }
        : { onLayoutChanged: (item: string) => changes.push(item) },
  };
  const bounds = menu.getBoundingClientRect.bind(menu);
  menu.getBoundingClientRect = () => ({ ...bounds(), width });
  const layout = createTeachingPanelLayout(document);
  const originalViewport = window.innerWidth;
  try {
    layout.sync(true);
    expect(width).toBe(400);
    layout.sync(true);
    expect(changes).toEqual(['rightmenu']);
    Object.defineProperty(window, 'innerWidth', { value: 700, configurable: true });
    window.dispatchEvent(new Event('resize'));
    expect(width).toBe(260);
    Object.defineProperty(window, 'innerWidth', { value: originalViewport, configurable: true });
    window.dispatchEvent(new Event('resize'));
    expect(width).toBe(400);
    layout.sync(false);
    expect(width).toBe(260);
    layout.sync(true);
    layout.dispose();
    expect(width).toBe(260);
  } finally {
    Object.defineProperty(window, 'innerWidth', { value: originalViewport, configurable: true });
    layout.dispose();
    delete native.PE;
    document.body.replaceChildren();
  }
});
