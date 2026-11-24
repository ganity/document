import { expect, it, vi } from 'vitest';
const state = vi.hoisted(() => ({ context: null as any }));
vi.mock('../../lib/agent-plugin/editor-bridge', () => ({
  requireEditorContext: () => state.context,
  getEditorApi: () => state.context?.api,
}));
import { prepareSpeechNativeSchedule } from '../../lib/deckwise-teaching-native-schedule';

function fixture(kinds: number[]) {
  const createEffect = vi.fn(() => {
    throw new Error('Playback must never invent an animation');
  });
  const effect = (kind: number, Id: string, delay: number): any => ({
    Id,
    cTn: { presetClass: kind, presetID: 10, presetSubtype: 0, setNodeType() {} },
    getObjectId: () => 'shape',
    asc_getDuration: () => 600,
    asc_getDelay: () => delay,
    asc_putDelay() {},
    isIdle: () => true,
    isAtEnd: () => false,
    createDuplicate: () => effect(kind, `${Id}-copy`, delay),
  });
  const effects = kinds.map((kind, index) => effect(kind, `effect-${index}`, index * 1000));
  const saved = { getAllAnimEffects: () => effects };
  const slide = {
    timing: saved,
    cSld: { spTree: [{ Id: 'shape', getUniNvProps: () => ({ cNvPr: { name: 'Title' } }) }] },
  };
  class Timing {
    effects: any[] = [];
    createEffect = createEffect;
    getAllAnimEffects() {
      return this.effects;
    }
    buildTree(entries: any[][]) {
      this.effects = entries.flatMap((entry) => entry.slice(1));
    }
    setParent() {}
  }
  state.context = {
    api: { WordControl: { m_oLogicDocument: { Slides: [slide] } } },
    AscFormat: { CTiming: Timing, ExecuteNoHistory: (operation: () => unknown) => operation() },
  };
  const actions: Parameters<typeof prepareSpeechNativeSchedule>[1] = [
    {
      id: 'speech',
      type: 'speech',
      visualCues: kinds.length
        ? kinds.map((kind, index) => ({
            id: `cue-${index}`,
            atMs: index * 1000,
            endMs: index * 1000 + 600,
            effect: { type: 'element_animation', elementId: 'logical', effect: kind === 1 ? 'fadeIn' : 'fadeOut' },
          }))
        : [
            {
              id: 'cue',
              atMs: 0,
              endMs: 600,
              effect: { type: 'element_animation', elementId: 'logical', effect: 'fadeIn' },
            },
          ],
    },
  ];
  const run = () =>
    prepareSpeechNativeSchedule(
      1,
      actions,
      [{ objectId: 'logical', status: 'mapped', effectTargetEligible: true, sourceObjectName: 'Title' }],
      {
        nativeMotionPreset: (name) => ({
          presetClass: name === 'fadeIn' ? 1 : 2,
          presetId: 10,
          subtype: 0,
          ooxmlClass: name === 'fadeIn' ? 'entr' : 'exit',
        }),
      },
    );
  return { run, slide, saved, createEffect };
}

it('refuses a missing persisted native animation instead of creating a playback-only effect', () => {
  const f = fixture([]);
  expect(f.run).toThrow('尚未写入 PPT');
  expect(f.createEffect).not.toHaveBeenCalled();
  expect(f.slide.timing).toBe(f.saved);
});

it('binds entrance and exit on the same object to their actual saved native effects', () => {
  const f = fixture([1, 2]),
    schedule = f.run()!;
  expect(schedule.scheduled.map((entry) => entry.effect.Id)).toEqual(['effect-0-copy', 'effect-1-copy']);
  expect(schedule.unboundCount).toBe(0);
  expect(f.createEffect).not.toHaveBeenCalled();
  schedule.install();
  expect(f.slide.timing).toBe(schedule.timing);
  schedule.restore();
  expect(f.slide.timing).toBe(f.saved);
});

it('identifies a stale native delay before starting the speech', () => {
  const f = fixture([1, 2]);
  f.saved.getAllAnimEffects()[1].asc_getDelay = () => 0;
  expect(f.run).toThrow('与讲解时间不一致');
  expect(f.slide.timing).toBe(f.saved);
});

it('suppresses a page-entry exit on a speech-owned ancestor and preserves a sibling animation', () => {
  const f = fixture([2]);
  const child = f.slide.cSld.spTree[0];
  const group = { Id: 'group', spTree: [child], getUniNvProps: () => ({ cNvPr: { name: 'Group' } }) };
  Object.assign(child, { group });
  f.slide.cSld.spTree.splice(0, 1, group);
  const saved = f.saved.getAllAnimEffects();
  const additional = (id: string) => ({ ...saved[0], getObjectId: () => id });
  saved.push(additional('group'), additional('unrelated'));
  const schedule = f.run()!;
  expect(schedule.unboundCount).toBe(1);
  expect(schedule.scheduled).toHaveLength(1);
  expect(saved).toHaveLength(3);
});
