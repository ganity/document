import { expect, it, vi } from 'vitest';
const state = vi.hoisted(() => ({ api: null as any, dirty: false }));
vi.mock('../../lib/agent-plugin/editor-bridge', () => ({
  getEditorApi: () => state.api,
  requireEditorContext: () => ({ api: state.api, AscCommon: { changestype_Timing: 7 }, AscDFH: {} }),
}));
vi.mock('../../lib/unsaved-guard', () => ({ hasUnsavedChanges: () => state.dirty }));
vi.mock('../../lib/agent-plugin/slide-image-transform', () => ({
  slideImageFingerprint: (s: object) => JSON.stringify(s),
}));
import {
  integrateGeneratedNativeMotions,
  generatedMotionBindings,
  nativeSpeechMotionSteps,
} from '../../lib/deckwise-teaching-native-motion';
import { setReadonlyState } from '../../lib/onlyoffice/readonly';
import type { Runtime } from '../../lib/deckwise-teaching-preview';
const runtime = {
  nativeMotionPreset: () => ({ presetClass: 1, presetId: 10, subtype: 0, ooxmlClass: 'entr' }),
} satisfies Pick<Runtime, 'nativeMotionPreset'>;
function setup() {
  state.dirty = false;
  setReadonlyState(false);
  const shapes = ['one', 'two'].map((Id) => ({ Id, getUniNvProps: () => ({ cNvPr: { name: Id } }) }));
  const effects: any[] = [];
  const controller = {
    selectedObjects: [] as typeof shapes,
    resetSelection() {
      this.selectedObjects = [];
    },
    selectObject(s: (typeof shapes)[number]) {
      this.selectedObjects.push(s);
    },
  };
  const timing = {
    getAllAnimEffects: () => effects,
    getEffectsSequences: () => (effects.length ? [[null, ...effects]] : []),
  };
  const slide = { cSld: { spTree: shapes }, timing };
  state.api = {
    isDocumentLoadComplete: true,
    isLoadFullApi: true,
    canEdit: () => true,
    getCurrentPage: () => 0,
    asc_IsStartedAnimationPreview: () => false,
    asc_registerCallback() {},
    asc_unregisterCallback() {},
    StartDemonstration() {},
    EndDemonstration() {},
    WordControl: {
      DemonstrationManager: { IsMainSeqFinished() {}, OnNextSlideAnimPlayer() {}, StartAnimation() {} },
      m_oLogicDocument: {
        Slides: [slide],
        CanEdit: () => true,
        GetCurrentController: () => controller,
        IsSelectionLocked: () => false,
      },
    },
    asc_AddAnimation: vi.fn((presetClass, presetID, presetSubtype) => {
      const target = controller.selectedObjects[0].Id;
      effects.push(effect({ presetClass, presetID, presetSubtype, nodeType: 2 }, target, 500));
    }),
    asc_SetAnimationProperties: vi.fn((properties) => {
      effects[effects.length - 1] = properties;
    }),
  };
  function effect(cTn: object, id: string, duration: number): any {
    return {
      cTn,
      getObjectId: () => id,
      asc_getDuration: () => duration,
      asc_getDelay: () => 0,
      isIdle: () => true,
      isAtEnd: () => false,
      merge() {},
      createDuplicate() {
        return effect({ ...cTn }, id, duration);
      },
      asc_putDuration(value: number) {
        duration = value;
      },
      asc_putStartType(value: number) {
        Reflect.set(cTn, 'nodeType', value);
      },
    };
  }
  const actions = [
    {
      id: 'speech',
      type: 'speech',
      visualCues: shapes.map((s, i) => ({
        id: `cue${i}`,
        atMs: 0,
        endMs: 600,
        effect: { type: 'element_animation', effect: 'fadeIn', elementId: s.Id },
      })),
    },
  ];
  const objects = shapes.map((s) => ({
    objectId: s.Id,
    status: 'mapped',
    effectTargetEligible: true,
    sourceObjectName: s.Id,
  }));
  return { effects, shapes, controller, actions, objects };
}
it('projects two speech targets as one click with simultaneous native effects, preserves selection, and does not append again', () => {
  const f = setup();
  f.controller.selectedObjects = [f.shapes[1]];
  expect(integrateGeneratedNativeMotions(1, f.actions, f.objects, runtime)).toBe(true);
  expect(f.effects.map((e) => e.cTn.nodeType)).toEqual([2, 7]);
  expect(f.effects.map((e) => e.asc_getDuration())).toEqual([600, 600]);
  expect(f.controller.selectedObjects).toEqual([f.shapes[1]]);
  expect(
    generatedMotionBindings(1, f.actions, f.objects, runtime)?.bindings.every((b, i) =>
      generatedMotionBindings(1, f.actions, f.objects, runtime)!.matches(f.effects[i], b),
    ),
  ).toBe(true);
  expect(integrateGeneratedNativeMotions(1, f.actions, f.objects, runtime)).toBe(false);
  expect(state.api.asc_AddAnimation).toHaveBeenCalledTimes(2);
});
it('rejects dirty and read-only state before writing, and never replaces existing native animation', () => {
  const f = setup();
  state.dirty = true;
  expect(() => integrateGeneratedNativeMotions(1, f.actions, f.objects, runtime)).toThrow();
  state.dirty = false;
  setReadonlyState(true);
  expect(() => integrateGeneratedNativeMotions(1, f.actions, f.objects, runtime)).toThrow();
  expect(state.api.asc_AddAnimation).not.toHaveBeenCalled();
  setReadonlyState(false);
  integrateGeneratedNativeMotions(1, f.actions, f.objects, runtime);
  f.actions[0].visualCues[0].endMs = 500;
  expect(integrateGeneratedNativeMotions(1, f.actions, f.objects, runtime)).toBe(false);
  expect(f.effects[0].asc_getDuration()).toBe(600);
});

it('uses edited native presets and durations and includes extra click steps for the same speech without rewriting PPT', () => {
  const f = setup();
  integrateGeneratedNativeMotions(1, f.actions, f.objects, runtime);
  f.effects[0].asc_putDuration(1200);
  f.effects[0].cTn.presetID = 1;
  const added = f.effects[1].createDuplicate();
  added.asc_putStartType(2);
  f.effects.push(added);
  expect(nativeSpeechMotionSteps(1, f.actions, f.objects, runtime)).toEqual([
    { actionId: 'speech', stepIndex: 0 },
    { actionId: 'speech', stepIndex: 1 },
  ]);
  expect(f.effects[0].asc_getDuration()).toBe(1200);
  expect(state.api.asc_AddAnimation).toHaveBeenCalledTimes(2);
});
it('keeps native steps without a speech target manual instead of guessing or blocking the lesson', () => {
  const f = setup();
  integrateGeneratedNativeMotions(1, f.actions, f.objects, runtime);
  f.actions[0].visualCues.pop();
  expect(nativeSpeechMotionSteps(1, f.actions, f.objects, runtime)).toBeUndefined();
});
it('routes edited presets outside the legacy binding subset through the native schedule', () => {
  const f = setup();
  integrateGeneratedNativeMotions(1, f.actions, f.objects, runtime);
  f.effects[0].cTn.presetID = 37;
  expect(nativeSpeechMotionSteps(1, f.actions, f.objects, runtime)).toBeUndefined();
  expect(integrateGeneratedNativeMotions(1, f.actions, f.objects, runtime)).toBe(false);
});
it('retains successive speech bindings when several speeches animate the same objects', () => {
  const f = setup();
  const second = structuredClone(f.actions[0]);
  second.id = 'second';
  second.visualCues.forEach((cue) => {
    cue.id += '-second';
  });
  f.actions.push(second);
  integrateGeneratedNativeMotions(1, f.actions, f.objects, runtime);
  expect(nativeSpeechMotionSteps(1, f.actions, f.objects, runtime)).toEqual([
    { actionId: 'speech', stepIndex: 0 },
    { actionId: 'second', stepIndex: 1 },
  ]);
});
