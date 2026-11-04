import { requireEditorContext } from './agent-plugin/editor-bridge';
import { attentionTarget, type PreviewObject, type Runtime } from './deckwise-teaching-preview';
import { type AnimationCue } from './deckwise-teaching-animation-bindings';

type Effect = {
  Id: string;
  cTn: { presetClass: number; presetID: number; presetSubtype: number; setNodeType(type: number): void };
  getObjectId(): string;
  asc_getDuration(): number;
  createDuplicate(): Effect;
  asc_putDuration(ms: number): void;
  asc_putDelay(ms: number): void;
  isAtEnd(): boolean;
  isIdle(): boolean;
};
type Timing = {
  parent: unknown;
  getAllAnimEffects(): Effect[];
  buildTree(sequences: Array<Array<Effect | string | null>>): void;
  createEffect(shape: string, kind: number, id: number, subtype: number, color: null): Effect | null;
  setParent(parent: unknown): void;
};
export function prepareSpeechNativeSchedule(
  page: number,
  actions: Array<{ id: string; type: string; visualCues?: AnimationCue[] }>,
  objects: PreviewObject[],
  runtime: Pick<Runtime, 'nativeMotionPreset'>,
) {
  const cues = actions.flatMap((action) =>
    action.type === 'speech'
      ? (action.visualCues ?? [])
          .filter((cue) => cue.effect?.type === 'element_animation')
          .map((cue) => ({ actionId: action.id, cue }))
      : [],
  );
  if (!cues.some(({ cue }) => cue.trigger || cue.nativeAnimation)) return undefined;
  if (actions.some((action) => action.type === 'native_animation'))
    throw new Error('Remove legacy page animation bindings before using speech animation links');
  const context = requireEditorContext();
  const api = context.api;
  const control = Reflect.get(api, 'WordControl');
  if (!control || typeof control !== 'object') return undefined;
  const slide = Reflect.get(control, 'm_oLogicDocument')?.Slides?.[page - 1];
  const saved: Timing | undefined = slide?.timing;
  const existing = saved?.getAllAnimEffects() ?? [];
  const noHistory = context.AscFormat?.ExecuteNoHistory;
  const CTiming = context.AscFormat?.CTiming;
  if (typeof noHistory !== 'function' || typeof CTiming !== 'function')
    throw new Error('Native animation scheduling is unavailable');
  const claimed = new Set<Effect>();
  const build = () => {
    const timing = Reflect.construct(CTiming, []) as Timing;
    const scheduled = cues.map(({ actionId, cue }, index) => {
      const shape = attentionTarget(slide.cSld.spTree, objects, cue.effect.elementId);
      if (!shape.Id) throw new Error('Native animation target is missing');
      const candidates = existing.filter((effect) => effect.getObjectId() === shape.Id);
      let effect: Effect;
      if (cue.nativeAnimation) {
        if (candidates.length !== 1 || claimed.has(candidates[0]))
          throw new Error('Linked native animation is missing or ambiguous');
        claimed.add(candidates[0]);
        effect = candidates[0].createDuplicate();
      } else {
        const preset = runtime.nativeMotionPreset(cue.effect.effect);
        if (!preset) throw new Error('Unsupported teaching animation');
        const mirror =
          !cue.trigger &&
          candidates.find(
            (effect) =>
              !claimed.has(effect) &&
              effect.cTn.presetClass === preset.presetClass &&
              effect.cTn.presetID === preset.presetId &&
              effect.cTn.presetSubtype === preset.subtype,
          );
        if (mirror) claimed.add(mirror);
        if (cue.trigger && candidates.some((effect) => !claimed.has(effect)))
          throw new Error('Link the existing native animation instead of adding a second owner on the same object');
        const created = mirror
          ? mirror.createDuplicate()
          : timing.createEffect(shape.Id, preset.presetClass, preset.presetId, preset.subtype, null);
        if (!created) throw new Error('Native animation preset could not be created');
        effect = created;
        effect.asc_putDuration(cue.endMs - cue.atMs);
      }
      effect.cTn.setNodeType(2);
      effect.asc_putDelay(0);
      return {
        actionId,
        stepIndex: index,
        atMs: cue.atMs,
        endMs: cue.atMs + effect.asc_getDuration(),
        trigger: cue.trigger,
        effect,
        token: effect.Id,
      };
    });
    // Independent interactive sequences reuse the native renderer without advancing unrelated effects.
    const unbound = existing.filter((effect) => !claimed.has(effect)).map((effect) => effect.createDuplicate());
    timing.buildTree([
      ...(unbound.length ? [[null, ...unbound]] : []),
      ...scheduled.map((entry) => [entry.token, entry.effect]),
    ]);
    timing.setParent(slide);
    return { timing, scheduled, hasUnbound: unbound.length > 0, unboundCount: unbound.length };
  };
  const table = Reflect.get(context.AscCommon ?? {}, 'g_oTableId') as
    { m_aPairs?: Record<string, unknown> } | undefined;
  const previousIds = new Set(Object.keys(table?.m_aPairs ?? {}));
  let failure: unknown;
  const built = noHistory(
    () => {
      try {
        return build();
      } catch (error) {
        failure = error;
        return undefined;
      }
    },
    null,
    [],
    true,
  ) as ReturnType<typeof build> | undefined;
  const createdIds = Object.entries(table?.m_aPairs ?? {}).filter(([id]) => !previousIds.has(id));
  const cleanup = () => {
    const pairs = table?.m_aPairs;
    if (pairs) for (const [id, object] of createdIds) if (pairs[id] === object) delete pairs[id];
  };
  if (failure) {
    cleanup();
    throw failure;
  }
  if (!built || built.timing.getAllAnimEffects().length !== built.scheduled.length + built.unboundCount) {
    cleanup();
    throw new Error('Native animation schedule could not be built');
  }
  return {
    ...built,
    saved,
    slide,
    motions: built.scheduled.map(({ effect: _effect, token: _token, ...motion }) => motion),
    install() {
      if (slide.timing !== saved) throw new Error('Native timing changed before playback');
      slide.timing = built.timing;
    },
    restore() {
      if (slide.timing === built.timing) slide.timing = saved;
      cleanup();
    },
  };
}
export type SpeechNativeSchedule = NonNullable<ReturnType<typeof prepareSpeechNativeSchedule>>;
