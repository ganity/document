import { requireEditorContext } from './agent-plugin/editor-bridge';
import {
  attentionTarget,
  teachingShapes,
  type Shape,
  type PreviewObject,
  type Runtime,
} from './deckwise-teaching-preview';
import { type AnimationCue } from './deckwise-teaching-animation-bindings';

type Effect = {
  Id: string;
  cTn: { presetClass: number; presetID: number; presetSubtype: number; setNodeType(type: number): void };
  getObjectId(): string;
  asc_getDuration(): number;
  asc_getDelay(): number;
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
  if (!cues.length) return undefined;
  if (actions.some((action) => action.type === 'native_animation')) {
    if (cues.some(({ cue }) => cue.trigger || cue.nativeAnimation))
      throw new Error('Remove legacy page animation bindings before using speech animation links');
    return undefined;
  }
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
    const shapes = teachingShapes<Shape>(slide.cSld.spTree);
    const scheduled = cues
      .flatMap(({ actionId, cue }) => {
        const shape = attentionTarget(slide.cSld.spTree, objects, cue.effect.elementId);
        if (!shape.Id) throw new Error('Native animation target is missing');
        // A pictured teaching target may inherit its entrance from a saved PPT group.
        let candidates: Effect[] = [];
        for (
          let owner: Shape | undefined = shape;
          owner && shapes.includes(owner);
          owner = owner.group as Shape | undefined
        ) {
          const own = existing.filter((effect) => effect.getObjectId() === owner!.Id);
          candidates.push(...own);
          if (cue.nativeAnimation && own.length) break;
        }
        let effect: Effect;
        if (cue.nativeAnimation) {
          if (candidates.length !== 1 || claimed.has(candidates[0]))
            throw new Error('Linked native animation is missing or ambiguous');
          claimed.add(candidates[0]);
          effect = candidates[0].createDuplicate();
        } else {
          const preset = runtime.nativeMotionPreset(cue.effect.effect);
          if (!preset) throw new Error('Unsupported teaching animation');
          const mirror = candidates.find(
            (effect) =>
              !claimed.has(effect) &&
              effect.cTn.presetClass === preset.presetClass &&
              effect.cTn.presetID === preset.presetId &&
              (effect.cTn.presetSubtype ?? 0) === preset.subtype &&
              effect.asc_getDelay() === cue.atMs &&
              effect.asc_getDuration() === cue.endMs - cue.atMs,
          );
          if (!mirror) throw new Error('讲解关联的动画尚未写入 PPT 或与讲解时间不一致，请在当前页重新同步并保存该动画。');
          claimed.add(mirror);
          effect = mirror.createDuplicate();
        }
        effect.cTn.setNodeType(2);
        effect.asc_putDelay(0);
        return [
          {
            actionId,
            atMs: cue.atMs,
            endMs: cue.atMs + effect.asc_getDuration(),
            trigger: cue.trigger,
            effect,
            token: effect.Id,
          },
        ];
      })
      .map((entry, stepIndex) => ({ ...entry, stepIndex }));
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
