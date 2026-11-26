import { requireEditorContext } from './agent-plugin/editor-bridge';
import { nativeTeachingTarget } from './deckwise-native-teaching';
import { nativeTeachingMotionPreset, type AnimationCue } from './deckwise-teaching-animation-bindings';

type Effect = {
  cTn: {
    presetClass: number;
    presetID: number;
    presetSubtype?: number;
    nodeType?: number;
    setNodeType(value: number): void;
  };
  getObjectId(): string;
  asc_getDuration(): number;
  asc_getDelay(): number;
  asc_putDuration(value: number): void;
  asc_putDelay(value: number): void;
};
type Timing = {
  createDuplicate(): Timing;
  setParent(slide: Slide): void;
  getAllAnimEffects(): Effect[];
  createEffect(id: string, kind: number, preset: number, subtype: number, color: null): Effect;
  addEffectsToMainSequence(effects: Effect[]): void;
  removeEffects(effects: Effect[]): void;
};
type Slide = { timing?: Timing | null; setTiming(value: Timing | null): void };

type MotionAction = { id?: string; type: string; visualCues?: unknown };
export type SpeechMotionOwners = { count: number; indices: Record<string, number> };

/** Stage a detached timing tree; publish it with media and metadata in one native Undo action. */
export function prepareNativeTeachingMotions(
  pageNumber: number,
  actions: MotionAction[],
  previousActions: MotionAction[],
  resolveTarget: (objectId: string) => unknown = nativeTeachingTarget,
  ownership?: { legacyGenerated?: boolean; previous?: SpeechMotionOwners; save(value: SpeechMotionOwners): void },
) {
  const motionGroups = (values: MotionAction[]) =>
    values
      .filter((action) => action.type === 'speech')
      .map((action) =>
        (Array.isArray(action.visualCues) ? (action.visualCues as AnimationCue[]) : []).filter(
          (cue) => cue.effect?.type === 'element_animation' && !cue.nativeAnimation,
        ),
      );
  const groups = motionGroups(actions),
    previousGroups = motionGroups(previousActions);
  const layout = (list: AnimationCue[][]) =>
    JSON.stringify(
      list
        .filter((group) => group.length)
        .map((group) => group.map((cue) => [cue.id, cue.effect.elementId, cue.effect.effect, cue.atMs, cue.endMs])),
    );
  if (layout(groups) === layout(previousGroups)) return;
  if (groups.flat().length > 200) throw new Error('Too many native speech animations');
  const context = requireEditorContext();
  const logic = (context.api.WordControl as { m_oLogicDocument: { Slides: Slide[] } }).m_oLogicDocument;
  const slide = logic.Slides[pageNumber - 1],
    previous = slide?.timing ?? null;
  const format = context.AscFormat as unknown as { CTiming: new () => Timing };
  if (!slide || typeof slide.setTiming !== 'function' || !format?.CTiming)
    throw new Error('Native speech animation persistence is unavailable');
  const existing = previous?.getAllAnimEffects() ?? [],
    claimed = new Set<Effect>(),
    owned = new Set<Effect>();
  const describe = (cue: AnimationCue) => {
    const preset = nativeTeachingMotionPreset(cue.effect.effect);
    const shape = resolveTarget(cue.effect.elementId) as { Id?: string };
    const duration = cue.endMs - cue.atMs;
    if (
      !preset ||
      !shape?.Id ||
      !Number.isInteger(cue.atMs) ||
      cue.atMs < 0 ||
      !Number.isInteger(duration) ||
      duration < 1 ||
      duration > 5000
    )
      throw new Error('Invalid native speech animation target or duration');
    return { shapeId: shape.Id, preset, duration, delay: cue.atMs, cueId: cue.id };
  };
  const matches = (effect: Effect, motion: ReturnType<typeof describe>) =>
    effect.getObjectId() === motion.shapeId &&
    effect.cTn.presetClass === motion.preset.presetClass &&
    effect.cTn.presetID === motion.preset.presetId &&
    (effect.cTn.presetSubtype ?? 0) === motion.preset.subtype &&
    effect.asc_getDuration() === motion.duration &&
    effect.asc_getDelay() === motion.delay;
  const before = ownership?.previous;
  if (before?.count === existing.length && before.indices && typeof before.indices === 'object') {
    for (const cue of previousGroups.flat()) {
      const index = Object.hasOwn(before.indices, cue.id) ? before.indices[cue.id] : -1;
      if (Number.isSafeInteger(index) && existing[index] && matches(existing[index], describe(cue)))
        owned.add(existing[index]);
    }
  } else if (!before && ownership && ownership.legacyGenerated) {
    // Legacy generated lessons have no owner ledger. Require the complete original click sequence.
    const old = previousActions
      .filter((a) => a.type === 'speech')
      .flatMap((a, i) =>
        previousGroups[i].map((cue, j) => ({ cue, generated: cue.id === `motion-${a.id}-${j}`, nodeType: j ? 7 : 2 })),
      );
    if (
      old.length === existing.length &&
      old.every(
        ({ cue, generated, nodeType }, i) =>
          generated && existing[i].cTn.nodeType === nodeType && matches(existing[i], describe(cue)),
      )
    )
      existing.forEach((effect) => owned.add(effect));
  }
  const additions = groups.map((group) =>
    group.flatMap((cue) => {
      const motion = describe(cue);
      const match = existing.find((effect) => !owned.has(effect) && !claimed.has(effect) && matches(effect, motion));
      if (match) {
        claimed.add(match);
        return [];
      }
      return [motion];
    }),
  );
  const removals = existing.flatMap((effect, index) => (owned.has(effect) ? [index] : []));
  if (!additions.some((group) => group.length) && !removals.length) {
    ownership?.save({ count: existing.length, indices: {} });
    return;
  }
  if (actions.some((action) => action.type === 'native_animation'))
    throw new Error('Rebind legacy native animation steps before changing speech animations');
  const expectedCount = existing.length - removals.length + additions.flat().length;
  if (expectedCount > 200) throw new Error('Native speech animation limit exceeded');
  let applied = false;
  return {
    apply() {
      // Register the timing tree in the same SDK history action as the assignment so peers can replay it.
      const timing = previous?.createDuplicate() ?? new format.CTiming();
      timing.setParent(slide);
      if (removals.length) {
        const copied = timing.getAllAnimEffects();
        timing.removeEffects(removals.map((index) => copied[index]));
      }
      const created = new Map<string, Effect>();
      const effects = additions.flatMap((group) =>
        group.map(({ shapeId, preset, duration, delay, cueId }, index) => {
          const effect = timing.createEffect(shapeId, preset.presetClass, preset.presetId, preset.subtype, null);
          if (!effect) throw new Error('Native speech animation preset is unavailable');
          effect.asc_putDuration(duration);
          effect.asc_putDelay(delay);
          created.set(cueId, effect);
          effect.cTn.setNodeType(index ? 7 : 2);
          return effect;
        }),
      );
      if (effects.length) timing.addEffectsToMainSequence(effects);
      const finalEffects = timing.getAllAnimEffects();
      const remaining = [...finalEffects];
      if (
        remaining.length !== expectedCount ||
        groups.flat().some((cue) => {
          const index = remaining.findIndex((effect) => matches(effect, describe(cue)));
          if (index < 0) return true;
          remaining.splice(index, 1);
          return false;
        })
      )
        throw new Error('Native speech animation readback differs');
      const indices: Record<string, number> = Object.create(null);
      for (const [cueId, effect] of created) {
        const index = finalEffects.findIndex((entry) => entry === effect || entry.cTn === effect.cTn);
        if (index < 0) throw new Error('Cannot identify the saved native speech animation');
        indices[cueId] = index;
      }
      applied = true;
      slide.setTiming(timing);
      ownership?.save({ count: finalEffects.length, indices });
    },
    restore() {
      if (applied) slide.setTiming(previous);
    },
  };
}
