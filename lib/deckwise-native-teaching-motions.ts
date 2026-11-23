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
  if (!groups.some((group) => group.length) && !previousGroups.some((group) => group.length)) return;
  if (groups.flat().length > 200) throw new Error('Too many native speech animations');
  const context = requireEditorContext();
  const logic = (context.api.WordControl as { m_oLogicDocument: { Slides: Slide[] } }).m_oLogicDocument;
  const slide = logic.Slides[pageNumber - 1],
    previous = slide?.timing ?? null;
  const format = context.AscFormat as unknown as {
    CTiming: new () => Timing;
    ExecuteNoHistory<T>(fn: () => T, scope: null, args: [], keepIds: boolean): T;
  };
  if (!slide || typeof slide.setTiming !== 'function' || !format?.CTiming || !format.ExecuteNoHistory)
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
  const layout = (list: AnimationCue[][]) =>
    JSON.stringify(
      list
        .filter((g) => g.length)
        .map((g) => g.map((c) => [c.id, c.effect.elementId, c.effect.effect, c.atMs, c.endMs])),
    );
  if (layout(groups) === layout(previousGroups)) {
    const remaining = [...existing];
    if (
      groups.flat().every((cue) => {
        const index = remaining.findIndex((effect) => matches(effect, describe(cue)));
        if (index < 0) return false;
        remaining.splice(index, 1);
        return true;
      })
    )
      return;
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
  const table = Reflect.get(context.AscCommon ?? {}, 'g_oTableId') as
    { m_aPairs?: Record<string, unknown> } | undefined;
  const previousIds = new Set(Object.keys(table?.m_aPairs ?? {}));
  const created = new Map<string, Effect>();
  let failure: unknown;
  const timing = format.ExecuteNoHistory(
    () => {
      try {
        const next = previous?.createDuplicate() ?? new format.CTiming();
        next.setParent(slide);
        if (removals.length) {
          const copied = next.getAllAnimEffects();
          next.removeEffects(removals.map((index) => copied[index]));
        }
        // Each SDK append rebuilds the tree, replacing the objects from previous appends.
        const effects = additions.flatMap(group =>
          group.map(({ shapeId, preset, duration, delay, cueId }, index) => {
            const effect = next.createEffect(shapeId, preset.presetClass, preset.presetId, preset.subtype, null);
            if (!effect) throw new Error('Native speech animation preset is unavailable');
            effect.asc_putDuration(duration);
            effect.asc_putDelay(delay);
            created.set(cueId, effect);
            effect.cTn.setNodeType(index ? 7 : 2);
            return effect;
          }),
        );
        if (effects.length) next.addEffectsToMainSequence(effects);
        const remaining = [...next.getAllAnimEffects()];
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
        return next;
      } catch (error) {
        failure = error;
        return undefined;
      }
    },
    null,
    [],
    true,
  );
  const createdIds = Object.entries(table?.m_aPairs ?? {}).filter(([id]) => !previousIds.has(id));
  const cleanup = () => {
    const pairs = table?.m_aPairs;
    if (pairs) for (const [id, object] of createdIds) if (pairs[id] === object) delete pairs[id];
  };
  if (failure || !timing) {
    cleanup();
    throw failure ?? new Error('Native speech animation persistence failed');
  }
  const finalEffects = timing.getAllAnimEffects();
  const indices: Record<string, number> = Object.create(null);
  for (const [cueId, effect] of created) {
    const index = finalEffects.findIndex((e) => e === effect || e.cTn === effect.cTn);
    if (index < 0) {
      cleanup();
      throw new Error('Cannot identify the saved native speech animation');
    }
    indices[cueId] = index;
  }
  ownership?.save({ count: finalEffects.length, indices });
  let applied = false;
  return {
    apply() {
      applied = true;
      slide.setTiming(timing);
    },
    restore() {
      try {
        if (applied) slide.setTiming(previous);
      } finally {
        cleanup();
      }
    },
  };
}
