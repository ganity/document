import { requireEditorContext } from './agent-plugin/editor-bridge';
import { nativeTeachingTarget } from './deckwise-native-teaching';
import { nativeTeachingMotionPreset, type AnimationCue } from './deckwise-teaching-animation-bindings';

type Effect = {
  cTn: { presetClass: number; presetID: number; presetSubtype?: number; setNodeType(value: number): void };
  getObjectId(): string;
  asc_getDuration(): number;
  asc_putDuration(value: number): void;
  asc_putDelay(value: number): void;
};
type Timing = {
  createDuplicate(): Timing;
  setParent(slide: Slide): void;
  getAllAnimEffects(): Effect[];
  createEffect(id: string, kind: number, preset: number, subtype: number, color: null): Effect;
  addEffectsToMainSequence(effects: Effect[]): void;
};
type Slide = { timing?: Timing | null; setTiming(value: Timing | null): void };

/** Stage a detached timing tree; publish it with media and metadata in one native Undo action. */
export function prepareNativeTeachingMotions(pageNumber: number, actions: Array<{ type: string; visualCues?: unknown }>) {
  const groups = actions.filter(action => action.type === 'speech').map(action =>
    (Array.isArray(action.visualCues) ? action.visualCues as AnimationCue[] : []).filter(cue => cue.effect?.type === 'element_animation' && !cue.nativeAnimation));
  if (!groups.some(group => group.length)) return;
  if (groups.flat().length > 200) throw new Error('Too many native speech animations');
  const context = requireEditorContext();
  const logic = (context.api.WordControl as { m_oLogicDocument: { Slides: Slide[] } }).m_oLogicDocument;
  const slide = logic.Slides[pageNumber - 1], previous = slide?.timing ?? null;
  const format = context.AscFormat as unknown as { CTiming: new () => Timing; ExecuteNoHistory<T>(fn: () => T, scope: null, args: [], keepIds: boolean): T };
  if (!slide || typeof slide.setTiming !== 'function' || !format?.CTiming || !format.ExecuteNoHistory)
    throw new Error('Native speech animation persistence is unavailable');
  const existing = previous?.getAllAnimEffects() ?? [], claimed = new Set<Effect>();
  const additions = groups.map(group => group.flatMap(cue => {
    const preset = nativeTeachingMotionPreset(cue.effect.effect);
    const shape = nativeTeachingTarget(cue.effect.elementId) as { Id?: string };
    const duration = cue.endMs - cue.atMs;
    if (!preset || !shape?.Id || !Number.isInteger(cue.atMs) || cue.atMs < 0 || !Number.isInteger(duration) || duration < 1 || duration > 5000)
      throw new Error('Invalid native speech animation target or duration');
    const match = existing.find(effect => !claimed.has(effect) && effect.getObjectId() === shape.Id &&
      effect.cTn.presetClass === preset.presetClass && effect.cTn.presetID === preset.presetId &&
      (effect.cTn.presetSubtype ?? 0) === preset.subtype && effect.asc_getDuration() === duration);
    if (match) { claimed.add(match); return []; }
    return [{ shapeId: shape.Id, preset, duration }];
  }));
  if (!additions.some(group => group.length)) return;
  if (actions.some(action => action.type === 'native_animation')) throw new Error('Rebind legacy native animation steps before adding speech animations');
  if (existing.length + additions.flat().length > 200) throw new Error('Native speech animation limit exceeded');
  const table = Reflect.get(context.AscCommon ?? {}, 'g_oTableId') as { m_aPairs?: Record<string, unknown> } | undefined;
  const previousIds = new Set(Object.keys(table?.m_aPairs ?? {}));
  let failure: unknown;
  const timing = format.ExecuteNoHistory(() => {
    try {
      const next = previous?.createDuplicate() ?? new format.CTiming();
      next.setParent(slide);
      for (const group of additions) {
        const effects = group.map(({ shapeId, preset, duration }, index) => {
          const effect = next.createEffect(shapeId, preset.presetClass, preset.presetId, preset.subtype, null);
          if (!effect) throw new Error('Native speech animation preset is unavailable');
          effect.asc_putDuration(duration); effect.asc_putDelay(0);
          effect.cTn.setNodeType(index ? 7 : 2);
          return effect;
        });
        if (effects.length) next.addEffectsToMainSequence(effects);
      }
      if (next.getAllAnimEffects().length !== existing.length + additions.flat().length)
        throw new Error('Native speech animation readback differs');
      return next;
    } catch (error) { failure = error; return undefined; }
  }, null, [], true);
  const createdIds = Object.entries(table?.m_aPairs ?? {}).filter(([id]) => !previousIds.has(id));
  const cleanup = () => {
    const pairs = table?.m_aPairs;
    if (pairs) for (const [id, object] of createdIds) if (pairs[id] === object) delete pairs[id];
  };
  if (failure || !timing) { cleanup(); throw failure ?? new Error('Native speech animation persistence failed'); }
  let applied = false;
  return {
    apply() { applied = true; slide.setTiming(timing); },
    restore() { try { if (applied) slide.setTiming(previous); } finally { cleanup(); } },
  };
}
