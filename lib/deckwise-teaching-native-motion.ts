import { hasNativeTeaching, nativeTeachingTarget } from './deckwise-native-teaching';
import { requireEditorContext } from './agent-plugin/editor-bridge';
import { slideImageFingerprint } from './agent-plugin/slide-image-transform';
import { getReadonlyMode } from './onlyoffice/readonly';
import { hasUnsavedChanges } from './unsaved-guard';
import { attentionTarget, type PreviewObject, type Runtime } from './deckwise-teaching-preview';
import { inspectTeachingAnimations } from './deckwise-teaching-slideshow';

type Cue = { id: string; atMs: number; endMs: number; effect: { type: string; elementId?: string; effect?: string } };
type Speech = { id: string; type: string; visualCues?: Cue[] };
type Effect = {
  cTn: { presetClass: number; presetID: number; presetSubtype: number; nodeType: number };
  getObjectId(): string;
  asc_getDuration(): number;
  asc_getDelay(): number;
  createDuplicate(): Effect;
  merge(effect: Effect): void;
  asc_putDuration(duration: number): void;
  asc_putStartType(type: number): void;
};
export function generatedMotionBindings(
  page: number,
  actions: Speech[],
  objects: PreviewObject[],
  runtime: Pick<Runtime, 'nativeMotionPreset'>,
) {
  const motions = actions.flatMap((action) =>
    action.type === 'speech'
      ? (action.visualCues ?? [])
          .filter((c) => c.effect?.type === 'element_animation')
          .map((cue) => ({ actionId: action.id, cue }))
      : [],
  );
  if (!motions.length || actions.some((a) => a.type === 'native_animation')) return undefined;
  const inspected = inspectTeachingAnimations(page, true);
  if (motions.length > 200) throw new Error('Too many generated native motions');
  let stepIndex = -1;
  let previousAction: string | undefined;
  const bindings = motions.map(({ actionId, cue }) => {
    const preset = runtime.nativeMotionPreset(cue.effect.effect ?? '');
    if (
      !preset ||
      cue.atMs !== 0 ||
      !Number.isInteger(cue.endMs) ||
      cue.endMs < 1 ||
      cue.endMs > 5000 ||
      !cue.effect.elementId
    )
      throw new Error('Generated native motion requires a bounded speech-start cue');
    const nodeType = previousAction === actionId ? 7 : 2;
    if (nodeType === 2) stepIndex++;
    previousAction = actionId;
    const shape = attentionTarget(inspected.slide.cSld.spTree, objects, cue.effect.elementId);
    if ((!hasNativeTeaching() && shape.group) || typeof shape.Id !== 'string' || !shape.Id) throw new Error('Unsupported native motion target');
    return { actionId, cueId: cue.id, stepIndex, nodeType, shape, preset, duration: cue.endMs };
  });
  const matches = (effect: Effect, binding: (typeof bindings)[number]) =>
    !!effect &&
    effect.getObjectId() === binding.shape.Id &&
    effect.cTn.presetClass === binding.preset.presetClass &&
    effect.cTn.presetID === binding.preset.presetId &&
    effect.cTn.presetSubtype === binding.preset.subtype &&
    effect.cTn.nodeType === binding.nodeType &&
    effect.asc_getDelay() === 0 &&
    effect.asc_getDuration() === binding.duration;
  return { ...inspected, bindings, matches };
}

/** Native edits own effect properties; speech bindings use object identity and saved step order. */
export function nativeSpeechMotionSteps(
  page: number,
  actions: Speech[],
  objects: PreviewObject[],
  runtime: Pick<Runtime, 'nativeMotionPreset'>,
) {
  const planned = generatedMotionBindings(page, actions, objects, runtime);
  if (!planned?.steps.length) return undefined;
  const result: Array<{ actionId: string; stepIndex: number }> = [];
  let previous = -1;
  for (const [stepIndex, effects] of planned.steps.entries()) {
    const targets = effects.map((effect: Effect) => effect.getObjectId());
    const ownsTargets = (action: Speech) =>
      targets.every((id: string) =>
        planned.bindings.some((binding) => binding.actionId === action.id && binding.shape.Id === id),
      );
    const next = actions.findIndex((action, index) => index > previous && ownsTargets(action));
    const actionIndex = next >= 0 ? next : previous >= 0 && ownsTargets(actions[previous]) ? previous : -1;
    // Unbound or reordered manual steps keep the native/manual presentation path.
    if (actionIndex < 0) return undefined;
    previous = actionIndex;
    result.push({ actionId: actions[actionIndex].id, stepIndex });
  }
  return result;
}

/** Append to an empty sequence only. Existing native effects are never overwritten or retried. */
export function integrateGeneratedNativeMotions(
  page: number,
  actions: Speech[],
  objects: PreviewObject[],
  runtime: Pick<Runtime, 'nativeMotionPreset'>,
): boolean {
  const planned = generatedMotionBindings(page, actions, objects, runtime);
  if (!planned || !planned.bindings.length || planned.effects.length) return false;
  const context = requireEditorContext();
  const api = planned.api,
    logic = api.WordControl.m_oLogicDocument,
    controller = logic.GetCurrentController?.();
  const changeType = Reflect.get(context.AscCommon ?? {}, 'changestype_Timing');
  if (
    getReadonlyMode() ||
    (hasUnsavedChanges() && !hasNativeTeaching()) ||
    !api.isDocumentLoadComplete ||
    !api.isLoadFullApi ||
    api.asc_IsStartedAnimationPreview?.() !== false ||
    api.canEdit?.() !== true ||
    logic.CanEdit?.() !== true ||
    api.isLongAction?.() ||
    api.isViewMode ||
    logic.IsMasterMode?.() ||
    context.AscCommon?.CollaborativeEditing?.Get_GlobalLock() ||
    !Number.isSafeInteger(changeType) ||
    typeof api.asc_AddAnimation !== 'function' ||
    typeof api.asc_SetAnimationProperties !== 'function' ||
    typeof controller?.resetSelection !== 'function' ||
    typeof controller?.selectObject !== 'function'
  )
    throw new Error('Native teaching animation requires a current writable slide');
  const selected = [...(controller.selectedObjects ?? [])];
  const peers: Array<{ shape: object; fingerprint: ReturnType<typeof slideImageFingerprint> }> =
    planned.slide.cSld.spTree.map((shape: object) => ({ shape, fingerprint: slideImageFingerprint(shape) }));
  const effects = () => (planned.slide.timing?.getAllAnimEffects() ?? []) as Effect[];
  try {
    for (const [index, binding] of planned.bindings.entries()) {
      if (api.getCurrentPage() !== page - 1 || logic.Slides[page - 1] !== planned.slide)
        throw new Error('Teaching animation page changed');
      controller.resetSelection();
      controller.selectObject(binding.shape, page - 1);
      if (controller.selectedObjects?.length !== 1 || controller.selectedObjects[0] !== binding.shape)
        throw new Error('Native animation target selection failed');
      if (logic.IsSelectionLocked?.(changeType) !== false) throw new Error('Teaching animation target is locked');
      planned.slide.timing?.resetSelection?.();
      api.asc_AddAnimation(
        binding.preset.presetClass,
        binding.preset.presetId,
        binding.preset.subtype,
        null,
        false,
        false,
      );
      if (effects().length !== index + 1)
        throw new Error(
          `Native animation append failed at ${index}: selected=${JSON.stringify(controller.selectedObjects?.map((s: { Id?: string }) => s.Id))}, timingSelection=${planned.slide.timing?.getSelectedEffects?.().length}`,
        );
      const added = effects().at(-1);
      if (!added || typeof added.createDuplicate !== 'function') throw new Error('Native animation was not created');
      planned.slide.timing?.resetSelection?.();
      if (typeof Reflect.get(added, 'select') === 'function') Reflect.get(added, 'select').call(added);
      const properties = added.createDuplicate();
      properties.merge(added);
      properties.asc_putDuration(binding.duration);
      api.asc_SetAnimationProperties(properties);
      if (binding.nodeType !== 2) {
        const start = effects().at(-1)!.createDuplicate();
        start.merge(effects().at(-1)!);
        start.asc_putStartType(binding.nodeType);
        api.asc_SetAnimationProperties(start);
      }
      const current = effects();
      if (current.length !== index + 1 || current.some((effect, i) => !planned.matches(effect, planned.bindings[i])))
        throw new Error(`Native animation readback differs at step ${index}`);
    }
    if (
      planned.slide.cSld.spTree.length !== peers.length ||
      peers.some(
        ({ shape, fingerprint }, i) =>
          planned.slide.cSld.spTree[i] !== shape || slideImageFingerprint(shape) !== fingerprint,
      )
    )
      throw new Error('Native motion changed slide content');
    return true;
  } catch (cause) {
    throw new Error(
      'Native teaching animation could not be verified. Inspect the animation pane and Undo before retrying.',
      { cause },
    );
  } finally {
    if (api.getCurrentPage() === page - 1 && logic.Slides[page - 1] === planned.slide) {
      controller.resetSelection();
      for (const shape of selected) controller.selectObject(shape, page - 1);
      logic.Document_UpdateSelectionState?.();
    }
  }
}
