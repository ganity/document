import { hasNativeTeaching, nativeTeachingTarget } from './deckwise-native-teaching';
import { getEditorApi } from './agent-plugin/editor-bridge';
import { attentionTarget, type PreviewObject } from './deckwise-teaching-preview';
import { hasUnsavedChanges } from './unsaved-guard';

export type SpeechTrigger =
  | { version: 1; kind: 'start' | 'end' }
  | {
      version: 1;
      kind: 'text';
      text: string;
      start: number;
      end: number;
      phase: 'with' | 'after';
      audioAssetId: string;
      stale?: true;
    };
export type AnimationCue = {
  id: string;
  atMs: number;
  endMs: number;
  trigger?: SpeechTrigger;
  nativeAnimation?: string;
  effect: { type: 'element_animation'; elementId: string; effect: string };
};
type NativeEffect = {
  cTn: { presetClass: number; presetID: number; presetSubtype: number };
  getObjectId(): string;
  asc_getDuration(): number;
  asc_getDelay(): number;
  asc_getRepeatCount?(): number;
  asc_getRewind?(): boolean;
};
const presets: Array<[string, number, number, number]> = [
  ['fadeIn', 1, 10, 0],
  ['fadeOut', 2, 10, 0],
  ['pulse', 0, 26, 0],
  ['zoomIn', 1, 53, 16],
  ['slideInLeft', 1, 2, 8],
  ['slideInRight', 1, 2, 2],
  ['gentleRotate', 0, 32, 0],
];

export function nativeTeachingMotionPreset(effect: string) {
  const entry = presets.find(([name]) => name === effect);
  return entry && { presetClass: entry[1], presetId: entry[2], subtype: entry[3] };
}

// shortcut: multiple native effects on one object need a persistent effect ID before they can be linked safely.
/** A logical object with exactly one native effect is stable across SDK IDs and native reorder. */
export function listTeachingNativeAnimations(page: number, objects: PreviewObject[]) {
  const api = getEditorApi();
  const logic = Reflect.get(Reflect.get(api ?? {}, 'WordControl') ?? {}, 'm_oLogicDocument');
  const slide = logic?.Slides?.[page - 1];
  if (!slide || (hasUnsavedChanges() && !hasNativeTeaching())) throw new Error('Save the current slide before linking animations');
  const effects: NativeEffect[] = slide.timing?.getAllAnimEffects?.() ?? [];
  return effects.map((effect, index) => {
    const preset = presets.find(
      ([, kind, id, subtype]) =>
        kind === effect.cTn?.presetClass && id === effect.cTn.presetID && subtype === effect.cTn.presetSubtype,
    );
    const object = objects.find((object) => {
      try {
        return attentionTarget(slide.cSld.spTree, objects, object.objectId).Id === effect.getObjectId();
      } catch {
        return false;
      }
    });
    const unique = effects.filter((e) => e.getObjectId() === effect.getObjectId()).length === 1;
    const durationMs = effect.asc_getDuration();
    return {
      index,
      objectId: object?.objectId,
      effect: preset?.[0],
      durationMs,
      supported:
        !!object &&
        !!preset &&
        unique &&
        Number.isInteger(durationMs) &&
        durationMs > 0 &&
        durationMs <= 5000 &&
        effect.asc_getDelay() === 0 &&
        (effect.asc_getRepeatCount?.() ?? 1000) === 1000 &&
        !effect.asc_getRewind?.(),
      native: effect,
    };
  });
}

export function relocateAnimationTrigger(trigger: SpeechTrigger, text: string): SpeechTrigger {
  if (trigger.kind !== 'text') return trigger;
  const start = text.indexOf(trigger.text);
  return start >= 0 && text.indexOf(trigger.text, start + 1) < 0
    ? { ...trigger, start, end: start + trigger.text.length, stale: true }
    : { ...trigger, stale: true };
}

export function animationTriggerLabel(
  trigger: SpeechTrigger | undefined,
  labels: { start: string; end: string; with: string; after: string },
) {
  if (!trigger) return labels.start;
  if (trigger.kind === 'text') return `${trigger.text} · ${trigger.phase === 'with' ? labels.with : labels.after}`;
  return trigger.kind === 'start' ? labels.start : labels.end;
}
