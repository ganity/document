import { captureDocumentToolTarget } from './document-tool-action';
import { requireEditorContext } from './editor-bridge';
import { requireSlideImageTransform, slideImageFingerprint } from './slide-image-transform';
import { readSlideShapeText, type SlideTextShape } from './slide-text-read';
import { getReadonlyMode } from '../onlyoffice/readonly';

interface Shape {
  Id: string;
  group?: unknown;
  isPlaceholder?(): boolean;
  getObjectType?(): number;
  getDocContent?(): unknown;
  spPr?: { geometry?: { preset?: string }; xfrm?: Record<string, number | boolean | null | undefined> };
}
interface Effect {
  cTn: { presetClass: number; presetID: number; presetSubtype: number; nodeType: number };
  getObjectId(): string;
  asc_getDuration(): number;
  asc_getDelay(): number;
}
interface Timing {
  bldLst?: unknown;
  getAllAnimEffects(): Effect[];
  getSelectedEffects(): Effect[];
  getEffectsSequences(): Array<Array<Effect | null>>;
}
interface AnimationApi {
  isDocumentLoadComplete?: boolean;
  isLoadFullApi?: boolean;
  isViewMode?: boolean;
  canEdit?(): boolean;
  isSlideShow?(): boolean;
  isLongAction?(): boolean;
  getCurrentPage?(): number;
  asc_IsStartedAnimationPreview?(): boolean;
  asc_AddAnimation?(
    presetClass: number,
    presetId: number,
    subtype: number,
    color: null,
    replace: false,
    preview: false,
  ): void;
  WordControl?: {
    m_oLogicDocument?: {
      Slides: Array<{ cSld: { spTree: Shape[] }; timing?: Timing; timingLock?: { Lock?: { Is_Locked?(): boolean } } }>;
      CanEdit?(): boolean;
      IsMasterMode?(): boolean;
      GetWidthMM(): number;
      GetHeightMM(): number;
      GetCurrentController?(): { selectedObjects?: Shape[]; getTargetDocContent?(): unknown };
    };
  };
}

// The SDK copies effects and renumbers cTn IDs when appending. Preserve their data, not runtime identity.
function effectFingerprint(effect: Effect): string {
  const seen = new Set<object>();
  let nodes = 0;
  function visit(value: unknown, depth: number): unknown {
    if (value === null || typeof value === 'string' || typeof value === 'boolean') return value;
    if (typeof value === 'number' && Number.isFinite(value)) return value;
    if (value === undefined) return null;
    if (!value || typeof value !== 'object' || depth > 32 || ++nodes > 2048 || seen.has(value))
      throw new Error('Unsupported native animation data');
    seen.add(value);
    const result = Array.isArray(value)
      ? value.map((item) => visit(item, depth + 1))
      : Object.fromEntries(
          Object.keys(value)
            .sort()
            .filter(
              (key) =>
                !['Id', 'id', 'parent', 'selected', 'state', 'simpleDurationIdx', 'originalNode'].includes(key) &&
                typeof Reflect.get(value, key) !== 'function',
            )
            .map((key) => {
              if (key === 'tn' && Reflect.get(value, key) != null)
                throw new Error('Animation node references are unsupported');
              return [key, visit(Reflect.get(value, key), depth + 1)];
            }),
        );
    seen.delete(value);
    return result;
  }
  return JSON.stringify(visit(effect, 0));
}

function animations(slide: { cSld: { spTree: Shape[] }; timing?: Timing }) {
  const timing = slide.timing;
  if (!timing) return [];
  if (
    timing.bldLst != null ||
    typeof timing.getAllAnimEffects !== 'function' ||
    typeof timing.getSelectedEffects !== 'function' ||
    typeof timing.getEffectsSequences !== 'function'
  )
    throw new Error('Animation builds or unknown timing are unsupported');
  const effects = timing.getAllAnimEffects(),
    sequences = timing.getEffectsSequences();
  if (
    !Array.isArray(effects) ||
    effects.length > 128 ||
    !Array.isArray(sequences) ||
    (effects.length
      ? sequences.length !== 1 ||
        sequences[0]?.[0] !== null ||
        sequences[0].length !== effects.length + 1 ||
        effects.some((effect, i) => sequences[0][i + 1] !== effect)
      : sequences.length !== 0)
  )
    throw new Error('Only a simple native main animation sequence is supported');
  return effects.map((effect) => {
    if (
      !effect ||
      effect.cTn?.presetClass !== 1 ||
      ![1, 10].includes(effect.cTn.presetID) ||
      effect.cTn.presetSubtype !== 0 ||
      effect.cTn.nodeType !== 2 ||
      typeof effect.getObjectId !== 'function' ||
      !slide.cSld.spTree.some((shape) => shape.Id === effect.getObjectId()) ||
      typeof effect.asc_getDuration !== 'function' ||
      typeof effect.asc_getDelay !== 'function' ||
      effect.asc_getDuration() !== (effect.cTn.presetID === 1 ? 1 : 500) ||
      effect.asc_getDelay() !== 0
    )
      throw new Error('Existing animation is outside the verified on-click appear/fade scope');
    return { effect, fingerprint: effectFingerprint(effect) };
  });
}

function current() {
  const context = requireEditorContext(),
    api = context.api as unknown as AnimationApi;
  const logic = api.WordControl?.m_oLogicDocument,
    page = api.getCurrentPage?.();
  const slide = page === undefined ? undefined : logic?.Slides[page];
  const controller = logic?.GetCurrentController?.(),
    shape = controller?.selectedObjects?.[0];
  const format = context.AscFormat,
    changeType = Reflect.get(context.AscCommon ?? {}, 'changestype_Timing');
  if (
    !format ||
    format.PRESET_CLASS_ENTR !== 1 ||
    format.ENTRANCE_APPEAR !== 1 ||
    format.ENTRANCE_FADE !== 10 ||
    format.NODE_TYPE_CLICKEFFECT !== 2 ||
    !Number.isSafeInteger(changeType) ||
    getReadonlyMode() ||
    api.isViewMode ||
    api.canEdit?.() !== true ||
    logic?.CanEdit?.() !== true ||
    !api.isDocumentLoadComplete ||
    !api.isLoadFullApi ||
    api.isSlideShow?.() ||
    api.isLongAction?.() ||
    api.asc_IsStartedAnimationPreview?.() !== false ||
    logic.IsMasterMode?.() ||
    context.AscCommon?.CollaborativeEditing?.Get_GlobalLock() ||
    typeof api.asc_AddAnimation !== 'function' ||
    !slide ||
    // The SDK's IsSelectionLocked requests a lock. Context refresh must only read it.
    slide.timingLock?.Lock?.Is_Locked?.() !== false ||
    !shape ||
    controller?.selectedObjects?.length !== 1 ||
    !slide.cSld.spTree.includes(shape) ||
    !shape.Id ||
    shape.group ||
    shape.isPlaceholder?.() !== false
  )
    throw new Error('Animation requires one current writable ordinary object');
  const imageType = Reflect.get(context.AscDFH ?? {}, 'historyitem_type_ImageShape');
  const shapeType = Reflect.get(context.AscDFH ?? {}, 'historyitem_type_Shape');
  if (Number.isSafeInteger(imageType) && shape.getObjectType?.() === imageType) {
    if (requireSlideImageTransform().image !== shape) throw new Error('Selected image changed');
  } else if (Number.isSafeInteger(shapeType) && shape.getObjectType?.() === shapeType) {
    const content = controller.getTargetDocContent?.(),
      shapeContent = shape.getDocContent?.();
    if (
      (content && content !== shapeContent) ||
      (shapeContent && readSlideShapeText(shape as SlideTextShape) === undefined)
    )
      throw new Error('Text selection changed');
    const t = shape.spPr?.xfrm,
      width = logic.GetWidthMM(),
      height = logic.GetHeightMM();
    if (
      !t ||
      shape.spPr?.geometry?.preset !== 'rect' ||
      (t.rot != null && t.rot !== 0) ||
      t.flipH ||
      t.flipV ||
      ![t.offX, t.offY, t.extX, t.extY, width, height].every((v) => typeof v === 'number' && Number.isFinite(v)) ||
      Number(t.offX) < 0 ||
      Number(t.offY) < 0 ||
      Number(t.extX) <= 0 ||
      Number(t.extY) <= 0 ||
      Number(t.offX) + Number(t.extX) > width ||
      Number(t.offY) + Number(t.extY) > height
    )
      throw new Error('Only ordinary on-page unrotated rectangular shapes and text boxes are supported');
  } else throw new Error('This object type does not support verified animation');
  const effects = animations(slide);
  if (slide.timing?.getSelectedEffects().some((effect) => effect.getObjectId() !== shape.Id))
    throw new Error('Animation-pane selection differs from the selected object');
  return { api, logic, slide, shape, page: page! + 1, effects };
}

export function canAddSlideObjectAnimation(): boolean {
  try {
    const target = captureDocumentToolTarget(),
      before = current();
    before.slide.cSld.spTree.forEach((shape) => slideImageFingerprint(shape));
    return target.context.kind === 'slide' && target.isCurrent(true);
  } catch {
    return false;
  }
}

/** Append one fixed native entrance without replacing existing effects or starting a preview. */
export function addSlideObjectAnimation(input: unknown, signal?: AbortSignal) {
  if (
    !input ||
    typeof input !== 'object' ||
    Array.isArray(input) ||
    Object.keys(input).length !== 1 ||
    typeof Reflect.get(input, 'effect') !== 'string' ||
    !['appear', 'fade'].includes(Reflect.get(input, 'effect'))
  )
    throw new Error('Animation requires only effect: appear or fade');
  const effect = Reflect.get(input, 'effect') as 'appear' | 'fade';
  if (!Object.hasOwn(input, 'effect')) throw new Error('Animation effect is required');
  signal?.throwIfAborted();
  const before = current(),
    target = captureDocumentToolTarget();
  if (!target.isCurrent(true)) throw new Error('Read fresh editor context');
  const peers = before.slide.cSld.spTree.map((shape) => ({ shape, value: slideImageFingerprint(shape) }));
  try {
    before.api.asc_AddAnimation!(1, effect === 'appear' ? 1 : 10, 0, null, false, false);
    signal?.throwIfAborted();
    const after = current(),
      added = after.effects.at(-1)?.effect;
    if (
      after.api !== before.api ||
      after.logic !== before.logic ||
      after.slide !== before.slide ||
      after.shape !== before.shape ||
      after.page !== before.page ||
      after.effects.length !== before.effects.length + 1 ||
      before.effects.some((item, index) => after.effects[index]?.fingerprint !== item.fingerprint) ||
      !added ||
      added.getObjectId() !== before.shape.Id ||
      added.cTn.presetID !== (effect === 'appear' ? 1 : 10) ||
      after.slide.cSld.spTree.length !== peers.length ||
      peers.some(
        ({ shape, value }, index) => after.slide.cSld.spTree[index] !== shape || slideImageFingerprint(shape) !== value,
      )
    )
      throw new Error('Native animation readback mismatch');
    return { effect, page: before.page, verified: true as const };
  } catch (cause) {
    throw new Error('Object animation unverified. Inspect and Undo if needed; do not retry automatically.', { cause });
  }
}
