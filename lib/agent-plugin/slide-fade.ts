import { captureDocumentToolTarget } from './document-tool-action';
import { requireEditorContext, type EditorApi } from './editor-bridge';
import { readSlideShapeText, type SlideTextShape } from './slide-text-read';
import { getReadonlyMode } from '../onlyoffice/readonly';

interface Transition {
  TransitionType?: number;
  TransitionOption?: number;
  TransitionDuration?: number;
  SlideAdvanceOnMouseClick?: boolean;
  SlideAdvanceAfter?: boolean;
  SlideAdvanceDuration?: number;
  ShowLoop?: boolean;
}
const transitionKeys = [
  'TransitionType',
  'TransitionOption',
  'TransitionDuration',
  'SlideAdvanceOnMouseClick',
  'SlideAdvanceAfter',
  'SlideAdvanceDuration',
  'ShowLoop',
] as const;
interface Shape extends SlideTextShape {
  Id?: string;
  spPr?: { xfrm?: Record<string, unknown> };
  blipFill?: unknown;
}
interface Slide {
  cSld: { spTree: Shape[] };
  transition: Transition;
  timing?: unknown;
  notesShape?: Shape;
  deleteLock?: { Lock?: { Type?: number } };
  transitionLock?: { Lock?: { Type?: number } };
}
interface FadeApi extends EditorApi {
  canEdit?(): boolean;
  isSlideShow?(): boolean;
  isLongAction?(): boolean;
  getCurrentPage?(): number;
  ApplySlideTransition(value: Transition): void;
  WordControl?: {
    m_oLogicDocument?: {
      Slides: Slide[];
      FocusOnNotes?: boolean;
      CanEdit?(): boolean;
      IsMasterMode?(): boolean;
      GetSelectedSlides?(): number[];
      isLoopShowMode(): boolean;
      GetCurrentController?(): {
        selectedObjects: unknown[];
        getTargetDocContent?(): { GetSelectionState?(): unknown };
      };
    };
  };
}
const transitionValue = (value: Transition) => JSON.stringify(transitionKeys.map((key) => value[key]));
function shapeValue(shape: Shape) {
  return JSON.stringify([
    shape.Id,
    readSlideShapeText(shape),
    shape.spPr?.xfrm && ['offX', 'offY', 'extX', 'extY', 'rot', 'flipH', 'flipV'].map((key) => shape.spPr!.xfrm![key]),
    shape.blipFill,
  ]);
}
function current() {
  const context = requireEditorContext(),
    api = context.api as FadeApi;
  const common = context.AscCommon as typeof context.AscCommon & {
    c_oAscLockTypes?: { kLockTypeNone: number; kLockTypeMine: number };
  };
  const logic = api.WordControl?.m_oLogicDocument,
    page = api.getCurrentPage?.();
  const slide = Number.isSafeInteger(page) ? logic?.Slides?.[page!] : undefined;
  const controller = logic?.GetCurrentController?.(),
    selected = logic?.GetSelectedSlides?.();
  const types = context.Asc.c_oAscSlideTransitionTypes as { None?: number; Fade?: number } | undefined;
  const params = context.Asc.c_oAscSlideTransitionParams as { Fade_Smoothly?: number } | undefined;
  const Constructor = context.Asc.CAscSlideTransition as
    (new () => Transition & { get_ShowLoop(): boolean | undefined }) | undefined;
  const lockTypes = common?.c_oAscLockTypes;
  if (
    getReadonlyMode() ||
    api.isViewMode ||
    api.canEdit?.() !== true ||
    logic?.CanEdit?.() !== true ||
    !api.isDocumentLoadComplete ||
    !api.isLoadFullApi ||
    api.isLongAction?.() ||
    api.isSlideShow?.() ||
    logic.IsMasterMode?.() !== false ||
    logic.FocusOnNotes ||
    common?.CollaborativeEditing?.Get_GlobalLock() ||
    !slide ||
    !controller ||
    !Array.isArray(controller.selectedObjects) ||
    selected?.length !== 1 ||
    selected[0] !== page ||
    !lockTypes ||
    ![slide.deleteLock, slide.transitionLock].every(
      (lock) =>
        Number.isInteger(lock?.Lock?.Type) &&
        [lockTypes.kLockTypeNone, lockTypes.kLockTypeMine].includes(lock!.Lock!.Type!),
    ) ||
    typeof api.ApplySlideTransition !== 'function' ||
    typeof logic.isLoopShowMode !== 'function' ||
    typeof Constructor !== 'function' ||
    ![types?.None, types?.Fade, params?.Fade_Smoothly].every(Number.isSafeInteger) ||
    types!.None === types!.Fade ||
    typeof logic.isLoopShowMode() !== 'boolean' ||
    !slide.transition ||
    Object.keys(slide.transition).some((key) => !transitionKeys.includes(key as (typeof transitionKeys)[number])) ||
    ![types!.None, types!.Fade].includes(slide.transition.TransitionType) ||
    (slide.transition.TransitionType === types!.Fade && slide.transition.TransitionOption !== params!.Fade_Smoothly) ||
    slide.transition.SlideAdvanceAfter !== false ||
    slide.transition.SlideAdvanceOnMouseClick !== true ||
    !Number.isFinite(slide.transition.TransitionDuration) ||
    slide.transition.TransitionDuration! < 0
  )
    throw new Error('Fade requires one writable current slide with a simple click-advanced transition and native APIs');
  return {
    api,
    logic,
    slide,
    controller,
    page: page!,
    Constructor,
    fade: types!.Fade!,
    smooth: params!.Fade_Smoothly!,
  };
}

export function canSetSlideFade(): boolean {
  try {
    const target = captureDocumentToolTarget();
    return target.context.kind === 'slide' && target.isCurrent(true) && !!current();
  } catch {
    return false;
  }
}

/** Host-only: first validate/review the latest clean PPTX transition; SDK state cannot prove absence of sound/extensions. */
export function prepareSlideFade(signal?: AbortSignal) {
  signal?.throwIfAborted();
  const before = current(),
    target = captureDocumentToolTarget();
  if (target.context.kind !== 'slide' || !target.isCurrent(true)) throw new Error('Read fresh editor context');
  if (
    before.slide.transition.TransitionType === before.fade &&
    before.slide.transition.TransitionOption === before.smooth &&
    before.slide.transition.TransitionDuration === 750
  )
    throw new Error('Current slide already has the requested fade');
  const loop = before.logic.isLoopShowMode(),
    selection = before.controller.selectedObjects.slice();
  const content = before.controller.getTargetDocContent?.();
  const position = JSON.stringify(content?.GetSelectionState?.());
  const slides = before.logic.Slides.map((slide) => ({
    slide,
    transition: transitionValue(slide.transition),
    timing: slide.timing,
    notes: slide.notesShape,
    notesText: slide.notesShape && readSlideShapeText(slide.notesShape),
    shapes: slide.cSld.spTree.map((shape) => ({ shape, value: shapeValue(shape) })),
  }));
  const unchanged = (after: boolean) =>
    before.logic.Slides.length === slides.length &&
    slides.every(
      (saved, index) =>
        before.logic.Slides[index] === saved.slide &&
        saved.slide.timing === saved.timing &&
        saved.slide.notesShape === saved.notes &&
        (saved.notes && readSlideShapeText(saved.notes)) === saved.notesText &&
        ((after && saved.slide === before.slide) || transitionValue(saved.slide.transition) === saved.transition) &&
        saved.slide.cSld.spTree.length === saved.shapes.length &&
        saved.shapes.every(
          ({ shape, value }, i) => saved.slide.cSld.spTree[i] === shape && shapeValue(shape) === value,
        ),
    );
  const assertCurrent = (after = false) => {
    signal?.throwIfAborted();
    const now = current();
    if (
      (!after && !target.isCurrent(true)) ||
      now.api !== before.api ||
      now.logic !== before.logic ||
      now.slide !== before.slide ||
      now.controller !== before.controller ||
      now.page !== before.page ||
      now.logic.isLoopShowMode() !== loop ||
      now.controller.selectedObjects.length !== selection.length ||
      !selection.every((shape, index) => now.controller.selectedObjects[index] === shape) ||
      now.controller.getTargetDocContent?.() !== content ||
      JSON.stringify(content?.GetSelectionState?.()) !== position ||
      !unchanged(after)
    )
      throw new Error('Fade context or slide contents changed');
  };
  let consumed = false;
  return Object.freeze({
    page: before.page + 1,
    apply: () => {
      if (consumed) throw new Error('Fade was already consumed');
      consumed = true;
      assertCurrent();
      const value = new before.Constructor();
      value.TransitionType = before.fade;
      value.TransitionOption = before.smooth;
      value.TransitionDuration = 750;
      // Omitting ShowLoop preserves the presentation-wide loop setting in ApplySlideTransition.
      if (
        transitionKeys.slice(3).some((key) => value[key] !== undefined) ||
        Object.keys(value).some((key) => !transitionKeys.includes(key as (typeof transitionKeys)[number])) ||
        typeof value.get_ShowLoop !== 'function' ||
        value.get_ShowLoop() !== undefined
      )
        throw new Error('Unsupported native transition constructor');
      try {
        before.api.ApplySlideTransition(value);
        assertCurrent(true);
        const expected = JSON.parse(slides[before.page].transition);
        expected[0] = before.fade;
        expected[1] = before.smooth;
        expected[2] = 750;
        if (transitionValue(before.slide.transition) !== JSON.stringify(expected))
          throw new Error('Native fade readback differs');
      } catch (error) {
        throw new Error(
          'Fade could not be verified. Inspect the document and use Undo if needed; do not retry automatically.',
          { cause: error },
        );
      }
      return {
        page: before.page + 1,
        transition: 'fade' as const,
        durationMs: 750,
        verified: true as const,
        saved: false as const,
      };
    },
  });
}
