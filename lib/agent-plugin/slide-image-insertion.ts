import { captureDocumentToolTarget } from './document-tool-action';
import { requireEditorContext } from './editor-bridge';
import { validateSlidePng, preloadSlidePng, type PngApi, type PngCommon } from './slide-image-png';
import { readSlideShapeText, type SlideTextShape } from './slide-text-read';
import { getReadonlyMode } from '../onlyoffice/readonly';

export interface SlideImageBox {
  xMm: number;
  yMm: number;
  widthMm: number;
  heightMm: number;
}
interface ImageShape {
  Id: string;
  setDescription?(description: string): void;
  getCNvProps?(): { descr?: string };
  getDocContent?: SlideTextShape['getDocContent'];
  group?: unknown;
  isPlaceholder?(): boolean;
  getObjectType?(): number;
  blipFill?: { RasterImageId?: string; srcRect?: unknown; tile?: unknown; transparent?: unknown };
  spPr?: { geometry?: { preset?: string }; xfrm?: Record<string, unknown> };
  setParent(slide: unknown): void;
  addToDrawingObjects(): void;
}
interface ImageController {
  selectedObjects: ImageShape[];
  getTargetDocContent?(): unknown;
  resetSelection(): void;
  createImage(url: string, x: number, y: number, width: number, height: number): ImageShape;
  selectObject(image: ImageShape, page: number): void;
}
interface ImageSlide {
  cSld: { spTree: ImageShape[] };
  deleteLock?: { Lock?: { Type?: number } };
}
interface ImageApi extends PngApi {
  canEdit?(): boolean;
  isSlideShow?(): boolean;
  isLongAction?(): boolean;
  getCurrentPage?(): number;
  WordControl?: {
    m_oLogicDocument?: {
      Slides: ImageSlide[];
      FocusOnNotes?: boolean;
      CanEdit?(): boolean;
      IsMasterMode?(): boolean;
      GetWidthMM(): number;
      GetHeightMM(): number;
      GetCurrentController?(): ImageController;
      StartAction(description: number): void;
      FinalizeAction(): void;
      Document_UpdateInterfaceState(): void;
      CheckEmptyPlaceholderNotes(): void;
    };
  };
}
interface ImageCommon extends PngCommon {
  CollaborativeEditing?: { Get_GlobalLock(): boolean };
  c_oAscLockTypes?: { kLockTypeNone: number; kLockTypeMine: number };
}
const geometryKeys = ['offX', 'offY', 'extX', 'extY', 'rot', 'flipH', 'flipV'] as const;
function fingerprint(shape: ImageShape) {
  return JSON.stringify({
    id: shape.Id,
    geometry: geometryKeys.map((key) => shape.spPr?.xfrm?.[key] ?? null),
    text: readSlideShapeText(shape),
    image: shape.blipFill,
  });
}
function current(emptySelection: boolean) {
  const context = requireEditorContext();
  const api = context.api as ImageApi;
  const common = context.AscCommon as typeof context.AscCommon & Partial<ImageCommon>;
  const logic = api.WordControl?.m_oLogicDocument;
  const page = api.getCurrentPage?.();
  const slide = Number.isSafeInteger(page) ? logic?.Slides?.[page!] : undefined;
  const controller = logic?.GetCurrentController?.();
  const description = Reflect.get(context.AscDFH ?? {}, 'historydescription_Presentation_AddFlowImage');
  const imageType = Reflect.get(context.AscDFH ?? {}, 'historyitem_type_ImageShape');
  const lock = slide?.deleteLock?.Lock?.Type;
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
    (emptySelection && controller.selectedObjects.length !== 0) ||
    controller.getTargetDocContent?.() ||
    !common.c_oAscLockTypes ||
    !Number.isInteger(lock) ||
    ![common.c_oAscLockTypes.kLockTypeNone, common.c_oAscLockTypes.kLockTypeMine].includes(lock!) ||
    !Number.isSafeInteger(description) ||
    !Number.isSafeInteger(imageType) ||
    !['StartAction', 'FinalizeAction', 'Document_UpdateInterfaceState', 'CheckEmptyPlaceholderNotes'].every(
      (key) => typeof Reflect.get(logic, key) === 'function',
    ) ||
    !['resetSelection', 'createImage', 'selectObject'].every(
      (key) => typeof Reflect.get(controller, key) === 'function',
    ) ||
    typeof api.ImageLoader?.LoadImage !== 'function' ||
    typeof api.ImageLoader.LoadImagesWithCallback !== 'function' ||
    typeof common.sendImgUrls !== 'function' ||
    typeof common.getFullImageSrc2 !== 'function' ||
    typeof common.g_oDocumentUrls?.getImageLocal !== 'function' ||
    typeof common.g_oDocumentUrls.imagePath2Local !== 'function'
  )
    throw new Error(
      emptySelection
        ? 'Image insertion requires a writable slide with an empty canvas selection and native image APIs'
        : 'Image insertion requires a writable slide outside text editing and native image APIs',
    );
  const width = logic.GetWidthMM(),
    height = logic.GetHeightMM();
  if (![width, height].every((value) => Number.isFinite(value) && value > 0))
    throw new Error('Invalid slide dimensions');
  return {
    api,
    common: common as ImageCommon,
    logic,
    slide,
    controller,
    page: page! + 1,
    width,
    height,
    description: description as number,
    imageType: imageType as number,
  };
}
export function canInsertSlideImage(): boolean {
  try {
    const target = captureDocumentToolTarget();
    return target.context.kind === 'slide' && target.isCurrent(true) && !!current(true);
  } catch {
    return false;
  }
}
/** Host-only: caller must authorize/review these PNG bytes and the explicit box. No public tool or save. */
export function centeredSlideImageBox(width: number, height: number): SlideImageBox {
  const slide = current(false);
  const scale = Math.min((slide.width * 0.55) / width, (slide.height * 0.32) / height);
  const widthMm = width * scale;
  const heightMm = height * scale;
  return { xMm: (slide.width - widthMm) / 2, yMm: (slide.height - heightMm) / 2, widthMm, heightMm };
}

export async function prepareSlideImageInsertion(
  png: Uint8Array,
  box: SlideImageBox,
  signal?: AbortSignal,
  { description, allowSelectedObjects = false }: { description?: string; allowSelectedObjects?: boolean } = {},
) {
  const dimensions = validateSlidePng(png, signal);
  const { xMm, yMm, widthMm, heightMm } = box;
  if (
    ![xMm, yMm, widthMm, heightMm].every((value) => typeof value === 'number' && Number.isFinite(value)) ||
    xMm < 0 ||
    yMm < 0 ||
    widthMm < 1 ||
    heightMm < 1
  )
    throw new TypeError('Image box requires finite nonnegative millimeter positions and dimensions of at least 1 mm');
  const before = current(!allowSelectedObjects),
    target = captureDocumentToolTarget();
  if (target.context.kind !== 'slide' || !target.isCurrent(true)) throw new Error('Read fresh editor context');
  if (xMm + widthMm > before.width || yMm + heightMm > before.height)
    throw new Error('Image box must fit entirely on the current slide');
  const peers = before.slide.cSld.spTree.map((shape) => ({ shape, fingerprint: fingerprint(shape) }));
  const samePeers = () =>
    peers.every(
      ({ shape, fingerprint: original }, index) =>
        before.slide.cSld.spTree[index] === shape && fingerprint(shape) === original,
    );
  const assertCurrent = () => {
    signal?.throwIfAborted();
    const now = current(!allowSelectedObjects);
    if (
      !target.isCurrent(true) ||
      now.api !== before.api ||
      now.logic !== before.logic ||
      now.slide !== before.slide ||
      now.controller !== before.controller ||
      now.width !== before.width ||
      now.height !== before.height ||
      before.slide.cSld.spTree.length !== peers.length ||
      !samePeers()
    )
      throw new Error('Image insertion context changed; prepare again from fresh context');
  };
  assertCurrent();
  const { full, ready } = await preloadSlidePng(png, before, dimensions, signal);
  assertCurrent();
  ready();
  let consumed = false;
  return Object.freeze({
    page: before.page,
    box: Object.freeze({ xMm, yMm, widthMm, heightMm }),
    apply: () => {
      if (consumed) throw new Error('Image insertion was already consumed');
      consumed = true;
      assertCurrent();
      ready();
      try {
        // The SDK's addImages uses this same native action/controller path; explicit geometry avoids a second edit.
        before.logic.StartAction(before.description);
        let image: ImageShape;
        try {
          before.controller.resetSelection();
          image = before.controller.createImage(full, xMm, yMm, widthMm, heightMm);
          if (description !== undefined) {
            image.setDescription?.(description);
            if (image.getCNvProps?.()?.descr !== description)
              throw new Error('Native image description could not be saved');
          }
          image.setParent(before.slide);
          image.addToDrawingObjects();
          before.controller.selectObject(image, 0);
        } finally {
          before.logic.FinalizeAction();
        }
        before.logic.Document_UpdateInterfaceState();
        before.logic.CheckEmptyPlaceholderNotes();
        const after = current(false),
          transform = image.spPr?.xfrm;
        if (
          signal?.aborted ||
          after.api !== before.api ||
          after.logic !== before.logic ||
          after.slide !== before.slide ||
          after.controller !== before.controller ||
          after.width !== before.width ||
          after.height !== before.height ||
          before.slide.cSld.spTree.length !== peers.length + 1 ||
          before.slide.cSld.spTree[peers.length] !== image ||
          !samePeers() ||
          !image.Id ||
          peers.some((peer) => peer.shape.Id === image.Id) ||
          after.controller.selectedObjects.length !== 1 ||
          after.controller.selectedObjects[0] !== image ||
          after.controller.getTargetDocContent?.() ||
          image.group ||
          image.isPlaceholder?.() !== false ||
          image.getObjectType?.() !== before.imageType ||
          image.spPr?.geometry?.preset !== 'rect' ||
          image.blipFill?.RasterImageId !== before.common.g_oDocumentUrls.getImageLocal(full) ||
          image.blipFill.srcRect != null ||
          image.blipFill.tile != null ||
          image.blipFill.transparent != null ||
          transform?.offX !== xMm ||
          transform.offY !== yMm ||
          transform.extX !== widthMm ||
          transform.extY !== heightMm ||
          (transform.rot != null && transform.rot !== 0) ||
          transform.flipH ||
          transform.flipV
        )
          throw new Error('Native image insertion readback mismatch');
        return {
          page: before.page,
          imageId: image.Id,
          box: { xMm, yMm, widthMm, heightMm },
          verified: true as const,
          saved: false as const,
        };
      } catch (cause) {
        throw new Error('Image insertion unverified. Inspect and Undo if needed; do not retry automatically.', {
          cause,
        });
      }
    },
  });
}
