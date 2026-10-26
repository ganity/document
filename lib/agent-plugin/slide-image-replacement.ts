import { validateSlidePng, preloadSlidePng, type PngCommon, type PngApi } from './slide-image-png';
import { captureDocumentToolTarget } from './document-tool-action';
import { requireEditorContext } from './editor-bridge';
import { getReadonlyMode } from '../onlyoffice/readonly';

interface ImageShape {
  Id: string;
  group?: unknown;
  isPlaceholder?(): boolean;
  getObjectType?(): number;
  blipFill?: { RasterImageId?: string; srcRect?: unknown; tile?: unknown; transparent?: unknown };
  spPr?: { geometry?: { preset?: string }; xfrm?: Record<string, number | boolean | null | undefined> };
}
interface ImageApi extends PngApi {
  canEdit?(): boolean;
  isSlideShow?(): boolean;
  isLongAction?(): boolean;
  getCurrentPage?(): number;
  ImgApply?(props: { ImageUrl?: string; Width?: number; Height?: number; Position?: { X: number; Y: number } }): void;
  WordControl?: {
    m_oLogicDocument?: {
      Slides: Array<{ cSld: { spTree: ImageShape[] } }>;
      CanEdit?(): boolean;
      IsMasterMode?(): boolean;
      GetWidthMM(): number;
      GetHeightMM(): number;
      GetCurrentController?(): { selectedObjects?: ImageShape[]; getTargetDocContent?(): unknown };
    };
  };
}
interface ImageCommon extends PngCommon {
  CollaborativeEditing?: { Get_GlobalLock(): boolean };
}
const geometryKeys = ['offX', 'offY', 'extX', 'extY', 'rot', 'flipH', 'flipV'] as const;
export function requireSlideImage() {
  const context = requireEditorContext();
  const api = context.api as ImageApi;
  const common = context.AscCommon as typeof context.AscCommon & Partial<ImageCommon>;
  const logic = api.WordControl?.m_oLogicDocument;
  const page = api.getCurrentPage?.();
  const slide = page === undefined ? undefined : logic?.Slides?.[page];
  const controller = logic?.GetCurrentController?.();
  const image = controller?.selectedObjects?.[0];
  const xfrm = image?.spPr?.xfrm;
  const imageType = Reflect.get(context.AscDFH ?? {}, 'historyitem_type_ImageShape');
  if (
    getReadonlyMode() ||
    api.isViewMode ||
    api.canEdit?.() !== true ||
    logic?.CanEdit?.() !== true ||
    !api.isDocumentLoadComplete ||
    !api.isLoadFullApi ||
    api.isSlideShow?.() ||
    api.isLongAction?.() ||
    logic.IsMasterMode?.() ||
    common?.CollaborativeEditing?.Get_GlobalLock() ||
    typeof api.ImgApply !== 'function' ||
    typeof api.ImageLoader?.LoadImage !== 'function' ||
    typeof api.ImageLoader?.LoadImagesWithCallback !== 'function' ||
    typeof common?.sendImgUrls !== 'function' ||
    typeof common.getFullImageSrc2 !== 'function' ||
    typeof common.g_oDocumentUrls?.getImageLocal !== 'function' ||
    typeof common.g_oDocumentUrls.imagePath2Local !== 'function' ||
    !Number.isSafeInteger(imageType) ||
    !slide ||
    !image ||
    controller?.selectedObjects?.length !== 1 ||
    controller.getTargetDocContent?.() ||
    !slide.cSld.spTree.includes(image) ||
    image.getObjectType?.() !== imageType ||
    image.isPlaceholder?.() !== false ||
    image.group ||
    !image.Id ||
    !image.blipFill?.RasterImageId ||
    image.blipFill.srcRect != null ||
    image.blipFill.tile != null ||
    image.blipFill.transparent != null ||
    image.spPr?.geometry?.preset !== 'rect' ||
    !xfrm
  )
    throw new Error('Image replacement requires one writable, ungrouped, uncropped ordinary image');
  const width = logic.GetWidthMM(),
    height = logic.GetHeightMM();
  const { offX, offY, extX, extY, rot, flipH, flipV } = xfrm;
  if (
    ![offX, offY, extX, extY, width, height].every((v) => typeof v === 'number' && Number.isFinite(v)) ||
    (rot != null && (typeof rot !== 'number' || !Number.isFinite(rot))) ||
    [flipH, flipV].some((value) => value != null && typeof value !== 'boolean') ||
    Number(offX) < 0 ||
    Number(offY) < 0 ||
    Number(extX) <= 0 ||
    Number(extY) <= 0 ||
    Number(offX) + Number(extX) > width ||
    Number(offY) + Number(extY) > height
  )
    throw new Error('Image must fit entirely on the current slide');
  const angle = Number(rot ?? 0);
  const halfWidth = (Math.abs(Math.cos(angle)) * Number(extX) + Math.abs(Math.sin(angle)) * Number(extY)) / 2;
  const halfHeight = (Math.abs(Math.sin(angle)) * Number(extX) + Math.abs(Math.cos(angle)) * Number(extY)) / 2;
  const centerX = Number(offX) + Number(extX) / 2,
    centerY = Number(offY) + Number(extY) / 2;
  if (
    centerX - halfWidth < -1e-9 ||
    centerY - halfHeight < -1e-9 ||
    centerX + halfWidth > width + 1e-9 ||
    centerY + halfHeight > height + 1e-9
  )
    throw new Error('Rotated image must fit entirely on the current slide');
  const geometry = JSON.stringify(geometryKeys.map((key) => xfrm[key] ?? null));
  return { api, common: common as ImageCommon, logic, slide, image, page: page! + 1, geometry, width, height };
}

export function canReplaceSlideImage(): boolean {
  try {
    const target = captureDocumentToolTarget();
    return target.context.kind === 'slide' && target.isCurrent(true) && !!requireSlideImage();
  } catch {
    return false;
  }
}

/** Host-only preparation. No model URLs, tool registration, save or source-file writes. */
export async function prepareSlideImageReplacement(png: Uint8Array, signal?: AbortSignal) {
  const dimensions = validateSlidePng(png, signal);
  const before = requireSlideImage();
  const target = captureDocumentToolTarget();
  if (target.context.kind !== 'slide' || !target.isCurrent(true)) throw new Error('Read fresh editor context');
  const assertCurrent = () => {
    signal?.throwIfAborted();
    const now = requireSlideImage();
    if (
      !target.isCurrent(true) ||
      now.api !== before.api ||
      now.slide !== before.slide ||
      now.image !== before.image ||
      now.geometry !== before.geometry ||
      now.width !== before.width ||
      now.height !== before.height ||
      now.image.blipFill?.RasterImageId !== beforeUrl
    )
      throw new Error('Image context changed; prepare again from fresh context');
  };
  const beforeUrl = before.image.blipFill!.RasterImageId;
  assertCurrent();
  const { full, ready } = await preloadSlidePng(png, before, dimensions, signal);
  assertCurrent();
  ready();
  let consumed = false;
  return Object.freeze({
    page: before.page,
    imageId: before.image.Id,
    apply: () => {
      if (consumed) throw new Error('Image replacement was already consumed');
      consumed = true;
      assertCurrent();
      ready();
      // Native ImgApply is synchronous only for a registered, already-loaded image.
      try {
        before.api.ImgApply!({ ImageUrl: full });
        const after = requireSlideImage();
        const expected = before.common.g_oDocumentUrls.getImageLocal(full);
        if (
          signal?.aborted ||
          after.api !== before.api ||
          after.logic !== before.logic ||
          after.slide !== before.slide ||
          after.image !== before.image ||
          after.geometry !== before.geometry ||
          after.width !== before.width ||
          after.height !== before.height ||
          after.image.blipFill?.RasterImageId !== expected
        )
          throw new Error('Native image readback mismatch');
      } catch (cause) {
        throw new Error('Image change unverified. Inspect and Undo if needed; do not retry automatically.', { cause });
      }
      return { page: before.page, imageId: before.image.Id, verified: true as const };
    },
  });
}
