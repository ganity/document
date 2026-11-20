import { requireEditorContext, getEditorContext } from './editor-bridge';
import { getReadonlyMode } from '../onlyoffice/readonly';
import { preloadSlidePng, validateSlidePng, type PngApi, type PngCommon } from './slide-image-png';

type Box = [number, number, number, number];

/** Pixel partition, not generative inpainting: the crop restores every removed pixel. */
export function splitBackgroundPixels(pixels: Uint8ClampedArray, width: number, height: number, bounds: number[]) {
  if (
    ![width, height].every((n) => Number.isSafeInteger(n) && n > 0 && n <= 4096) ||
    pixels.length !== width * height * 4 ||
    !Array.isArray(bounds) ||
    bounds.length !== 4 ||
    !bounds.every(Number.isFinite) ||
    bounds[0] < 0 ||
    bounds[1] < 0 ||
    bounds[2] <= 0 ||
    bounds[3] <= 0 ||
    bounds[0] + bounds[2] > 1 ||
    bounds[1] + bounds[3] > 1
  )
    throw new Error('Invalid background region');
  const x = Math.floor(bounds[0] * width),
    y = Math.floor(bounds[1] * height);
  const w = Math.ceil((bounds[0] + bounds[2]) * width) - x,
    h = Math.ceil((bounds[1] + bounds[3]) * height) - y;
  if (w * h >= width * height || pixels.some((v, i) => i % 4 === 3 && v !== 255))
    throw new Error('Region repair requires a partial region of an opaque background');
  const background = pixels.slice(),
    region = new Uint8ClampedArray(w * h * 4);
  for (let row = 0; row < h; row++) {
    const offset = ((y + row) * width + x) * 4;
    region.set(pixels.subarray(offset, offset + w * 4), row * w * 4);
    background.fill(255, offset, offset + w * 4);
  }
  return { background, region, box: [x, y, w, h] as Box };
}

type Fill = {
  Effects?: Array<{ bright?: number | null; contrast?: number | null }>;
  stretch?: { fillRect?: unknown };
  RasterImageId?: string;
  srcRect?: unknown;
  tile?: unknown;
  transparent?: unknown;
};
type Background = { bgPr?: { Fill?: { fill?: Fill; transparent?: unknown } }; createFullCopy(): Background };
type Shape = {
  Id: string;
  setName(name: string): void;
  setParent(slide: Slide): void;
  getUniNvProps?(): { cNvPr?: { id?: number | string; name?: string } };
};
type Slide = {
  cSld: { Bg: Background; spTree: Shape[] };
  changeBackground(bg: Background): void;
  addToSpTreeToPos(index: number, shape: Shape): void;
  removeFromSpTreeByPos(index: number): void;
};
type Logic = {
  Slides: Slide[];
  CanEdit(): boolean;
  IsMasterMode(): boolean;
  IsSelectionLocked(kind: unknown, slides?: Slide[]): boolean;
  GetWidthMM(): number;
  GetHeightMM(): number;
  GetCurrentController(): { createImage(url: string, x: number, y: number, w: number, h: number): Shape };
  Recalculate(data: unknown): void;
  Document_UpdateInterfaceState(): void;
  DrawingDocument: { UpdateThumbnailsAttack(): void };
};
type Api = PngApi & {
  canEdit(): boolean;
  isSlideShow(): boolean;
  isLongAction(): boolean;
  WordControl: { m_oLogicDocument: Logic };
};
type Common = PngCommon & {
  changestype_SlideBg: unknown;
  changestype_Drawing_Props: unknown;
  CollaborativeEditing?: { Get_GlobalLock(): boolean };
  History: { Index: number; Points: unknown[]; _getLongPointIndex(): number };
};

function context(page: number) {
  const native = requireEditorContext();
  const format = native.AscFormat as unknown as {
    ExecuteNoHistory<T>(fn: () => T, scope: null, args: [], keepIds: boolean): T;
  };
  const api = native.api as Api,
    common = native.AscCommon as unknown as Common;
  const logic = api.WordControl?.m_oLogicDocument,
    slide = logic?.Slides?.[page - 1];
  const bg = slide?.cSld.Bg,
    fill = bg?.bgPr?.Fill?.fill;
  if (
    !Number.isSafeInteger(page) ||
    page < 1 ||
    getReadonlyMode() ||
    !slide ||
    !bg ||
    !fill?.RasterImageId ||
    typeof bg.createFullCopy !== 'function' ||
    typeof format?.ExecuteNoHistory !== 'function' ||
    typeof slide.changeBackground !== 'function' ||
    typeof slide.addToSpTreeToPos !== 'function' ||
    typeof slide.removeFromSpTreeByPos !== 'function' ||
    fill.Effects?.some(
      (effect) =>
        Object.keys(effect).some((key) => !['bright', 'contrast'].includes(key)) ||
        (effect.bright != null && effect.bright !== 0) ||
        (effect.contrast != null && effect.contrast !== 0),
    ) ||
    fill.stretch?.fillRect != null ||
    fill.srcRect != null ||
    fill.tile != null ||
    fill.transparent != null ||
    bg.bgPr?.Fill?.transparent != null ||
    api.canEdit?.() !== true ||
    logic.CanEdit?.() !== true ||
    logic.IsMasterMode?.() !== false ||
    api.isSlideShow?.() ||
    api.isLongAction?.() ||
    common.CollaborativeEditing?.Get_GlobalLock() ||
    !common.History ||
    common.History._getLongPointIndex?.() !== -1 ||
    common.changestype_SlideBg === undefined ||
    common.changestype_Drawing_Props === undefined ||
    logic.IsSelectionLocked(common.changestype_SlideBg, [slide]) ||
    logic.IsSelectionLocked(common.changestype_Drawing_Props)
  )
    throw new Error('Region repair requires an editable slide with its own ordinary image background');
  return { api, common, logic, slide, bg, fill, format, width: logic.GetWidthMM(), height: logic.GetHeightMM() };
}

export function canExtractBackgroundRegion(page: number): boolean {
  try {
    context(page);
    return true;
  } catch {
    return false;
  }
}

/** Prepare detached images only; caller commits background, crop, cues and carrier in one Undo action. */
export async function prepareBackgroundRegion(page: number, bounds: number[], objectId: string, signal?: AbortSignal) {
  const before = context(page),
    { api, common, logic, slide, bg } = before;
  if (!/^[\w.-]{1,128}$/.test(objectId) || ![before.width, before.height].every((n) => Number.isFinite(n) && n > 0))
    throw new Error('Invalid region identity or canvas');
  const history = common.History,
    index = history.Index,
    point = history.Points[index];
  const originals = [...slide.cSld.spTree],
    originalFill = JSON.stringify(bg);
  const assertCurrent = () => {
    signal?.throwIfAborted();
    const now = context(page);
    if (
      getEditorContext()?.api !== api ||
      now.slide !== slide ||
      now.bg !== bg ||
      JSON.stringify(bg) !== originalFill ||
      now.width !== before.width ||
      now.height !== before.height ||
      history.Index !== index ||
      history.Points[index] !== point ||
      slide.cSld.spTree.length !== originals.length ||
      originals.some((shape, i) => slide.cSld.spTree[i] !== shape)
    )
      throw new Error('Background region source changed during preparation');
  };
  assertCurrent();
  const url = common.getFullImageSrc2(before.fill.RasterImageId!);
  const image = api.ImageLoader?.LoadImage(url, 1)?.Image as HTMLImageElement | undefined;
  if (!image || !image.naturalWidth || !image.naturalHeight || image.naturalWidth > 4096 || image.naturalHeight > 4096)
    throw new Error('Background pixels are not loaded or exceed the region budget');
  const canvas = document.createElement('canvas');
  canvas.width = image.naturalWidth;
  canvas.height = image.naturalHeight;
  const ctx = canvas.getContext('2d');
  if (!ctx) throw new Error('Region pixel canvas unavailable');
  let split: ReturnType<typeof splitBackgroundPixels>;
  try {
    ctx.drawImage(image, 0, 0);
    split = splitBackgroundPixels(
      ctx.getImageData(0, 0, canvas.width, canvas.height).data,
      canvas.width,
      canvas.height,
      bounds,
    );
  } finally {
    canvas.width = canvas.height = 0;
  }
  const png = (pixels: Uint8ClampedArray, width: number, height: number) => {
    const output = document.createElement('canvas');
    output.width = width;
    output.height = height;
    try {
      const target = output.getContext('2d');
      if (!target) throw new Error('Region output canvas unavailable');
      const data = target.createImageData(width, height);
      data.data.set(pixels);
      target.putImageData(data, 0, 0);
      return Uint8Array.from(atob(output.toDataURL('image/png').split(',')[1]), (c) => c.charCodeAt(0));
    } finally {
      output.width = output.height = 0;
    }
  };
  const [x, y, w, h] = split.box;
  const background = png(split.background, image.naturalWidth, image.naturalHeight),
    crop = png(split.region, w, h);
  const plate = await preloadSlidePng(background, before, validateSlidePng(background), signal);
  assertCurrent();
  const region = await preloadSlidePng(crop, before, validateSlidePng(crop), signal);
  assertCurrent();
  plate.ready();
  region.ready();
  const nextBackground = bg.createFullCopy();
  const local = common.g_oDocumentUrls.getImageLocal(plate.full);
  if (!local || !nextBackground.bgPr?.Fill?.fill) throw new Error('Prepared background unavailable');
  nextBackground.bgPr.Fill.fill.RasterImageId = local;
  const nativeShape = before.format.ExecuteNoHistory(
    () => {
      const shape = logic
        .GetCurrentController()
        .createImage(
          region.full,
          (x / image.naturalWidth) * before.width,
          (y / image.naturalHeight) * before.height,
          (w / image.naturalWidth) * before.width,
          (h / image.naturalHeight) * before.height,
        );
      shape.setName(objectId);
      shape.setParent(slide);
      return shape;
    },
    null,
    [],
    true,
  );
  assertCurrent();
  const redraw = () => {
    logic.Recalculate({ Drawings: { All: true, Map: {} } });
    logic.Document_UpdateInterfaceState();
    logic.DrawingDocument.UpdateThumbnailsAttack();
  };
  return {
    nativeShape,
    objectId,
    assertCurrent,
    apply() {
      plate.ready();
      region.ready();
      slide.changeBackground(nextBackground);
      // Keep every existing foreground in its original order above the background crop.
      slide.addToSpTreeToPos(0, nativeShape);
      redraw();
      if (
        slide.cSld.Bg !== nextBackground ||
        slide.cSld.spTree[0] !== nativeShape ||
        slide.cSld.spTree.length !== originals.length + 1 ||
        originals.some((s, i) => slide.cSld.spTree[i + 1] !== s)
      )
        throw new Error('Region readback differs from the prepared crop');
    },
    restore() {
      const position = slide.cSld.spTree.indexOf(nativeShape);
      if (position >= 0) slide.removeFromSpTreeByPos(position);
      slide.changeBackground(bg);
      redraw();
    },
  };
}
