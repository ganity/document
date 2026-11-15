/**
 * 21. Render supported source image soft edges without changing media.
 */
type Graphics = { drawImage2?: (image: CanvasImageSource, ...args: unknown[]) => unknown };
type Shape = {
  extX?: number;
  extY?: number;
  blipFill?: { RasterImageId?: string; srcRect?: unknown; tile?: unknown; Effects?: unknown[] };
  spPr?: { geometry?: { preset?: string }; effectProps?: { EffectDag?: unknown; EffectLst?: Record<string, unknown> } };
  pen?: { w?: number } | null;
};
type Scope = Window & {
  AscFormat?: {
    CImageShape?: { prototype: { draw?: (this: Shape, graphics: Graphics, ...args: unknown[]) => unknown } };
  };
  AscCommon?: { getFullImageSrc2?: (id: string) => string };
  Asc?: { editor?: { ImageLoader?: { map_image_index?: Record<string, { Image?: HTMLImageElement }> } } };
  __imageSoftEdgeGuard?: boolean;
};

/** An inward ellipse feather in source-image coordinates; preserve existing alpha. */
export function featherEllipseAlpha(
  data: Uint8ClampedArray,
  width: number,
  height: number,
  widthMM: number,
  heightMM: number,
  radiusMM: number,
): void {
  const rx = widthMM / 2,
    ry = heightMM / 2;
  for (let y = 0; y < height; y++)
    for (let x = 0; x < width; x++) {
      const px = ((x + 0.5) / width - 0.5) * widthMM;
      const py = ((y + 0.5) / height - 0.5) * heightMM;
      const q = Math.hypot(px / rx, py / ry);
      const gradient = q > 0 ? Math.hypot(px / (rx * rx), py / (ry * ry)) / q : 0;
      const distance = gradient > 0 ? (1 - q) / gradient : Math.min(rx, ry);
      const t = Math.max(0, Math.min(1, distance / radiusMM));
      const index = (y * width + x) * 4 + 3;
      data[index] = Math.round(data[index] * t * t * (3 - 2 * t));
    }
}

/** Support a lone softEdge on an ordinary uncropped ellipse image at drawing time. */
export function installImageSoftEdgeGuard(win: Window): boolean {
  const scope = win as Scope;
  const prototype = scope.AscFormat?.CImageShape?.prototype;
  const original = prototype?.draw;
  if (!prototype || typeof original !== 'function') return false;
  if (scope.__imageSoftEdgeGuard) return true;
  const cache = new WeakMap<HTMLImageElement, Map<string, HTMLCanvasElement>>();
  prototype.draw = function (graphics, ...args) {
    const list = this.spPr?.effectProps?.EffectLst;
    const radius = (list?.softEdge as { rad?: number } | undefined)?.rad;
    const widthMM = this.extX!,
      heightMM = this.extY!;
    const id = this.blipFill?.RasterImageId;
    const url = id && scope.AscCommon?.getFullImageSrc2?.(id);
    const source = url ? scope.Asc?.editor?.ImageLoader?.map_image_index?.[url]?.Image : undefined;
    const drawImage = graphics.drawImage2;
    if (
      this.spPr?.geometry?.preset !== 'ellipse' ||
      this.spPr.effectProps?.EffectDag ||
      !list ||
      Object.entries(list).some(([key, value]) => key !== 'softEdge' && value != null) ||
      !Number.isFinite(radius) ||
      radius! <= 0 ||
      !Number.isFinite(widthMM) ||
      !Number.isFinite(heightMM) ||
      widthMM <= 0 ||
      heightMM <= 0 ||
      this.blipFill?.srcRect ||
      this.blipFill?.tile ||
      this.blipFill?.Effects?.length ||
      (this.pen && this.pen.w !== 0) ||
      !source?.complete ||
      typeof drawImage !== 'function'
    ) {
      return original.call(this, graphics, ...args);
    }
    const key = `${widthMM}:${heightMM}:${radius}`;
    let entries = cache.get(source),
      canvas = entries?.get(key);
    if (!canvas) {
      const width = source.naturalWidth,
        height = source.naturalHeight;
      if (width <= 0 || height <= 0 || width * height > 16_000_000) return original.call(this, graphics, ...args);
      canvas = win.document.createElement('canvas');
      canvas.width = width;
      canvas.height = height;
      const context = canvas.getContext('2d', { willReadFrequently: true });
      if (!context) return original.call(this, graphics, ...args);
      try {
        context.drawImage(source, 0, 0);
        const image = context.getImageData(0, 0, width, height);
        featherEllipseAlpha(image.data, width, height, widthMM, heightMM, radius! / 36000);
        context.putImageData(image, 0, 0);
      } catch {
        return original.call(this, graphics, ...args);
      }
      entries ??= new Map();
      if (entries.size >= 4) entries.clear();
      entries.set(key, canvas);
      cache.set(source, entries);
    }
    const own = Object.prototype.hasOwnProperty.call(graphics, 'drawImage2'),
      rendered = canvas;
    graphics.drawImage2 = function (image, ...imageArgs) {
      return drawImage.call(this, image === source ? rendered : image, ...imageArgs);
    };
    try {
      return original.call(this, graphics, ...args);
    } finally {
      if (own) graphics.drawImage2 = drawImage;
      else delete graphics.drawImage2;
    }
  };
  scope.__imageSoftEdgeGuard = true;
  return true;
}
