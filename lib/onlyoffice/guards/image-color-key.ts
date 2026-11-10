type Color = { color?: { RGBA?: { R: number; G: number; B: number }; constructor?: unknown }; Mods?: { Mods?: Array<{ name: string; val: number }> } | null };
type ColorKeyEffect = { clrFrom?: Color; clrTo?: Color; useA?: boolean | null };
type ImageGraphics = { drawImage2?: (image: CanvasImageSource, ...args: unknown[]) => unknown };
type ImageShape = { blipFill?: { RasterImageId?: string; Effects?: unknown[] } };
type EditorWindow = Window & {
  AscFormat?: {
    CImageShape?: { prototype: { draw?: (this: ImageShape, graphics: ImageGraphics, ...args: unknown[]) => unknown } };
    CClrChange?: new (...args: never[]) => ColorKeyEffect;
    CRGBColor?: new (...args: never[]) => unknown;
  };
  AscCommon?: { getFullImageSrc2?: (id: string) => string };
  Asc?: { editor?: { ImageLoader?: { map_image_index?: Record<string, { Image?: HTMLImageElement }> } } };
  __imageColorKeyGuard?: boolean;
};

/** Adapt MAIC applyPptxImageColorChange: exact RGB matches, no heuristic color removal. */
export function applyImageColorKey(pixels: Uint8ClampedArray, rgb: readonly number[]): void {
  for (let i = 0; i < pixels.length; i += 4) {
    if (pixels[i] === rgb[0] && pixels[i + 1] === rgb[1] && pixels[i + 2] === rgb[2]) pixels[i + 3] = 0;
  }
}

/** Render supported clrChange transparency without rewriting media or the document model. */
export function installImageColorKeyGuard(win: Window): boolean {
  const scope = win as EditorWindow;
  const format = scope.AscFormat;
  const prototype = format?.CImageShape?.prototype;
  const original = prototype?.draw;
  if (!format?.CClrChange || !format.CRGBColor || !prototype || typeof original !== 'function') return false;
  if (scope.__imageColorKeyGuard) return true;
  const cache = new WeakMap<HTMLImageElement, Map<string, HTMLCanvasElement>>();

  prototype.draw = function (graphics, ...args) {
    const effects = this.blipFill?.Effects;
    const effect = effects?.length === 1 && effects[0] instanceof format.CClrChange!
      ? effects[0] as ColorKeyEffect : null;
    const from = effect?.clrFrom;
    const to = effect?.clrTo;
    const alpha = to?.Mods?.Mods;
    const rgb = from?.color?.RGBA;
    const target = to?.color?.RGBA;
    const drawImage = graphics.drawImage2;
    const id = this.blipFill?.RasterImageId;
    const fullId = id && scope.AscCommon?.getFullImageSrc2?.(id);
    const source = fullId ? scope.Asc?.editor?.ImageLoader?.map_image_index?.[fullId]?.Image : undefined;
    if (!effect || effect.useA === true || !(from?.color instanceof format.CRGBColor!) ||
        !(to?.color instanceof format.CRGBColor!) || from.Mods?.Mods?.length ||
        alpha?.length !== 1 || alpha[0].name !== 'alpha' || alpha[0].val !== 0 ||
        !rgb || !target || rgb.R !== target.R || rgb.G !== target.G || rgb.B !== target.B ||
        !source || !source.complete || typeof drawImage !== 'function') {
      return original.call(this, graphics, ...args);
    }
    const key = `${rgb.R},${rgb.G},${rgb.B}`;
    let entries = cache.get(source);
    let canvas = entries?.get(key);
    if (!canvas) {
      const width = source.naturalWidth, height = source.naturalHeight;
      if (width <= 0 || height <= 0 || width * height > 16_000_000) return original.call(this, graphics, ...args);
      canvas = win.document.createElement('canvas');
      canvas.width = width; canvas.height = height;
      const context = canvas.getContext('2d', { willReadFrequently: true });
      if (!context) return original.call(this, graphics, ...args);
      try {
        context.drawImage(source, 0, 0);
        const image = context.getImageData(0, 0, width, height);
        applyImageColorKey(image.data, [rgb.R, rgb.G, rgb.B]);
        context.putImageData(image, 0, 0);
      } catch {
        return original.call(this, graphics, ...args);
      }
      entries ??= new Map();
      if (entries.size >= 4) entries.clear();
      entries.set(key, canvas); cache.set(source, entries);
    }
    const hadOwnMethod = Object.prototype.hasOwnProperty.call(graphics, 'drawImage2');
    const processed = canvas;
    graphics.drawImage2 = function (image, ...imageArgs) {
      return drawImage.call(this, image === source ? processed : image, ...imageArgs);
    };
    try {
      return original.call(this, graphics, ...args);
    } finally {
      if (hadOwnMethod) graphics.drawImage2 = drawImage;
      else delete graphics.drawImage2;
    }
  };
  scope.__imageColorKeyGuard = true;
  return true;
}
