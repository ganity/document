import { captureDocumentToolTarget } from './document-tool-action';
import { requireSlideImage } from './slide-image-replacement';
import { readSlideShapeText, type SlideTextShape } from './slide-text-read';

const transformKeys = ['offX', 'offY', 'extX', 'extY', 'rot', 'flipH', 'flipV'] as const;
export function slideImageFingerprint(shape: object, omittedKeys: readonly string[] = []) {
  const properties = Reflect.get(shape, 'spPr');
  const transform = properties?.xfrm;
  return JSON.stringify({
    id: Reflect.get(shape, 'Id'),
    image: Reflect.get(shape, 'blipFill'),
    style: Reflect.get(shape, 'style'),
    fill: properties?.Fill,
    line: properties?.ln,
    effects: properties?.effectProps,
    geometry: properties?.geometry?.preset,
    transform: transformKeys.filter((key) => !omittedKeys.includes(key)).map((key) => transform?.[key] ?? null),
    text: readSlideShapeText(shape as SlideTextShape),
  });
}
export function requireSlideImageTransform(): ReturnType<typeof requireSlideImage> {
  const selected = requireSlideImage();
  const transform = selected.image.spPr!.xfrm!;
  if ((transform.rot != null && transform.rot !== 0) || transform.flipH || transform.flipV)
    throw new Error('Image resize requires one unrotated, unflipped ordinary image');
  return selected;
}

export function canSetSlideImageTransform(): boolean {
  try {
    const target = captureDocumentToolTarget();
    const selected = requireSlideImageTransform();
    selected.slide.cSld.spTree.forEach((shape) => slideImageFingerprint(shape));
    return target.context.kind === 'slide' && target.isCurrent(true);
  } catch {
    return false;
  }
}
