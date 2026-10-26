import { captureDocumentToolTarget } from './document-tool-action';
import { canSetSlideImageTransform, requireSlideImageTransform, slideImageFingerprint } from './slide-image-transform';

export const canSetSlideImageSize = canSetSlideImageTransform;

/** Finite native resize only; no asset loading, target selection, file writes or save. */
export function setSlideImageSize(input: unknown, signal?: AbortSignal) {
  if (
    !input ||
    typeof input !== 'object' ||
    Array.isArray(input) ||
    Object.keys(input).length !== 2 ||
    Object.keys(input).some((key) => !['widthMm', 'heightMm'].includes(key))
  )
    throw new Error('Image resize requires only widthMm and heightMm');
  const widthMm: unknown = Reflect.get(input, 'widthMm'),
    heightMm: unknown = Reflect.get(input, 'heightMm');
  if (
    typeof widthMm !== 'number' ||
    typeof heightMm !== 'number' ||
    ![widthMm, heightMm].every((value) => Number.isFinite(value) && value >= 1 && value <= 1000)
  )
    throw new Error('Image dimensions must be finite millimeters between 1 and 1000');
  signal?.throwIfAborted();
  const before = requireSlideImageTransform(),
    target = captureDocumentToolTarget();
  if (!target.isCurrent(true)) throw new Error('Read fresh editor context');
  const transform = before.image.spPr!.xfrm!;
  if (Number(transform.offX) + widthMm > before.width || Number(transform.offY) + heightMm > before.height)
    throw new Error('The resized image must fit entirely on the current slide');
  const peers = before.slide.cSld.spTree.map((shape) => ({
    shape,
    value: slideImageFingerprint(shape, shape === before.image ? ['extX', 'extY'] : []),
  }));
  try {
    before.api.ImgApply!({ Width: widthMm, Height: heightMm });
    signal?.throwIfAborted();
    const after = requireSlideImageTransform(),
      resized = after.image.spPr!.xfrm!;
    if (
      after.api !== before.api ||
      after.logic !== before.logic ||
      after.slide !== before.slide ||
      after.image !== before.image ||
      after.page !== before.page ||
      after.width !== before.width ||
      after.height !== before.height ||
      resized.extX !== widthMm ||
      resized.extY !== heightMm ||
      after.slide.cSld.spTree.length !== peers.length ||
      peers.some(
        ({ shape, value }, index) =>
          after.slide.cSld.spTree[index] !== shape ||
          slideImageFingerprint(shape, shape === before.image ? ['extX', 'extY'] : []) !== value,
      )
    )
      throw new Error('Native image-size readback mismatch');
    return { widthMm, heightMm, page: before.page, verified: true as const };
  } catch (cause) {
    throw new Error('Image resize unverified. Inspect and Undo if needed; do not retry automatically.', { cause });
  }
}
