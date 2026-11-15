import { captureDocumentToolTarget } from './document-tool-action';
import { canSetSlideImageTransform, requireSlideImageTransform, slideImageFingerprint } from './slide-image-transform';

export const canSetSlideImagePosition = canSetSlideImageTransform;

/** Finite native positioning only; no asset loading, target selection, file writes or save. */
export function setSlideImagePosition(input: unknown, signal?: AbortSignal) {
  if (
    !input ||
    typeof input !== 'object' ||
    Array.isArray(input) ||
    Object.keys(input).length !== 2 ||
    Object.keys(input).some((key) => !['xMm', 'yMm'].includes(key))
  )
    throw new Error('Image positioning requires only xMm and yMm');
  const xMm: unknown = Reflect.get(input, 'xMm'),
    yMm: unknown = Reflect.get(input, 'yMm');
  if (
    typeof xMm !== 'number' ||
    typeof yMm !== 'number' ||
    ![xMm, yMm].every((value) => Number.isFinite(value) && value >= 0 && value <= 1000)
  )
    throw new Error('Image position must use finite millimeters between 0 and 1000');
  signal?.throwIfAborted();
  const before = requireSlideImageTransform(),
    target = captureDocumentToolTarget();
  if (!target.isCurrent(true)) throw new Error('Read fresh editor context');
  const transform = before.image.spPr!.xfrm!;
  if (xMm + Number(transform.extX) > before.width || yMm + Number(transform.extY) > before.height)
    throw new Error('The positioned image must fit entirely on the current slide');
  const peers = before.slide.cSld.spTree.map((shape) => ({
    shape,
    value: slideImageFingerprint(shape, shape === before.image ? ['offX', 'offY'] : []),
  }));
  try {
    before.api.ImgApply!({ Position: { X: xMm, Y: yMm } });
    signal?.throwIfAborted();
    const after = requireSlideImageTransform(),
      positioned = after.image.spPr!.xfrm!;
    if (
      after.api !== before.api ||
      after.logic !== before.logic ||
      after.slide !== before.slide ||
      after.image !== before.image ||
      after.page !== before.page ||
      after.width !== before.width ||
      after.height !== before.height ||
      positioned.offX !== xMm ||
      positioned.offY !== yMm ||
      after.slide.cSld.spTree.length !== peers.length ||
      peers.some(
        ({ shape, value }, index) =>
          after.slide.cSld.spTree[index] !== shape ||
          slideImageFingerprint(shape, shape === before.image ? ['offX', 'offY'] : []) !== value,
      )
    )
      throw new Error('Native image-position readback mismatch');
    return { xMm, yMm, page: before.page, verified: true as const };
  } catch (cause) {
    throw new Error('Image positioning unverified. Inspect and Undo if needed; do not retry automatically.', { cause });
  }
}
