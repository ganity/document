/** Capture native slide duplication before the new slide enters the presentation. */
type Slide = { createDuplicate?: (...args: unknown[]) => Slide };
const tracked = new WeakSet<object>();
export function trackTeachingSlideCopy(slide: Slide, copied: (duplicate: Slide) => void) {
  if (tracked.has(slide) || typeof slide.createDuplicate !== 'function') return;
  const duplicate = slide.createDuplicate;
  slide.createDuplicate = function (...args) {
    const result = duplicate.apply(this, args);
    copied(result);
    return result;
  };
  tracked.add(slide);
}
