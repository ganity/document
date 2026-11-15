import { installNestedSlideAnimations } from './nested-slide-animations';

/** Import identity must be captured before ReadGroupShapeMain regenerates format IDs. */
interface Shape {
  getUniNvProps?(): { cNvPr?: { id?: number | string } } | null;
}
interface Loader {
  ReadNvUniProp(shape?: Shape): { cNvPr?: { id?: number | string } };
}
const importedIds = new WeakMap<object, string>();

export function importedSlideShapeId(shape: object): string | undefined {
  return importedIds.get(shape);
}

export function installSlideSourceIdentity(win: Window): boolean {
  const scope = win as unknown as {
    AscCommon?: { BinaryPPTYLoader?: new (...args: unknown[]) => Loader };
    __ooSlideSourceIdentity?: boolean;
  };
  if (scope.__ooSlideSourceIdentity) return true;
  const common = scope.AscCommon;
  if (!common) return false;
  const Constructor = common.BinaryPPTYLoader;
  // Keep native IDs and the loader's duplicate checks unchanged. Only remember
  // which package object produced a live shape. The session pins initial objects
  // so later clipboard imports cannot impersonate their source.
  const wrap = (Constructor: new (...args: unknown[]) => Loader) =>
    new Proxy(Constructor, {
      construct(target, args, newTarget) {
        installNestedSlideAnimations(win);
        const loader = Reflect.construct(target, args, newTarget) as Loader;
        const read = loader.ReadNvUniProp;
        loader.ReadNvUniProp = function (shape) {
          const properties = read.call(this, shape);
          const id = properties?.cNvPr?.id;
          if (shape && typeof shape === 'object' && /^\d{1,10}$/.test(String(id))) {
            importedIds.set(shape, String(id));
          }
          return properties;
        };
        return loader;
      },
    });
  if (Constructor) common.BinaryPPTYLoader = wrap(Constructor);
  else
    Object.defineProperty(common, 'BinaryPPTYLoader', {
      configurable: true,
      set(value: new (...args: unknown[]) => Loader) {
        Object.defineProperty(common, 'BinaryPPTYLoader', {
          value: wrap(value),
          writable: true,
          configurable: true,
          enumerable: true,
        });
      },
    });
  scope.__ooSlideSourceIdentity = true;
  return true;
}
