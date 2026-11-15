type Shape = { Id: string; group?: Shape; parent?: { cSld?: { spTree?: Shape[] } }; spTree?: Shape[] };
type Target = { spid: string | null; subSpId: string | null; parent?: { onRemoveChild(target: Target): void };
  setSpid(id: string): void; setSubSpId(id: string): void; assignConnection(shapes: Record<string, Shape>): void };
const installed = new WeakSet<object>();

/** SDK v9 drops group-child timing on import because its target check only accepts top-level shapes. */
export function installNestedSlideAnimations(win: Window): boolean {
  const prototype = (win as unknown as { AscFormat?: { CSpTgt?: { prototype: Target } } }).AscFormat?.CSpTgt?.prototype;
  if (!prototype?.assignConnection) return false;
  if (installed.has(prototype)) return true;
  const assign = prototype.assignConnection;
  prototype.assignConnection = function (shapes) {
    const target = this.spid === null ? undefined : shapes[this.spid];
    let owner = target;
    const visited = new Set<Shape>();
    while (owner?.group && !visited.has(owner) && visited.size < 1000) {
      visited.add(owner);
      if (!owner.group.spTree?.includes(owner)) return assign.call(this, shapes);
      owner = owner.group;
    }
    if (!target || owner === target || owner?.group || !owner?.parent?.cSld?.spTree?.includes(owner))
      return assign.call(this, shapes);
    this.setSpid(target.Id);
    if (this.subSpId !== null) {
      const sub = shapes[this.subSpId];
      if (sub) this.setSubSpId(sub.Id);
      else this.parent?.onRemoveChild(this);
    }
  };
  installed.add(prototype);
  return true;
}
