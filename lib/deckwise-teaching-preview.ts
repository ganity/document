import type { SpeechTrigger } from './deckwise-teaching-animation-bindings';
import { getEditorApi } from './agent-plugin/editor-bridge';
import { deckwisePageLineage } from './deckwise-page-lineage';
import { hasUnsavedChanges } from './unsaved-guard';

type Point = { X: number; Y: number; Error?: boolean };
type Rect = { x: number; y: number; w: number; h: number };
export type Shape = {
  Id?: string;
  group?: unknown;
  spTree?: Shape[];
  extX: number;
  extY: number;
  transform: { TransformPointX(x: number, y: number): number; TransformPointY(x: number, y: number): number };
  getUniNvProps?(): { cNvPr?: { id?: number | string; name?: string } };
};
export type PreviewObject = {
  objectId: string;
  sourceShapeId?: string;
  sourceObjectName?: string;
  status: string;
  effectTargetEligible?: boolean;
};
export type AttentionCue = {
  trigger?: SpeechTrigger;
  nativeAnimation?: string;
  id: string;
  atMs: number;
  endMs: number;
  effect: {
    type: string;
    effect?: string;
    elementId?: string;
    elementIds?: string[];
    color?: string;
    style?: string;
    opacity?: number;
    dimOpacity?: number;
  };
};
export type TimelinePreview = {
  start(): void;
  pause(): void;
  resume(): void;
  continueTeacher(): void;
  seek(ms: number): void;
  dispose(): void;
};
export type Runtime = {
  evaluateQuestionAnswer(
    response: { kind: string; options: string[]; answerIndices: number[] },
    selected: number[],
  ): boolean | null;
  nativeMotionPreset(
    effect: string,
  ): { presetClass: number; presetId: number; subtype: number; ooxmlClass: string } | undefined;
  createTeachingTimelinePreview(input: {
    pageId: string;
    video?: HTMLVideoElement;
    native?: {
      sourceSha256: string;
      stepCount: number;
      next(index: number): Promise<void>;
      pause(): void;
      resume(): void;
    };
    speechMotions?: Array<{
      actionId: string;
      stepIndex: number;
      atMs?: number;
      endMs?: number;
      trigger?: SpeechTrigger;
    }>;
    nodes: Array<Record<string, unknown>>;
    onMode(mode: string): void;
    onProgress(progress: { actionId: string; timeMs: number; durationMs: number } | null): void;
    onMediaProgress?(progress: { actionId: string; timeMs: number; durationMs: number } | null): void;
    onInteraction?(
      question: {
        checkpointId: string;
        question: string;
        referenceAnswer: { summary: string; keyPoints: string[] };
        response?: { kind: string; options: string[]; answerIndices: number[] };
      } | null,
    ): void;
    onNode(id: string): void;
    onError(error: Error): void;
    onComplete(): void;
  }): TimelinePreview;
  buildEvenOddPath(rects: Rect[]): string;
  SpeechCueScheduler: new (
    cues: [],
    time: () => number,
    cue: () => void,
    reconcile: undefined,
    onTime: (ms: number) => void,
  ) => {
    start(): void;
    pause(): void;
    resume(): void;
    stop(): void;
    seek(time?: number): void;
  };
};
let runtime: Promise<Runtime> | undefined;
export const loadTeachingPreviewRuntime = (): Promise<Runtime> => {
  const url = new URL('/.deckwise/teaching-preview.mjs?v=5', location.href).href;
  return (runtime ??= import(/* @vite-ignore */ url).catch((error) => {
    runtime = undefined;
    throw error;
  }));
};

/** SDK child transforms already include their ancestor groups. Reject detached children. */
export function teachingShapes<T extends { group?: unknown; spTree?: T[] }>(tree: T[], parent?: T): T[] {
  return tree.flatMap((shape) =>
    shape.group && shape.group !== parent ? [] : [shape, ...teachingShapes(shape.spTree ?? [], shape)],
  );
}

export function attentionTarget(tree: Shape[], objects: PreviewObject[], id: string): Shape {
  const matches = objects.filter(
    (o) =>
      o.objectId === id && o.status === 'mapped' && o.effectTargetEligible && (o.sourceShapeId || o.sourceObjectName),
  );
  if (matches.length !== 1) throw new Error('提示对象缺少当前映射，请保存并重新打开 PPT');
  const proof = matches[0];
  if (
    proof.sourceObjectName &&
    objects.filter((o) => o.status === 'mapped' && o.sourceObjectName === proof.sourceObjectName).length !== 1
  )
    throw new Error('提示对象映射不唯一，请重新读取本页');
  // The native converter renumbers OOXML IDs; generated scene names survive conversion.
  const shapes = teachingShapes(tree).filter((s) => {
    const props = s.getUniNvProps?.()?.cNvPr;
    return proof.sourceObjectName ? props?.name === proof.sourceObjectName : String(props?.id) === proof.sourceShapeId;
  });
  if (shapes.length !== 1) throw new Error('提示对象已删除或不再唯一，请重新选择目标');
  return shapes[0];
}

/** Resolve selected native objects through the same identity proof as canvas previews. */
export function selectedAttentionTargets(tree: Shape[], selected: Shape[], objects: PreviewObject[]): string[] {
  if (!selected.length || selected.length > 200 || new Set(selected).size !== selected.length)
    throw new Error('请先在画布上选择 1–200 个教学对象');
  return selected.map((shape) => {
    if (!teachingShapes(tree).includes(shape)) throw new Error('请选择当前页的教学对象');
    const matches = objects.filter((object) => {
      if (object.status !== 'mapped' || !object.effectTargetEligible) return false;
      const props = shape.getUniNvProps?.()?.cNvPr;
      return object.sourceObjectName
        ? props?.name === object.sourceObjectName
        : !!object.sourceShapeId && String(props?.id) === object.sourceShapeId;
    });
    if (matches.length !== 1 || attentionTarget(tree, objects, matches[0].objectId) !== shape)
      throw new Error('选中对象缺少唯一的教学映射，请保存并重新读取本页');
    return matches[0].objectId;
  });
}

/** Read live native transforms; do not infer identities from names or drawing order. */
export function attentionGeometry(
  tree: Shape[],
  objects: PreviewObject[],
  ids: string[],
  width: number,
  height: number,
  convert: (x: number, y: number) => Point,
) {
  if (![width, height].every((v) => Number.isFinite(v) && v > 0) || ids.length < 1 || ids.length > 200)
    throw new Error('Invalid native attention geometry');
  const origin = convert(0, 0),
    end = convert(width, height);
  if (
    [origin, end].some((p) => p.Error || !Number.isFinite(p.X) || !Number.isFinite(p.Y)) ||
    end.X <= origin.X ||
    end.Y <= origin.Y
  )
    throw new Error('Native slide viewport is unavailable');
  const targets = [...new Set(ids)].map((id) => {
    const shape = attentionTarget(tree, objects, id);
    if (![shape.extX, shape.extY].every((v) => Number.isFinite(v) && v > 0) || !shape.transform)
      throw new Error('Unsupported native target geometry');
    const points = [
      [0, 0],
      [shape.extX, 0],
      [0, shape.extY],
      [shape.extX, shape.extY],
    ].map(([x, y]) => convert(shape.transform.TransformPointX(x, y), shape.transform.TransformPointY(x, y)));
    if (points.some((p) => p.Error || !Number.isFinite(p.X) || !Number.isFinite(p.Y)))
      throw new Error('Invalid native transform');
    const x = Math.min(...points.map((p) => p.X)),
      y = Math.min(...points.map((p) => p.Y));
    return {
      id,
      shape,
      rect: {
        x: (100 * (x - origin.X)) / (end.X - origin.X) - 0.4,
        y: (100 * (y - origin.Y)) / (end.Y - origin.Y) - 0.6,
        w: (100 * (Math.max(...points.map((p) => p.X)) - x)) / (end.X - origin.X) + 0.8,
        h: (100 * (Math.max(...points.map((p) => p.Y)) - y)) / (end.Y - origin.Y) + 1.2,
      },
    };
  });
  return { origin, end, targets };
}

/** Read-only overlay: no history entries, document mutation or second playback clock. */
export async function createTeachingAttentionPreview(input: {
  doc: Document;
  page: number;
  sourceSha256: string;
  objects: PreviewObject[];
  onError(error: Error): void;
  slideshow?: HTMLElement;
}) {
  if (hasUnsavedChanges()) throw new Error('请先保存并重新打开 PPT，再预览教学提示');
  const shared = await loadTeachingPreviewRuntime();
  const api = getEditorApi() as any;
  const logic = api?.WordControl?.m_oLogicDocument;
  const slide = logic?.Slides?.[input.page - 1];
  const drawing = api?.WordControl?.m_oDrawingDocument;
  const viewport = input.slideshow ?? input.doc.getElementById('id_main_view');
  const manager = api?.WordControl?.DemonstrationManager;
  if (!slide || !drawing?.ConvertCoordsToCursor3 || !viewport || (!input.slideshow && api.isSlideShow?.()))
    throw new Error(
      `Native preview unavailable: slide=${!!slide}, coordinates=${!!drawing?.ConvertCoordsToCursor3}, viewport=${!!viewport}`,
    );
  if (deckwisePageLineage.context()?.sourceSha256 !== input.sourceSha256)
    throw new Error('页面来源已变化，请保存并重新打开当前 PPT 后预览教学提示');
  if (hasUnsavedChanges()) throw new Error('请先保存并重新打开 PPT，再预览教学提示');
  const layer = input.doc.createElement('div');
  layer.dataset.teachingPreview = '';
  layer.setAttribute('aria-hidden', 'true');
  layer.style.cssText = 'position:fixed;overflow:hidden;pointer-events:none;z-index:30;';
  (input.slideshow ?? input.doc.body).append(layer);
  let cues: AttentionCue[] = [],
    frame = 0,
    disposed = false,
    signature = '';
  const identities = new Map<string, Shape>();
  const dispose = () => {
    disposed = true;
    cancelAnimationFrame(frame);
    layer.remove();
  };
  const update = () => {
    if (disposed) return;
    try {
      if (
        hasUnsavedChanges() ||
        getEditorApi() !== api ||
        logic.Slides[input.page - 1] !== slide ||
        (input.slideshow
          ? !manager?.Mode || manager.DemonstrationDiv !== input.slideshow || manager.SlideNum !== input.page - 1
          : api.getCurrentPage() !== input.page - 1 || api.isSlideShow?.()) ||
        deckwisePageLineage.context()?.sourceSha256 !== input.sourceSha256
      )
        throw new Error('教学预览页面已变化');
      const supported = cues.filter((c) => ['highlight', 'spotlight'].includes(c.effect.type));
      const ids = supported.flatMap((c) =>
        c.effect.type === 'highlight'
          ? (c.effect.elementIds ?? [])
          : [c.effect.elementId!, ...(c.effect.elementIds ?? [])],
      );
      if (!ids.length) layer.replaceChildren();
      else {
        const g = attentionGeometry(
          slide.cSld.spTree,
          input.objects,
          [...new Set(ids)],
          logic.GetWidthMM(),
          logic.GetHeightMM(),
          (x, y) => {
            if (!input.slideshow) return drawing.ConvertCoordsToCursor3(x, y, true);
            const rect = manager.Transition.Rect,
              canvas = manager.Canvas;
            const bounds = canvas.getBoundingClientRect();
            const sx = bounds.width / canvas.width,
              sy = bounds.height / canvas.height;
            return {
              X: bounds.left + (rect.x + (x * rect.w) / logic.GetWidthMM()) * sx,
              Y: bounds.top + (rect.y + (y * rect.h) / logic.GetHeightMM()) * sy,
            };
          },
        );
        for (const t of g.targets) {
          if (identities.has(t.id) && identities.get(t.id) !== t.shape)
            throw new Error('提示对象已替换，请重新读取本页');
          identities.set(t.id, t.shape);
        }
        const clip = viewport.getBoundingClientRect();
        const next = JSON.stringify([cues, g.origin, g.end, g.targets.map((t) => t.rect), clip.toJSON()]);
        if (signature !== next) {
          signature = next;
          Object.assign(layer.style, {
            left: `${clip.left}px`,
            top: `${clip.top}px`,
            width: `${clip.width}px`,
            height: `${clip.height}px`,
          });
          const svg = input.doc.createElementNS('http://www.w3.org/2000/svg', 'svg');
          svg.setAttribute('viewBox', '0 0 100 100');
          svg.setAttribute('preserveAspectRatio', 'none');
          Object.assign(svg.style, {
            position: 'absolute',
            left: `${g.origin.X - clip.left}px`,
            top: `${g.origin.Y - clip.top}px`,
            width: `${g.end.X - g.origin.X}px`,
            height: `${g.end.Y - g.origin.Y}px`,
          });
          const spot = supported.findLast((c) => c.effect.type === 'spotlight');
          if (spot) {
            const ids = [spot.effect.elementId, ...(spot.effect.elementIds ?? [])];
            const path = input.doc.createElementNS(svg.namespaceURI, 'path');
            path.setAttribute(
              'd',
              shared.buildEvenOddPath(g.targets.filter((t) => ids.includes(t.id)).map((t) => t.rect)),
            );
            path.setAttribute('fill-rule', 'evenodd');
            path.setAttribute('fill', `rgba(0,0,0,${spot.effect.dimOpacity ?? 0.6})`);
            path.setAttribute('data-attention-kind', 'spotlight');
            svg.append(path);
          }
          for (const cue of supported.filter((c) => c.effect.type === 'highlight')) {
            for (const target of g.targets.filter((t) => cue.effect.elementIds?.includes(t.id))) {
              const rect = input.doc.createElementNS(svg.namespaceURI, 'rect');
              for (const [key, value] of Object.entries({
                x: target.rect.x,
                y: target.rect.y,
                width: target.rect.w,
                height: target.rect.h,
              }))
                rect.setAttribute(key, String(value));
              const color = /^#[\da-f]{6}$/i.test(cue.effect.color ?? '') ? cue.effect.color! : '#f59e0b';
              rect.setAttribute('fill', cue.effect.style === 'fill' ? color : 'none');
              rect.setAttribute('stroke', color);
              rect.setAttribute('stroke-width', '.5');
              rect.setAttribute('opacity', String(cue.effect.opacity ?? 0.8));
              if (cue.effect.style === 'glow') (rect as SVGElement).style.filter = `drop-shadow(0 0 2px ${color})`;
              rect.setAttribute('data-attention-kind', 'highlight');
              rect.setAttribute('data-object-id', target.id);
              svg.append(rect);
            }
          }
          layer.replaceChildren(svg);
        }
      }
      frame = requestAnimationFrame(update);
    } catch (error) {
      dispose();
      input.onError(error instanceof Error ? error : new Error(String(error)));
    }
  };
  update();
  return {
    dispose,
    show(next: AttentionCue[]) {
      if (!disposed) {
        cues = next;
        signature = '';
      }
    },
    shared,
  };
}
