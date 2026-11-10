import { readPresentationShapeText } from './agent-plugin/slide-text-read';
import { trackTeachingSlideCopy } from './onlyoffice/guards/slide-teaching-copy';
import { importedSlideShapeId } from './onlyoffice/guards/slide-source-identity';
import { getEditorApi } from './agent-plugin/editor-bridge';
import { getReadonlyMode } from './onlyoffice/readonly';

export const NATIVE_MEDIA_NAMESPACE = 'urn:deckwise:native-teaching-media:1';
export const NATIVE_TEACHING_NAMESPACE = 'urn:deckwise:native-teaching:1';
type Action = { id: string; type: string; [key: string]: unknown };
type Page = { pageId: string; locked: boolean; actions: Action[]; reviewState?: 'review' | 'ready' };
export type NativeTeaching = {
  schemaVersion: 'deckwise-native-teaching/1';
  documentId: string;
  revision: string;
  pages: Page[];
  media: Array<{ id: string; sha256: string; mimeType: string; byteLength: number; durationMs: number; data?: string }>;
  [key: string]: unknown;
};
type Part = { GetXml(): string; Delete(): void };
type Parts = { GetAll(): Part[]; GetByNamespace(namespace: string): Part[]; Add(xml: string): Part };
const decoded = new Map<string, NativeTeaching | undefined>();
type Logic = { StartAction(description: number): void; FinalizeAction(): void };

function session(): { parts: Parts; logic: Logic; description: number; history?: { TurnOff(): void; TurnOn(): void } } {
  const api = getEditorApi() as unknown as { WordControl?: { m_oLogicDocument?: Logic } } | null;
  for (const frame of document.querySelectorAll<HTMLIFrameElement>('iframe')) {
    const scope = frame.contentWindow as unknown as {
      editor?: unknown;
      AscBuilder?: { ApiCustomXmlParts?: new (logic: Logic) => Parts };
      AscCommon?: { History?: { TurnOff(): void; TurnOn(): void } };
      AscDFH?: { historydescription_Document_ApiBuilder?: number };
    } | null;
    if (scope?.editor !== api || !api?.WordControl?.m_oLogicDocument || !scope?.AscBuilder?.ApiCustomXmlParts) continue;
    const description = scope.AscDFH?.historydescription_Document_ApiBuilder;
    if (typeof description !== 'number') break;
    const logic = api.WordControl.m_oLogicDocument;
    return { parts: new scope.AscBuilder.ApiCustomXmlParts(logic), logic, description, history: scope.AscCommon?.History };
  }
  throw new Error('Native teaching API unavailable');
}
function decode(part: Part): NativeTeaching | undefined {
  const xml = part.GetXml();
  if (!xml.includes(NATIVE_TEACHING_NAMESPACE)) return;
  if (decoded.has(xml)) return decoded.get(xml);
  if (xml.length > 80 * 1024 * 1024 || /<!DOCTYPE|<!ENTITY/i.test(xml)) throw new Error('Invalid native teaching XML');
  const root = new DOMParser().parseFromString(xml, 'application/xml').documentElement;
  if (root.namespaceURI !== NATIVE_TEACHING_NAMESPACE || root.localName !== 'teaching' || root.children.length) throw new Error('Invalid native teaching XML');
  const value = JSON.parse(root.textContent ?? '') as NativeTeaching;
  if (
    !value || typeof value !== 'object' || value.schemaVersion !== 'deckwise-native-teaching/1' ||
    typeof value.documentId !== 'string' || !/^[\w.-]{1,128}$/.test(value.documentId) ||
    typeof value.revision !== 'string' || !/^[\w.-]{1,128}$/.test(value.revision) ||
    !Array.isArray(value.pages) ||
    value.pages.length > 5000 ||
    !Array.isArray(value.media) ||
    value.media.length > 1000 ||
    !value.pages.every(
      (p) =>
        p && typeof p === 'object' && typeof p.pageId === 'string' &&
        typeof p.locked === 'boolean' &&
        Array.isArray(p.actions) &&
        p.actions.length <= 1000 &&
        p.actions.every((a) => a && typeof a === 'object' && typeof a.id === 'string' && typeof a.type === 'string') && new Set(p.actions.map(a => a.id)).size === p.actions.length,
    ) ||
    new Set(value.pages.map((p) => p.pageId)).size !== value.pages.length
  )
    throw new Error('Unsupported native teaching data');
  validateMedia(value);
  if (decoded.size >= 2) decoded.clear();
  decoded.set(xml, value);
  return value;
}
function validateMedia(value: NativeTeaching) {
  let total = 0;
  const ids = new Set<string>();
  for (const media of value.media) {
    if (!media || !/^[\w.-]{1,128}$/.test(media.id) || ids.has(media.id) || !/^[a-f0-9]{64}$/.test(media.sha256) ||
        !['audio/mpeg', 'audio/wav', 'audio/ogg', 'audio/mp4', 'video/mp4'].includes(media.mimeType) ||
        !Number.isSafeInteger(media.byteLength) || media.byteLength < 1 || media.byteLength > 50 * 1024 * 1024 ||
        !Number.isSafeInteger(media.durationMs) || media.durationMs < 1 || media.durationMs > 86400000 ||
        (media.data !== undefined && (typeof media.data !== 'string' || media.data.length !== 4 * Math.ceil(media.byteLength / 3))) ||
        (total += media.byteLength) > 60 * 1024 * 1024) throw new Error('Invalid native teaching media');
    ids.add(media.id);
  }
}

function nativeTeachingParts(): Part[] {
  let active: ReturnType<typeof session>;
  try { active = session(); } catch (error) {
    if (error instanceof Error && error.message === 'Native teaching API unavailable') return [];
    throw error;
  }
  const parts = active.parts.GetByNamespace(NATIVE_TEACHING_NAMESPACE);
  if (parts.length > 1) throw new Error('Duplicate native teaching parts');
  return parts;
}
function readNativeTeachingCarrier(): NativeTeaching | undefined {
  // GetXml serializes the SDK tree; never serialize embedded audio just to locate metadata.
  const values = nativeTeachingParts()
    .map(decode)
    .filter((v): v is NativeTeaching => !!v);
  if (values.length > 1) throw new Error('Duplicate native teaching parts');
  return values[0];
}

// This guard runs per object and selection. Content validation belongs to the subsequent read.
export function hasNativeTeaching(): boolean { return nativeTeachingParts().length > 0; }
export function readNativeTeaching(): NativeTeaching | undefined {
  const carrier = readNativeTeachingCarrier();
  return carrier ? projectNativeTeaching(carrier) : undefined;
}

/** A single native action replaces the carrier; native keyboard undo sees the same change. */
export function replaceNativeTeachingPage(input: {
  documentId: string;
  expectedRevision: string;
  pageId: string;
  actions: Action[];
}) {
  if (getReadonlyMode()) throw new Error('Teaching is read-only');
  const current = readNativeTeaching();
  if (!current || current.documentId !== input.documentId || current.revision !== input.expectedRevision)
    throw new Error('Teaching revision changed');
  const page = current.pages.find((p) => p.pageId === input.pageId);
  if (!page || page.locked) throw new Error('Teaching page unavailable or locked');
  if (
    !Array.isArray(input.actions) ||
    input.actions.length > 1000 ||
    !input.actions.every((a) => a && typeof a.id === 'string' && typeof a.type === 'string') ||
    new Set(input.actions.map((a) => a.id)).size !== input.actions.length
  )
    throw new Error('Invalid teaching actions');
  const next = { ...current, pages: structuredClone(current.pages), media: [...current.media] };
  const actions = structuredClone(input.actions);
  for (const action of actions) {
    const previous = page.actions.find((a) => a.id === action.id);
    if (action.type === 'speech' && previous?.text !== action.text) {
      delete action.audioAssetId;
      delete action.durationMs;
      if (Array.isArray(action.visualCues)) for (const cue of action.visualCues) {
        if (cue?.trigger?.kind === 'text') cue.trigger.stale = true;
      }
    }
  }
  const changed = next.pages.find((p) => p.pageId === input.pageId)!;
  changed.actions = actions;
  changed.reviewState = 'review';
  return commitNativeTeaching(next);
}

function nativeAnimationSignature(pageNumber: number, stepIndex: number): string {
  const api = getEditorApi() as any;
  const slide = api?.WordControl?.m_oLogicDocument?.Slides?.[pageNumber - 1];
  const effect = slide?.timing?.getAllAnimEffects?.()[stepIndex];
  if (!effect) throw new Error('Native teaching animation is missing; bind the current animations again');
  const target = slideShapes(slide).find(shape => (shape as any).Id === effect.getObjectId?.());
  const object = target && live?.objects.get(target);
  if (!object) throw new Error('Native teaching animation target is unavailable');
  return JSON.stringify([object.logicalId, effect.cTn?.presetClass, effect.cTn?.presetID, effect.cTn?.presetSubtype,
    effect.cTn?.nodeType, effect.asc_getDuration?.(), effect.asc_getDelay?.(), effect.asc_getRepeatCount?.(), effect.asc_getRewind?.()]);
}

/** Explicit step links never silently move to a different native effect after reorder. */
export function assertNativeTeachingAnimations(pageNumber: number): void {
  const teaching = readNativeTeaching();
  if (!teaching) return;
  const page = teaching.pages[pageNumber - 1];
  const bindings = teaching.nativeAnimationBindings as Record<string, string> | undefined;
  const actions = page?.actions.filter(action => action.type === 'native_animation') ?? [];
  for (const action of actions) if (!bindings || bindings[action.id] !== nativeAnimationSignature(pageNumber, Number(action.stepIndex)))
    throw new Error('PPT 动画顺序或属性已变化，请移除动画绑定后重新关联');
}

function commitNativeTeaching(next: NativeTeaching, revise = true, mediaXml?: string): NativeTeaching {
  validateMedia(next);
  const bindings = (next.nativeAnimationBindings ?? {}) as Record<string, string>;
  const active: Record<string, string> = {};
  next.pages.forEach((page, index) => {
    for (const action of page.actions) if (action.type === 'native_animation')
      active[action.id] = bindings[action.id] ?? nativeAnimationSignature(index + 1, Number(action.stepIndex));
  });
  next.nativeAnimationBindings = active;
  // The saved action union retains one source anchor; live views and fixed snapshots
  // project it onto their independently verified current package hash.
  const motions = next.pages.flatMap(page => page.actions).filter(action => action.type === 'native_animation');
  const anchor = motions[0]?.sourceSha256;
  for (const motion of motions) motion.sourceSha256 = anchor;
  if (revise) next.revision = crypto.randomUUID();
  const xml = document.implementation.createDocument(NATIVE_TEACHING_NAMESPACE, 'teaching');
  xml.documentElement.textContent = JSON.stringify(next);
  const encoded = new XMLSerializer().serializeToString(xml);
  if (encoded.length > 80 * 1024 * 1024) throw new Error('Teaching exceeds package budget');
  const { parts, logic, description, history } = session();
  const previous = parts.GetByNamespace(NATIVE_TEACHING_NAMESPACE).filter((p) => decode(p));
  const before = previous.map(part => part.GetXml());
  const added: Part[] = [];
  // Export identity is derived metadata, so saving must not add an extra Undo step.
  if (revise) logic.StartAction(description); else history?.TurnOff();
  try {
    if (mediaXml) added.push(parts.Add(mediaXml));
    added.push(parts.Add(encoded));
    previous.forEach((p) => p.Delete());
  } catch (cause) {
    // A failed compound write must retain the old carrier and cannot strand audio.
    try {
      for (const part of added) if (parts.GetAll().includes(part)) part.Delete();
      const remaining = parts.GetAll().map(part => part.GetXml());
      for (const xml of before) if (!remaining.includes(xml)) parts.Add(xml);
    } catch (recovery) {
      throw new AggregateError([cause, recovery], 'Native teaching write and rollback failed');
    }
    throw cause;
  } finally {
    if (revise) logic.FinalizeAction(); else history?.TurnOn();
  }
  if (revise) window.dispatchEvent(new Event('deckwise-native-teaching-changed'));
  return projectNativeTeaching(next);
}

export async function nativeTeachingAudio(assetId: string): Promise<Blob> {
  const teaching = readNativeTeaching();
  const media = teaching?.media.filter((m) => m.id === assetId);
  if (!media || media.length !== 1) throw new Error('Native teaching audio missing');
  let item = media[0];
  if (item.data === undefined) {
    const matches = session().parts.GetAll().flatMap(part => {
      const xml = part.GetXml();
      if (!xml.slice(0, 512).includes(NATIVE_MEDIA_NAMESPACE)) return [];
      if (xml.length > 70 * 1024 * 1024 || /<!DOCTYPE|<!ENTITY/i.test(xml)) throw new Error('Invalid native media XML');
      const root = new DOMParser().parseFromString(xml, 'application/xml').documentElement;
      if (root.namespaceURI !== NATIVE_MEDIA_NAMESPACE || root.localName !== 'media' || root.children.length) throw new Error('Invalid native media XML');
      const asset = JSON.parse(root.textContent ?? '');
      return asset.id === assetId ? [asset] : [];
    });
    if (matches.length !== 1 || matches[0].sha256 !== item.sha256) throw new Error('Native media part missing or duplicated');
    item = { ...item, data: matches[0].data };
  }
  if (
    !['audio/mpeg', 'audio/wav', 'audio/ogg', 'audio/mp4', 'video/mp4'].includes(item.mimeType) ||
    !/^[a-f0-9]{64}$/.test(item.sha256) ||
    !Number.isSafeInteger(item.byteLength) ||
    item.byteLength < 1 ||
    item.byteLength > 50 * 1024 * 1024 ||
    typeof item.data !== 'string' ||
    item.data.length > 70 * 1024 * 1024
  )
    throw new Error('Invalid native teaching audio');
  const bytes = Uint8Array.from(atob(item.data!), (c) => c.charCodeAt(0));
  const digest = Array.from(new Uint8Array(await crypto.subtle.digest('SHA-256', bytes)), (b) =>
    b.toString(16).padStart(2, '0'),
  ).join('');
  if (bytes.length !== item.byteLength || digest !== item.sha256)
    throw new Error('Native teaching audio digest mismatch');
  return new Blob([bytes], { type: item.mimeType });
}

/** Adopt only bytes verified against the exact text revision, in the same native history as the action. */
export async function adoptNativeTeachingAudio(input: {
  documentId: string; expectedRevision: string; pageId: string; actionId: string;
  assetId: string; sha256: string; mimeType: string; bytes: ArrayBuffer; durationMs: number;
  kind?: 'music' | 'video';
}) {
  if (getReadonlyMode()) throw new Error('Teaching is read-only');
  if (Object.prototype.toString.call(input.bytes) !== '[object ArrayBuffer]' || input.bytes.byteLength < 1 || input.bytes.byteLength > 50 * 1024 * 1024 ||
    !['audio/mpeg', 'audio/wav', 'audio/ogg', 'audio/mp4', 'video/mp4'].includes(input.mimeType) ||
    !Number.isSafeInteger(input.durationMs) || input.durationMs < 1 || input.durationMs > 86400000 ||
    !/^[\w.-]{1,128}$/.test(input.assetId)) throw new Error('Invalid adopted native media');
  const digest = Array.from(new Uint8Array(await crypto.subtle.digest('SHA-256', input.bytes)), b => b.toString(16).padStart(2, '0')).join('');
  if (digest !== input.sha256) throw new Error('Native media digest mismatch');
  const current = readNativeTeaching();
  if (!current || current.documentId !== input.documentId || current.revision !== input.expectedRevision || getReadonlyMode()) throw new Error('Teaching revision changed');
  const next = { ...current, pages: structuredClone(current.pages), media: [...current.media] };
  const page = next.pages.find(p => p.pageId === input.pageId);
  let action = page?.actions.find(a => a.id === input.actionId);
  if (page && !page.locked && input.kind && (!action || action.type === (input.kind === 'video' ? 'video_slot' : 'music_slot'))) {
    if (!/^[\w.-]{1,128}$/.test(input.actionId) || page.actions.length >= 1000) throw new Error('Invalid native media action');
    const replacement: Action = { id: input.actionId, type: input.kind === 'video' ? 'video_clip' : 'audio_clip' };
    if (action) page.actions[page.actions.indexOf(action)] = replacement; else page.actions.push(replacement);
    action = replacement;
  }
  if (!page || page.locked || !action || !['speech', 'audio_clip', 'video_clip'].includes(action.type)) throw new Error('Native media target unavailable');
  if ((action.type === 'video_clip') !== input.mimeType.startsWith('video/')) throw new Error('Native media type mismatch');
  let binary = '';
  const bytes = new Uint8Array(input.bytes);
  for (let offset = 0; offset < bytes.length; offset += 8192) binary += String.fromCharCode(...bytes.subarray(offset, offset + 8192));
  const data = btoa(binary);
  const asset = { id: input.assetId, sha256: digest, mimeType: input.mimeType, byteLength: bytes.length, durationMs: input.durationMs };
  const existing = next.media.find(m => m.id === asset.id);
  if (existing && existing.sha256 !== digest) throw new Error('Native media identity conflict');
  next.media = [...next.media.filter(m => m.id !== asset.id), existing?.data ? { ...asset, data: existing.data } : asset];
  action[action.type === 'video_clip' ? 'videoAssetId' : 'audioAssetId'] = asset.id;
  action.durationMs = input.durationMs;
  if (action.type === 'speech' && action.attention !== undefined) {
    const intent = action.attention as { type?: string; objectIds?: unknown[]; animation?: unknown };
    const ids = intent?.objectIds;
    const pageNumber = current.pages.findIndex(p => p.pageId === input.pageId) + 1;
    const objects = nativeTeachingObjects(pageNumber);
    if (!intent || !['highlight', 'spotlight'].includes(String(intent.type)) || !Array.isArray(ids) ||
        ids.length < 1 || ids.length > 8 || new Set(ids).size !== ids.length ||
        ids.some(id => typeof id !== 'string' || !objects.some(o => o.objectId === id && o.effectTargetEligible)) ||
        Object.keys(intent).some(key => !['type', 'objectIds', 'animation'].includes(key)) ||
        intent.animation !== undefined && !['fadeIn', 'pulse', 'zoomIn', 'slideInLeft', 'slideInRight', 'gentleRotate'].includes(String(intent.animation)) ||
        input.durationMs > 3600000) throw new Error('Native speech attention target or timing changed');
    const duration = Math.min(input.durationMs, 600);
    const reveal = ['fadeIn', 'zoomIn', 'slideInLeft', 'slideInRight'].includes(String(intent.animation));
    action.visualCues = [
      ...(intent.animation ? ids.map((id, index) => ({ id: `motion-${action!.id}-${index}`, atMs: 0, endMs: duration,
        trigger: { version: 1, kind: 'start' }, effect: { type: 'element_animation', elementId: id, effect: intent.animation } })) : []),
      { id: `focus-${action.id}`, atMs: reveal ? Math.min(duration, input.durationMs - 1) : 0, endMs: input.durationMs,
        effect: intent.type === 'highlight' ? { type: 'highlight', elementIds: ids } : { type: 'spotlight', elementId: ids[0], elementIds: ids.slice(1) } },
    ];
    delete action.attention;
  }
  const part = document.implementation.createDocument(NATIVE_MEDIA_NAMESPACE, 'media');
  part.documentElement.textContent = JSON.stringify({ id: asset.id, sha256: digest, data });
  return commitNativeTeaching(next, true, existing ? undefined : new XMLSerializer().serializeToString(part));
}

export function setNativeTeachingLock(input: { documentId: string; expectedRevision: string; pageId: string; locked: boolean }) {
  const current = readNativeTeaching();
  if (getReadonlyMode() || !current || current.documentId !== input.documentId || current.revision !== input.expectedRevision || typeof input.locked !== 'boolean') throw new Error('Teaching revision changed or is read-only');
  const next = { ...current, pages: structuredClone(current.pages), media: [...current.media] }, page = next.pages.find(p => p.pageId === input.pageId);
  if (!page) throw new Error('Native teaching page unavailable');
  page.locked = input.locked;
  return commitNativeTeaching(next);
}


type LiveShape = { spTree?: LiveShape[]; getUniNvProps?(): { cNvPr?: { id?: number | string; name?: string } }; group?: unknown };
type LiveSlide = { cSld?: { spTree?: LiveShape[] }; createDuplicate?: (...args: unknown[]) => LiveSlide };
type ObjectIdentity = { logicalId: string; sourceShapeId?: string; sourceObjectName?: string; type?: string };
function slideShapes(slide: LiveSlide): LiveShape[] {
  const visit = (shapes: LiveShape[]): LiveShape[] => shapes.flatMap(shape => [shape, ...visit(shape.spTree ?? [])]);
  return visit(slide.cSld?.spTree ?? []);
}
let live: { api: unknown; documentId: string; pages: WeakMap<LiveSlide, string>; objects: WeakMap<LiveShape, ObjectIdentity>; teaching: WeakMap<LiveSlide, Page> } | undefined;

/** Native object references keep identities stable through reorder, rename, undo and duplication. */
function projectNativeTeaching(source: NativeTeaching): NativeTeaching {
  const api = getEditorApi() as unknown as { WordControl?: { m_oLogicDocument?: { Slides?: LiveSlide[] } } };
  const slides = api?.WordControl?.m_oLogicDocument?.Slides;
  if (!Array.isArray(slides)) return structuredClone(source);
  const identity = source.identity as { pages?: Array<{ pageId?: string; objects?: ObjectIdentity[] }> } | undefined;
  if (!live || live.api !== api || live.documentId !== source.documentId) {
    live = { api, documentId: source.documentId, pages: new WeakMap(), objects: new WeakMap(), teaching: new WeakMap() };
    slides.forEach((slide, index) => {
      const page = source.pages[index];
      if (page) live!.pages.set(slide, page.pageId);
      for (const shape of slideShapes(slide)) {
        const id = importedSlideShapeId(shape);
        const matches = identity?.pages?.[index]?.objects?.filter(o => id !== undefined && o.sourceShapeId === id) ?? [];
        if (matches.length === 1) live!.objects.set(shape, structuredClone(matches[0]));
      }
    });
  }
  const next = { ...source, pages: [] } as NativeTeaching;
  const identities: Array<{ pageId: string; objects: ObjectIdentity[] }> = [];
  slides.forEach(slide => {
    let pageId = live!.pages.get(slide);
    if (!pageId) { pageId = crypto.randomUUID(); live!.pages.set(slide, pageId); }
    const page: Page = structuredClone(source.pages.find(p => p.pageId === pageId) ?? live!.teaching.get(slide) ?? { pageId, locked: false, reviewState: 'review', actions: [] });
    live!.teaching.set(slide, page);
    next.pages.push(page);
    const objects = (slideShapes(slide)).map(shape => {
      let object = live!.objects.get(shape);
      if (!object) { object = { logicalId: crypto.randomUUID(), type: 'shape' }; live!.objects.set(shape, object); }
      return { ...object, sourceObjectName: shape.getUniNvProps?.()?.cNvPr?.name ?? object.logicalId };
    });
    identities.push({ pageId, objects });
    trackTeachingSlideCopy(slide, duplicate => {
      if (!live || live.api !== api) return;
      const copy = duplicate as LiveSlide, from = live.teaching.get(slide);
      if (!from) return;
      const copied = structuredClone(from), remap = new Map<string, string>();
      copied.pageId = crypto.randomUUID();
      copied.locked = false;
      copied.reviewState = 'review';
      (slideShapes(slide)).forEach((shape, index) => {
        const previous = live!.objects.get(shape), destination = slideShapes(copy)[index];
        if (previous && destination) {
          const logicalId = crypto.randomUUID(); remap.set(previous.logicalId, logicalId);
          live!.objects.set(destination, { ...previous, logicalId, sourceShapeId: undefined });
        }
      });
      const rebind = (value: unknown): unknown => {
        if (Array.isArray(value)) return value.map(rebind);
        if (value && typeof value === 'object') return Object.fromEntries(Object.entries(value).map(([key, item]) =>
          [key, key === 'id' ? crypto.randomUUID() : rebind(item)]));
        return typeof value === 'string' ? remap.get(value) ?? value : value;
      };
      copied.actions = copied.actions.map(action => rebind(action) as Action);
      live.pages.set(copy, copied.pageId); live.teaching.set(copy, copied);
    });
  });
  next.identity = { ...identity, pages: identities };
  // Media bytes stay shared and immutable; cloning them on every selection would block the UI.
  return next;
}

export function nativeTeachingObjects(pageNumber: number) {
  const teaching = readNativeTeaching();
  const api = getEditorApi() as any;
  const slide = api?.WordControl?.m_oLogicDocument?.Slides?.[pageNumber - 1];
  const tree: LiveShape[] = slide ? slideShapes(slide) : [];
  if (!teaching || !live) return [];
  return tree.map(shape => {
    const object = live!.objects.get(shape)!;
    return { objectId: object.logicalId, name: shape.getUniNvProps?.()?.cNvPr?.name ?? object.logicalId,
      sourceShapeId: object.sourceShapeId, sourceObjectName: object.sourceObjectName, status: 'mapped', effectTargetEligible: true,
      nativeShape: shape };
  });
}

export function nativeTeachingTarget(objectId: string, tree?: LiveShape[]): unknown {
  const api = getEditorApi() as any;
  if (!live || live.api !== api) return;
  const shapes = tree ? slideShapes({ cSld: { spTree: tree } }) :
    (api?.WordControl?.m_oLogicDocument?.Slides ?? []).flatMap(slideShapes);
  const matches = shapes.filter((shape: LiveShape) => live!.objects.get(shape)?.logicalId === objectId);
  return matches.length === 1 ? matches[0] : undefined;
}

/** Export current slide order and package shape IDs together with the teaching carrier. */
export function prepareNativeTeachingSave(): void {
  const teaching = readNativeTeaching();
  if (!teaching) return;
  const api = getEditorApi() as any;
  const identity = teaching.identity as { pages: Array<{ objects: ObjectIdentity[] }> };
  api.WordControl.m_oLogicDocument.Slides.forEach((slide: LiveSlide, index: number) => {
    identity.pages[index].objects = (slideShapes(slide)).map(shape => ({
      ...live!.objects.get(shape)!, sourceShapeId: String(shape.getUniNvProps?.()?.cNvPr?.id ?? ''),
      sourceObjectName: shape.getUniNvProps?.()?.cNvPr?.name ?? '',
    }));
  });
  const current = readNativeTeachingCarrier();
  if (JSON.stringify(current) !== JSON.stringify(teaching)) commitNativeTeaching(teaching, false);
}


/** New-format activation reads notes already loaded in Office; it never opens a parser job or legacy database. */
export function initializeNativeTeaching(): NativeTeaching {
  const existing = readNativeTeaching();
  if (existing) return existing;
  if (getReadonlyMode()) throw new Error('Teaching is read-only');
  const api = getEditorApi() as any;
  const slides = api?.WordControl?.m_oLogicDocument?.Slides;
  if (!Array.isArray(slides) || !slides.length || slides.length > 100) throw new Error('Native teaching requires 1–100 slides');
  const teaching: NativeTeaching = { schemaVersion: 'deckwise-native-teaching/1', documentId: crypto.randomUUID(),
    revision: crypto.randomUUID(), pages: slides.map(slide => {
      const text = (slide.notes?.cSld?.spTree ?? []).map(readPresentationShapeText).filter(Boolean).join('\n').trim();
      return { pageId: crypto.randomUUID(), locked: false, reviewState: 'review', actions: text ? [{ id: crypto.randomUUID(), type: 'speech', text, source: 'notes' }] : [] };
    }), media: [], identity: { pages: [] } };
  return commitNativeTeaching(teaching);
}


export function confirmNativeTeachingPage(input: { documentId: string; expectedRevision: string; pageId: string }) {
  const current = readNativeTeaching();
  if (getReadonlyMode() || !current || current.documentId !== input.documentId || current.revision !== input.expectedRevision) throw new Error('Teaching revision changed or is read-only');
  const next = { ...current, pages: structuredClone(current.pages), media: [...current.media] }, page = next.pages.find(p => p.pageId === input.pageId);
  if (!page) throw new Error('Native teaching page unavailable');
  page.reviewState = 'ready';
  return commitNativeTeaching(next);
}
