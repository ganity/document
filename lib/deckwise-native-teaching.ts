import { isCoursewareEnrichment, type CoursewareEnrichment } from './deckwise-teaching-enrichment';
import { readPresentationShapeText } from './agent-plugin/slide-text-read';
import { trackTeachingSlideCopy } from './onlyoffice/guards/slide-teaching-copy';
import { importedSlideShapeId } from './onlyoffice/guards/slide-source-identity';
import { getEditorApi } from './agent-plugin/editor-bridge';
import { prepareNativeTeachingMotions, type SpeechMotionOwners } from './deckwise-native-teaching-motions';
import { getReadonlyMode } from './onlyoffice/readonly';
import { markDocumentDirty } from './unsaved-guard';

export const NATIVE_MEDIA_NAMESPACE = 'urn:deckwise:native-teaching-media:1';
export const NATIVE_TEACHING_NAMESPACE = 'urn:deckwise:native-teaching:1';
type Action = { id: string; type: string; [key: string]: unknown };
type Page = { pageId: string; locked: boolean; actions: Action[]; reviewState?: 'review' | 'ready' };
export type NativeTeaching = {
  schemaVersion: 'deckwise-native-teaching/1';
  documentId: string;
  revision: string;
  pages: Page[];
  enrichment?: CoursewareEnrichment;
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
    return {
      parts: new scope.AscBuilder.ApiCustomXmlParts(logic),
      logic,
      description,
      history: scope.AscCommon?.History,
    };
  }
  throw new Error('Native teaching API unavailable');
}
function decode(part: Part): NativeTeaching | undefined {
  const xml = part.GetXml();
  if (!xml.includes(NATIVE_TEACHING_NAMESPACE)) return;
  if (decoded.has(xml)) return decoded.get(xml);
  if (xml.length > 80 * 1024 * 1024 || /<!DOCTYPE|<!ENTITY/i.test(xml)) throw new Error('Invalid native teaching XML');
  const root = new DOMParser().parseFromString(xml, 'application/xml').documentElement;
  if (root.namespaceURI !== NATIVE_TEACHING_NAMESPACE || root.localName !== 'teaching' || root.children.length)
    throw new Error('Invalid native teaching XML');
  const value = JSON.parse(root.textContent ?? '') as NativeTeaching;
  if (value?.enrichment !== undefined && !isCoursewareEnrichment(value.enrichment))
    throw new Error('Unsupported or invalid teaching enrichment');
  if (
    !value ||
    typeof value !== 'object' ||
    value.schemaVersion !== 'deckwise-native-teaching/1' ||
    typeof value.documentId !== 'string' ||
    !/^[\w.-]{1,128}$/.test(value.documentId) ||
    typeof value.revision !== 'string' ||
    !/^[\w.-]{1,128}$/.test(value.revision) ||
    !Array.isArray(value.pages) ||
    value.pages.length > 5000 ||
    !Array.isArray(value.media) ||
    value.media.length > 1000 ||
    !value.pages.every(
      (p) =>
        p &&
        typeof p === 'object' &&
        typeof p.pageId === 'string' &&
        typeof p.locked === 'boolean' &&
        Array.isArray(p.actions) &&
        p.actions.length <= 1000 &&
        p.actions.every((a) => a && typeof a === 'object' && typeof a.id === 'string' && typeof a.type === 'string') &&
        new Set(p.actions.map((a) => a.id)).size === p.actions.length,
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
    if (
      !media ||
      !/^[\w.-]{1,128}$/.test(media.id) ||
      ids.has(media.id) ||
      !/^[a-f0-9]{64}$/.test(media.sha256) ||
      !['audio/mpeg', 'audio/wav', 'audio/ogg', 'audio/mp4', 'video/mp4'].includes(media.mimeType) ||
      !Number.isSafeInteger(media.byteLength) ||
      media.byteLength < 1 ||
      media.byteLength > 50 * 1024 * 1024 ||
      !Number.isSafeInteger(media.durationMs) ||
      media.durationMs < 1 ||
      media.durationMs > 86400000 ||
      (media.data !== undefined &&
        (typeof media.data !== 'string' || media.data.length !== 4 * Math.ceil(media.byteLength / 3))) ||
      (total += media.byteLength) > 60 * 1024 * 1024
    )
      throw new Error('Invalid native teaching media');
    ids.add(media.id);
  }
}

function nativeTeachingParts(): Part[] {
  let active: ReturnType<typeof session>;
  try {
    active = session();
  } catch (error) {
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
export function hasNativeTeaching(): boolean {
  return nativeTeachingParts().length > 0;
}
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
  enrichment?: CoursewareEnrichment;
  semantics?: Array<{ objectId: string; role: string; aliases: string[]; text: string; teachingTarget: boolean }>;
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
  const next: NativeTeaching = { ...current, pages: structuredClone(current.pages), media: [...current.media] };
  if (input.enrichment !== undefined) {
    if (!isCoursewareEnrichment(input.enrichment)) throw new Error('Unsupported or invalid teaching enrichment');
    if (
      current.enrichment &&
      JSON.stringify(input.enrichment.originalPages) !== JSON.stringify(current.enrichment.originalPages)
    )
      throw new Error('Original page identity cannot be replaced');
    next.enrichment = structuredClone(input.enrichment);
  }
  const actions = structuredClone(input.actions);
  for (const action of actions) {
    const previous = page.actions.find((a) => a.id === action.id);
    if (action.type === 'speech' && previous?.text !== action.text) {
      delete action.audioAssetId;
      delete action.durationMs;
      if (Array.isArray(action.visualCues))
        for (const cue of action.visualCues) {
          if (cue?.trigger?.kind === 'text') cue.trigger.stale = true;
        }
    }
  }
  for (const action of actions) {
    const media = current.media.find((m) => m.id === action.audioAssetId && m.mimeType.startsWith('audio/'));
    if (media) resolveNativeSpeechAttention(current, input.pageId, action, media.durationMs);
  }
  const changed = next.pages.find((p) => p.pageId === input.pageId)!;
  if (input.semantics) {
    const ids = new Set(
      nativeTeachingObjects(current.pages.findIndex((p) => p.pageId === input.pageId) + 1).map((o) => o.objectId),
    );
    if (
      input.semantics.length > 1000 ||
      new Set(input.semantics.map((o) => o.objectId)).size !== input.semantics.length ||
      input.semantics.some(
        (o) =>
          !ids.has(o.objectId) ||
          typeof o.role !== 'string' ||
          o.role.length > 1000 ||
          typeof o.text !== 'string' ||
          o.text.length > 4000 ||
          typeof o.teachingTarget !== 'boolean' ||
          !Array.isArray(o.aliases) ||
          o.aliases.length > 50 ||
          o.aliases.some((a) => typeof a !== 'string' || a.length > 500),
      )
    )
      throw new Error('Invalid native object semantics');
    next.objectSemantics = {
      ...((current.objectSemantics ?? {}) as object),
      [input.pageId]: structuredClone(input.semantics),
    };
  }
  changed.actions = actions;
  changed.reviewState = 'review';
  return commitNativeTeaching(next, true, undefined, () =>
    prepareNativeTeachingMotions(current.pages.findIndex((p) => p.pageId === input.pageId) + 1, actions, page.actions, nativeTeachingTarget, pageMotionOwners(next, input.pageId)),
  );
}

/** Add an exact background crop and its effects in the same native transaction. */
export function replaceNativeTeachingRegion(input: {
  documentId: string;
  expectedRevision: string;
  pageId: string;
  actionId: string;
  description: string;
  change: { objectId: string; nativeShape: LiveShape; assertCurrent(): void; apply(): void; restore(): void };
}) {
  const current = readNativeTeaching();
  if (
    getReadonlyMode() ||
    !current ||
    current.documentId !== input.documentId ||
    current.revision !== input.expectedRevision
  )
    throw new Error('Teaching revision changed or is read-only');
  const index = current.pages.findIndex((p) => p.pageId === input.pageId),
    page = current.pages[index];
  const plan = current.generation as
    { pages?: Array<{ pageId: string; teaching?: { pageType?: string; requiredEffects?: string[] } }> } | undefined;
  const planned = plan?.pages?.find((p) => p.pageId === input.pageId)?.teaching;
  if (
    !page ||
    page.locked ||
    planned?.pageType !== 'slide' ||
    !planned.requiredEffects?.includes('animation') ||
    !planned.requiredEffects.includes('spotlight') ||
    current.enrichment?.originalPages.some((p) => p.pageId === input.pageId) ||
    (current.backgroundRegions as Record<string, unknown> | undefined)?.[input.pageId] ||
    typeof input.description !== 'string' ||
    !input.description.trim() ||
    input.description.length > 1000 ||
    !/^[\w.-]{1,128}$/.test(input.change.objectId) ||
    nativeTeachingTarget(input.change.objectId)
  )
    throw new Error('Background region is unavailable, protected or already repaired');
  const next = structuredClone(current),
    changed = next.pages[index];
  const speech = changed.actions.find((a) => a.id === input.actionId && a.type === 'speech');
  const cues = speech?.visualCues as
    | Array<{
        id: string;
        atMs: number;
        endMs: number;
        trigger?: unknown;
        effect: { type: string; effect?: string; elementId?: string; elementIds?: string[] };
      }>
    | undefined;
  const fade = cues?.find((c) => c.effect.type === 'element_animation' && c.effect.effect === 'fadeIn');
  const spot = cues?.find((c) => c.effect.type === 'spotlight');
  if (
    !speech ||
    !speech.audioAssetId ||
    !Number.isSafeInteger(speech.durationMs) ||
    Number(speech.durationMs) <= 0 ||
    speech.attention ||
    !fade ||
    !spot ||
    !cues ||
    cues.length >= 199 ||
    !Number.isInteger(fade.atMs) ||
    !Number.isInteger(fade.endMs) ||
    fade.atMs < 0 ||
    fade.endMs <= fade.atMs ||
    fade.endMs > Number(speech.durationMs) ||
    fade.endMs - fade.atMs > 5000 ||
    !spot.effect.elementId ||
    (spot.effect.elementIds?.length ?? 0) >= 198
  )
    throw new Error('Region repair requires saved audio and existing fade-in/spotlight cues');
  const { objectId, nativeShape } = input.change;
  cues.push({
    ...structuredClone(fade),
    id: `region-${objectId}`,
    effect: { type: 'element_animation', effect: 'fadeIn', elementId: objectId },
  });
  spot.effect.elementIds = [...(spot.effect.elementIds ?? []), objectId];
  changed.reviewState = 'review';
  const semantics = (next.objectSemantics ?? {}) as Record<string, unknown[]>;
  next.objectSemantics = {
    ...semantics,
    [input.pageId]: [
      ...(semantics[input.pageId] ?? []),
      { objectId, role: input.description, aliases: [input.description], text: '', teachingTarget: true },
    ],
  };
  next.backgroundRegions = {
    ...((next.backgroundRegions ?? {}) as object),
    [input.pageId]: { objectId, actionId: input.actionId, description: input.description },
  };
  input.change.assertCurrent();
  const timing = prepareNativeTeachingMotions(index + 1, changed.actions, page.actions, (id) =>
    id === objectId ? nativeShape : nativeTeachingTarget(id), pageMotionOwners(next, input.pageId),
  );
  return commitNativeTeaching(next, true, undefined, () => ({
    apply() {
      input.change.apply();
      live!.objects.set(nativeShape, { logicalId: objectId, sourceObjectName: objectId, type: 'image' });
      timing?.apply();
    },
    restore() {
      try {
        timing?.restore();
      } finally {
        try {
          input.change.restore();
        } finally {
          live?.objects.delete(nativeShape);
        }
      }
    },
  }));
}

/** Insert one independent assessment; visible content, answer data and source binding share one Undo action. */
export function insertNativeTeachingAssessment(input: {
  documentId: string;
  expectedRevision: string;
  afterPageId: string;
  pageId: string;
  actions: Action[];
  enrichment: CoursewareEnrichment;
  quiz?: unknown;
  change: { assertCurrent(): void; apply(): LiveSlide; restore(): void };
}) {
  if (getReadonlyMode()) throw new Error('Teaching is read-only');
  const current = readNativeTeaching();
  if (
    !current ||
    current.documentId !== input.documentId ||
    current.revision !== input.expectedRevision ||
    !current.enrichment ||
    JSON.stringify(current.enrichment.originalPages) !== JSON.stringify(input.enrichment.originalPages)
  )
    throw new Error('Teaching or original page identity changed');
  const after = current.pages.findIndex((p) => p.pageId === input.afterPageId);
  if (
    after < 0 ||
    current.pages[after].locked ||
    current.pages.length >= 100 ||
    current.pages.some((p) => p.pageId === input.pageId) ||
    !/^[\w.-]{1,128}$/.test(input.pageId) ||
    !isCoursewareEnrichment(input.enrichment) ||
    !Array.isArray(input.actions) ||
    input.actions.length > 10 ||
    input.actions.some((a) => !a || typeof a.id !== 'string' || !['speech', 'interaction_checkpoint'].includes(a.type))
  )
    throw new Error('Invalid independent assessment insertion');
  input.change.assertCurrent();
  const next: NativeTeaching = {
    ...current,
    pages: structuredClone(current.pages),
    enrichment: structuredClone(input.enrichment),
  };
  next.pages.splice(after + 1, 0, {
    pageId: input.pageId,
    locked: false,
    reviewState: 'review',
    actions: structuredClone(input.actions),
  });
  if (input.quiz)
    next.assessments = { ...((current.assessments ?? {}) as object), [input.pageId]: structuredClone(input.quiz) };
  return commitNativeTeaching(next, true, undefined, () => ({
    apply() {
      const slide = input.change.apply();
      if (!live) throw new Error('Native page identity is unavailable');
      live.pages.set(slide, input.pageId);
    },
    restore() {
      input.change.restore();
    },
  }));
}

function nativeAnimationSignature(pageNumber: number, stepIndex: number): string {
  const api = getEditorApi() as any;
  const slide = api?.WordControl?.m_oLogicDocument?.Slides?.[pageNumber - 1];
  const effect = slide?.timing?.getAllAnimEffects?.()[stepIndex];
  if (!effect) throw new Error('Native teaching animation is missing; bind the current animations again');
  const target = slideShapes(slide).find((shape) => (shape as any).Id === effect.getObjectId?.());
  const object = target && live?.objects.get(target);
  if (!object) throw new Error('Native teaching animation target is unavailable');
  return JSON.stringify([
    object.logicalId,
    effect.cTn?.presetClass,
    effect.cTn?.presetID,
    effect.cTn?.presetSubtype,
    effect.cTn?.nodeType,
    effect.asc_getDuration?.(),
    effect.asc_getDelay?.(),
    effect.asc_getRepeatCount?.(),
    effect.asc_getRewind?.(),
  ]);
}

/** Explicit step links never silently move to a different native effect after reorder. */
export function assertNativeTeachingAnimations(pageNumber: number): void {
  const teaching = readNativeTeaching();
  if (!teaching) return;
  const page = teaching.pages[pageNumber - 1];
  const bindings = teaching.nativeAnimationBindings as Record<string, string> | undefined;
  const actions = page?.actions.filter((action) => action.type === 'native_animation') ?? [];
  for (const action of actions)
    if (!bindings || bindings[action.id] !== nativeAnimationSignature(pageNumber, Number(action.stepIndex)))
      throw new Error('PPT 动画顺序或属性已变化，请移除动画绑定后重新关联');
}

/** Assessment editor writes share the carrier CAS and native Undo transaction. */
export function replaceNativeAssessments(input: {
  documentId: string;
  expectedRevision: string;
  changedPageIds: string[];
  change?: { apply(): void; restore(): void };
  scenes: Array<{ id: string; title?: string; content: any; actions?: Action[] }>;
}) {
  if (getReadonlyMode()) throw new Error('Teaching is read-only');
  const current = readNativeTeaching();
  if (!current || current.documentId !== input.documentId || current.revision !== input.expectedRevision)
    throw new Error('Teaching revision changed');
  if (input.scenes.length !== current.pages.length || input.scenes.some((s, i) => s.id !== current.pages[i].pageId))
    throw new Error('Assessment page order changed');
  const next = structuredClone(current);
  const assessments = { ...((current.assessments ?? {}) as Record<string, unknown>) };
  input.scenes.forEach((scene, index) => {
    const page = next.pages[index];
    if (!input.changedPageIds.includes(scene.id)) return;
    if (page.locked) throw new Error('Teaching page locked');
    if (scene.title !== undefined && typeof scene.title !== 'string') throw new Error('Invalid assessment title');
    for (const speech of scene.actions ?? [])
      if (speech.type === 'speech') {
        if (typeof speech.text !== 'string' || speech.text.length > 32000 || !speech.id)
          throw new Error('Invalid question speech');
        const old = page.actions.find((a) => a.id === speech.id);
        if (old?.type === 'speech' && old.text !== speech.text) {
          old.text = speech.text;
          delete old.audioAssetId;
          delete old.durationMs;
          delete old.visualCues;
          delete old.attention;
        } else if (!old) page.actions.unshift({ id: speech.id, type: 'speech', text: speech.text });
      }

    if (scene.content.type !== 'quiz') {
      for (const action of scene.actions ?? [])
        if (action.type === 'interaction_checkpoint' && action.checkpointKind === 'chapter_summary') {
          const old = page.actions.find((a) => a.id === action.id && a.type === action.type);
          if (!old) throw new Error('Native checkpoint identity changed');
          const reference = action.referenceAnswer as { summary: string; keyPoints: string[]; label?: string };
          Object.assign(old, {
            question: action.question,
            referenceAnswer: { summary: reference.summary, keyPoints: reference.keyPoints },
            context: action.context,
          });
          next.checkpointDetails = {
            ...((next.checkpointDetails ?? {}) as object),
            [scene.id]: { actionId: action.id, constraintPrompt: action.constraintPrompt, label: reference.label },
          };
        }
      return;
    }
    if (scene.content.questions.length !== 1) throw new Error('Native assessment requires one question per page');
    const q = scene.content.questions[0];
    if (!page.actions.some((a) => a.type === 'interaction_checkpoint' && a.checkpointKind === 'chapter_summary'))
      throw new Error('Native question missing');
    page.actions = page.actions.map((a) =>
      a.type === 'interaction_checkpoint' && a.checkpointKind === 'chapter_summary'
        ? {
            ...a,
            chapterTitle: scene.title ?? a.chapterTitle,
            question: q.question,
            response: {
              kind: q.type,
              options: (q.options ?? []).map((o: any) => o.label),
              answerIndices: (q.options ?? []).flatMap((o: any, i: number) => (q.answer?.includes(o.value) ? [i] : [])),
            },
            referenceAnswer: {
              summary: q.referenceAnswer ?? q.analysis ?? (q.type === 'short_answer' ? q.answer?.join('；') : '') ?? '',
              keyPoints: q.keyPoints ?? (q.question === a.question ? (a.referenceAnswer as { keyPoints?: string[] })?.keyPoints : undefined) ?? [],
            },
          }
        : a,
    );
    const oldQuestion = current.pages[index].actions.find(
      (a) => a.type === 'interaction_checkpoint' && a.checkpointKind === 'chapter_summary',
    );
    const newQuestion = page.actions.find((a) => a.id === oldQuestion?.id);
    const oldResponse = oldQuestion?.response as Record<string, unknown> | undefined;
    const newResponse = newQuestion?.response as Record<string, unknown> | undefined;
    // Object key order can change across JSON readers without changing the question or its audio.
    if (
      JSON.stringify([oldQuestion?.question, oldResponse?.kind, oldResponse?.options, oldResponse?.answerIndices]) !==
      JSON.stringify([newQuestion?.question, newResponse?.kind, newResponse?.options, newResponse?.answerIndices])
    ) {
      const reading = [q.question, ...(q.options ?? []).map((o: any) => `${o.value}. ${o.label}`)].join('\n');
      page.actions = page.actions.map((a) => {
        if (
          a.type !== 'speech' ||
          !(
            a.text === oldQuestion?.question ||
            (typeof a.text === 'string' && a.text.startsWith(String(oldQuestion?.question) + '\n'))
          )
        )
          return a;
        const next: Action = { ...a, text: reading };
        delete next.audioAssetId;
        delete next.durationMs;
        delete next.visualCues;
        delete next.attention;
        return next;
      });
    }
    assessments[scene.id] = structuredClone(scene.content);
    page.reviewState = 'review';
  });
  next.assessments = assessments;
  return commitNativeTeaching(next, true, undefined, input.change ? () => input.change! : undefined);
}

function pageMotionOwners(next: NativeTeaching, pageId: string) {
  const pages = (next.speechMotionOwners ?? {}) as Record<string, SpeechMotionOwners>;
  return { legacyGenerated: !!next.generation, previous: Object.hasOwn(pages, pageId) ? pages[pageId] : undefined,
    save: (owners: SpeechMotionOwners) => { next.speechMotionOwners = {...pages, [pageId]: owners}; } };
}

function commitNativeTeaching(
  next: NativeTeaching,
  revise = true,
  mediaXml?: string,
  prepareTiming?: () => { apply(): void; restore(): void } | undefined,
): NativeTeaching {
  validateMedia(next);
  const bindings = (next.nativeAnimationBindings ?? {}) as Record<string, string>;
  const active: Record<string, string> = {};
  next.pages.forEach((page, index) => {
    for (const action of page.actions)
      if (action.type === 'native_animation')
        active[action.id] = bindings[action.id] ?? nativeAnimationSignature(index + 1, Number(action.stepIndex));
  });
  next.nativeAnimationBindings = active;
  // The saved action union retains one source anchor; live views and fixed snapshots
  // project it onto their independently verified current package hash.
  const motions = next.pages.flatMap((page) => page.actions).filter((action) => action.type === 'native_animation');
  const anchor = motions[0]?.sourceSha256;
  for (const motion of motions) motion.sourceSha256 = anchor;
  if (revise) next.revision = crypto.randomUUID();
  const { parts, logic, description, history } = session();
  const previous = parts.GetByNamespace(NATIVE_TEACHING_NAMESPACE).filter((p) => decode(p));
  const before = previous.map((part) => part.GetXml());
  const added: Part[] = [];
  const timingChange = prepareTiming?.();
  // Export identity is derived metadata, so saving must not add an extra Undo step.
  if (revise) logic.StartAction(description);
  else history?.TurnOff();
  try {
    const xml = document.implementation.createDocument(NATIVE_TEACHING_NAMESPACE, 'teaching');
    xml.documentElement.textContent = JSON.stringify(next);
    const encoded = new XMLSerializer().serializeToString(xml);
    if (encoded.length > 80 * 1024 * 1024) throw new Error('Teaching exceeds package budget');
    if (mediaXml) added.push(parts.Add(mediaXml));
    added.push(parts.Add(encoded));
    previous.forEach((p) => p.Delete());
    timingChange?.apply();
  } catch (cause) {
    // A failed compound write must retain the old carrier and cannot strand audio.
    const recoveryErrors: unknown[] = [];
    try {
      timingChange?.restore();
    } catch (error) {
      recoveryErrors.push(error);
    }
    try {
      for (const part of added) if (parts.GetAll().includes(part)) part.Delete();
      const remaining = parts.GetAll().map((part) => part.GetXml());
      for (const xml of before) if (!remaining.includes(xml)) parts.Add(xml);
    } catch (recovery) {
      recoveryErrors.push(recovery);
    }
    if (recoveryErrors.length)
      throw new AggregateError([cause, ...recoveryErrors], 'Native teaching write and rollback failed');
    throw cause;
  } finally {
    if (revise) logic.FinalizeAction();
    else history?.TurnOn();
  }
  if (revise) {
    markDocumentDirty();
    window.dispatchEvent(new Event('deckwise-native-teaching-changed'));
  }
  return projectNativeTeaching(next);
}

export async function nativeTeachingAudio(assetId: string): Promise<Blob> {
  const teaching = readNativeTeaching();
  const media = teaching?.media.filter((m) => m.id === assetId);
  if (!media || media.length !== 1) throw new Error('Native teaching audio missing');
  let item = media[0];
  if (item.data === undefined) {
    const matches = session()
      .parts.GetAll()
      .flatMap((part) => {
        const xml = part.GetXml();
        if (!xml.slice(0, 512).includes(NATIVE_MEDIA_NAMESPACE)) return [];
        if (xml.length > 70 * 1024 * 1024 || /<!DOCTYPE|<!ENTITY/i.test(xml))
          throw new Error('Invalid native media XML');
        const root = new DOMParser().parseFromString(xml, 'application/xml').documentElement;
        if (root.namespaceURI !== NATIVE_MEDIA_NAMESPACE || root.localName !== 'media' || root.children.length)
          throw new Error('Invalid native media XML');
        const asset = JSON.parse(root.textContent ?? '');
        return asset.id === assetId ? [asset] : [];
      });
    if (matches.length !== 1 || matches[0].sha256 !== item.sha256)
      throw new Error('Native media part missing or duplicated');
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

function resolveNativeSpeechAttention(current: NativeTeaching, pageId: string, action: Action, durationMs: number) {
  if (action.type === 'speech' && action.attention !== undefined) {
    const intent = action.attention as { type?: string; objectIds?: unknown[]; animation?: unknown };
    const ids = intent?.objectIds;
    const pageNumber = current.pages.findIndex((p) => p.pageId === pageId) + 1;
    const objects = nativeTeachingObjects(pageNumber);
    if (
      !intent ||
      !['highlight', 'spotlight'].includes(String(intent.type)) ||
      !Array.isArray(ids) ||
      ids.length < 1 ||
      // Match selection, spotlight geometry and the existing native animation target bound.
      ids.length > 200 ||
      new Set(ids).size !== ids.length ||
      ids.some((id) => typeof id !== 'string' || !objects.some((o) => o.objectId === id && o.effectTargetEligible)) ||
      Object.keys(intent).some((key) => !['type', 'objectIds', 'animation'].includes(key)) ||
      (intent.animation !== undefined &&
        !['fadeIn', 'fadeOut', 'pulse', 'zoomIn', 'slideInLeft', 'slideInRight', 'gentleRotate'].includes(
          String(intent.animation),
        )) ||
      durationMs > 3600000
    )
      throw new Error('Native speech attention target or timing changed');
    const duration = Math.min(durationMs, 600);
    const exit = intent.animation === 'fadeOut';
    const atMs = exit ? durationMs : 0;
    if (atMs + duration > 3600000) throw new Error('Speech animation exceeds the supported timeline');
    const reveal = ['fadeIn', 'zoomIn', 'slideInLeft', 'slideInRight'].includes(String(intent.animation));
    const added = [
      ...(intent.animation
        ? ids.map((id, index) => ({
            id: `motion-${action!.id}-${index}`,
            atMs,
            endMs: atMs + duration,
            trigger: { version: 1, kind: exit ? 'end' : 'start' },
            effect: { type: 'element_animation', elementId: id, effect: intent.animation },
          }))
        : []),
      {
        id: `focus-${action.id}`,
        atMs: reveal ? Math.min(duration, durationMs - 1) : 0,
        endMs: durationMs,
        effect:
          intent.type === 'highlight'
            ? { type: 'highlight', elementIds: ids }
            : { type: 'spotlight', elementId: ids[0], elementIds: ids.slice(1) },
      },
    ];
    const existing = Array.isArray(action.visualCues) ? action.visualCues : [];
    action.visualCues = [...existing.filter((cue) => !added.some((next) => next.id === cue.id)), ...added];
    delete action.attention;
  }
}

/** Adopt only bytes verified against the exact text revision, in the same native history as the action. */
export async function adoptNativeTeachingAudio(input: {
  documentId: string;
  expectedRevision: string;
  pageId: string;
  actionId: string;
  assetId: string;
  sha256: string;
  mimeType: string;
  bytes: ArrayBuffer;
  durationMs: number;
  kind?: 'music' | 'video';
}) {
  if (getReadonlyMode()) throw new Error('Teaching is read-only');
  if (
    Object.prototype.toString.call(input.bytes) !== '[object ArrayBuffer]' ||
    input.bytes.byteLength < 1 ||
    input.bytes.byteLength > 50 * 1024 * 1024 ||
    !['audio/mpeg', 'audio/wav', 'audio/ogg', 'audio/mp4', 'video/mp4'].includes(input.mimeType) ||
    !Number.isSafeInteger(input.durationMs) ||
    input.durationMs < 1 ||
    input.durationMs > 86400000 ||
    !/^[\w.-]{1,128}$/.test(input.assetId)
  )
    throw new Error('Invalid adopted native media');
  const digest = Array.from(new Uint8Array(await crypto.subtle.digest('SHA-256', input.bytes)), (b) =>
    b.toString(16).padStart(2, '0'),
  ).join('');
  if (digest !== input.sha256) throw new Error('Native media digest mismatch');
  const current = readNativeTeaching();
  if (
    !current ||
    current.documentId !== input.documentId ||
    current.revision !== input.expectedRevision ||
    getReadonlyMode()
  )
    throw new Error('Teaching revision changed');
  const next = { ...current, pages: structuredClone(current.pages), media: [...current.media] };
  const page = next.pages.find((p) => p.pageId === input.pageId);
  let action = page?.actions.find((a) => a.id === input.actionId);
  if (
    page &&
    !page.locked &&
    input.kind &&
    (!action || action.type === (input.kind === 'video' ? 'video_slot' : 'music_slot'))
  ) {
    if (!/^[\w.-]{1,128}$/.test(input.actionId) || page.actions.length >= 1000)
      throw new Error('Invalid native media action');
    const replacement: Action = { id: input.actionId, type: input.kind === 'video' ? 'video_clip' : 'audio_clip' };
    if (action) page.actions[page.actions.indexOf(action)] = replacement;
    else page.actions.push(replacement);
    action = replacement;
  }
  if (!page || page.locked || !action || !['speech', 'audio_clip', 'video_clip'].includes(action.type))
    throw new Error('Native media target unavailable');
  if ((action.type === 'video_clip') !== input.mimeType.startsWith('video/'))
    throw new Error('Native media type mismatch');
  let binary = '';
  const bytes = new Uint8Array(input.bytes);
  for (let offset = 0; offset < bytes.length; offset += 8192)
    binary += String.fromCharCode(...bytes.subarray(offset, offset + 8192));
  const data = btoa(binary);
  const asset = {
    id: input.assetId,
    sha256: digest,
    mimeType: input.mimeType,
    byteLength: bytes.length,
    durationMs: input.durationMs,
  };
  const existing = next.media.find((m) => m.id === asset.id);
  if (existing && existing.sha256 !== digest) throw new Error('Native media identity conflict');
  next.media = [
    ...next.media.filter((m) => m.id !== asset.id),
    existing?.data ? { ...asset, data: existing.data } : asset,
  ];
  action[action.type === 'video_clip' ? 'videoAssetId' : 'audioAssetId'] = asset.id;
  action.durationMs = input.durationMs;
  resolveNativeSpeechAttention(current, input.pageId, action, input.durationMs);
  const part = document.implementation.createDocument(NATIVE_MEDIA_NAMESPACE, 'media');
  part.documentElement.textContent = JSON.stringify({ id: asset.id, sha256: digest, data });
  const motions =
    action.type === 'speech'
      ? () =>
          prepareNativeTeachingMotions(
            current.pages.findIndex((p) => p.pageId === input.pageId) + 1,
            page.actions,
            current.pages.find((p) => p.pageId === input.pageId)!.actions,
            nativeTeachingTarget, pageMotionOwners(next, input.pageId),
          )
      : undefined;
  return commitNativeTeaching(next, true, existing ? undefined : new XMLSerializer().serializeToString(part), motions);
}

export function setNativeTeachingLock(input: {
  documentId: string;
  expectedRevision: string;
  pageId: string;
  locked: boolean;
}) {
  const current = readNativeTeaching();
  if (
    getReadonlyMode() ||
    !current ||
    current.documentId !== input.documentId ||
    current.revision !== input.expectedRevision ||
    typeof input.locked !== 'boolean'
  )
    throw new Error('Teaching revision changed or is read-only');
  const next = { ...current, pages: structuredClone(current.pages), media: [...current.media] },
    page = next.pages.find((p) => p.pageId === input.pageId);
  if (!page) throw new Error('Native teaching page unavailable');
  page.locked = input.locked;
  return commitNativeTeaching(next);
}

type LiveShape = {
  spTree?: LiveShape[];
  getUniNvProps?(): { cNvPr?: { id?: number | string; name?: string } };
  group?: unknown;
};
type LiveSlide = { cSld?: { spTree?: LiveShape[] }; createDuplicate?: (...args: unknown[]) => LiveSlide };
type ObjectIdentity = { logicalId: string; sourceShapeId?: string; sourceObjectName?: string; type?: string };
function slideShapes(slide: LiveSlide): LiveShape[] {
  const visit = (shapes: LiveShape[]): LiveShape[] => shapes.flatMap((shape) => [shape, ...visit(shape.spTree ?? [])]);
  return visit(slide.cSld?.spTree ?? []);
}
type CopySeed = { page: Page; assessment?: unknown; semantics?: unknown; details?: unknown };
let live:
  | {
      copies: WeakMap<LiveSlide, CopySeed>;
      api: unknown;
      documentId: string;
      pages: WeakMap<LiveSlide, string>;
      objects: WeakMap<LiveShape, ObjectIdentity>;
      teaching: WeakMap<LiveSlide, Page>;
      assessments: WeakMap<LiveSlide, unknown>;
      semantics: WeakMap<LiveSlide, unknown>;
      checkpointDetails: WeakMap<LiveSlide, unknown>;
    }
  | undefined;

/** Native object references keep identities stable through reorder, rename, undo and duplication. */
function projectNativeTeaching(source: NativeTeaching): NativeTeaching {
  const api = getEditorApi() as unknown as { WordControl?: { m_oLogicDocument?: { Slides?: LiveSlide[] } } };
  const slides = api?.WordControl?.m_oLogicDocument?.Slides;
  if (!Array.isArray(slides)) return structuredClone(source);
  const identity = source.identity as { pages?: Array<{ pageId?: string; objects?: ObjectIdentity[] }> } | undefined;
  if (!live || live.api !== api || live.documentId !== source.documentId) {
    live = {
      copies: new WeakMap(),
      api,
      documentId: source.documentId,
      pages: new WeakMap(),
      objects: new WeakMap(),
      teaching: new WeakMap(),
      assessments: new WeakMap(),
      semantics: new WeakMap(),
      checkpointDetails: new WeakMap(),
    };
    slides.forEach((slide, index) => {
      const page = source.pages[index];
      if (page) live!.pages.set(slide, page.pageId);
      for (const shape of slideShapes(slide)) {
        const id = importedSlideShapeId(shape);
        const matches =
          identity?.pages?.[index]?.objects?.filter((o) => id !== undefined && o.sourceShapeId === id) ?? [];
        if (matches.length === 1) live!.objects.set(shape, structuredClone(matches[0]));
      }
    });
  }
  const next = { ...source, pages: [] } as NativeTeaching;
  const assessments: Record<string, unknown> = {},
    semantics: Record<string, unknown> = {},
    checkpointDetails: Record<string, unknown> = {};
  const identities: Array<{ pageId: string; objects: ObjectIdentity[] }> = [];
  slides.forEach((slide) => {
    let pageId = live!.pages.get(slide);
    if (!pageId) {
      pageId = crypto.randomUUID();
      live!.pages.set(slide, pageId);
    }
    const page: Page = structuredClone(
      source.pages.find((p) => p.pageId === pageId) ??
        live!.copies.get(slide)?.page ?? { pageId, locked: false, reviewState: 'review', actions: [] },
    );
    live!.teaching.set(slide, page);
    next.pages.push(page);
    const carried = source.pages.some((p) => p.pageId === pageId);
    const assessment =
      (source.assessments as Record<string, unknown> | undefined)?.[pageId] ??
      (!carried ? live!.copies.get(slide)?.assessment : undefined);
    const meaning =
      (source.objectSemantics as Record<string, unknown> | undefined)?.[pageId] ??
      (!carried ? live!.copies.get(slide)?.semantics : undefined);
    if (assessment) {
      assessments[pageId] = assessment;
      live!.assessments.set(slide, assessment);
    } else live!.assessments.delete(slide);
    if (meaning) {
      semantics[pageId] = meaning;
      live!.semantics.set(slide, meaning);
    } else live!.semantics.delete(slide);
    const details =
      (source.checkpointDetails as Record<string, unknown> | undefined)?.[pageId] ??
      (!carried ? live!.copies.get(slide)?.details : undefined);
    if (details) {
      checkpointDetails[pageId] = details;
      live!.checkpointDetails.set(slide, details);
    } else live!.checkpointDetails.delete(slide);
    const objects = slideShapes(slide).map((shape) => {
      let object = live!.objects.get(shape);
      if (!object) {
        object = { logicalId: crypto.randomUUID(), type: 'shape' };
        live!.objects.set(shape, object);
      }
      return { ...object, sourceObjectName: shape.getUniNvProps?.()?.cNvPr?.name ?? object.logicalId };
    });
    identities.push({ pageId, objects });
    trackTeachingSlideCopy(slide, (duplicate) => {
      if (!live || live.api !== api) return;
      const copy = duplicate as LiveSlide,
        from = live.teaching.get(slide);
      if (!from) return;
      const copied = structuredClone(from),
        remap = new Map<string, string>();
      copied.pageId = crypto.randomUUID();
      copied.locked = false;
      copied.reviewState = 'review';
      slideShapes(slide).forEach((shape, index) => {
        const previous = live!.objects.get(shape),
          destination = slideShapes(copy)[index];
        if (previous && destination) {
          const logicalId = crypto.randomUUID();
          remap.set(previous.logicalId, logicalId);
          live!.objects.set(destination, { ...previous, logicalId, sourceShapeId: undefined });
        }
      });
      const copiedId = (id: string) => {
        if (!remap.has(id)) remap.set(id, crypto.randomUUID());
        return remap.get(id)!;
      };
      const rebind = (value: unknown): unknown => {
        if (Array.isArray(value)) return value.map(rebind);
        if (value && typeof value === 'object')
          return Object.fromEntries(
            Object.entries(value).map(([key, item]) => [
              key,
              key === 'id' && typeof item === 'string' ? copiedId(item) : rebind(item),
            ]),
          );
        return typeof value === 'string' ? (remap.get(value) ?? value) : value;
      };
      remap.set(from.pageId, copied.pageId);
      for (const action of copied.actions) {
        copiedId(action.id);
        if (typeof action.checkpointId === 'string') copiedId(action.checkpointId);
      }
      copied.actions = copied.actions.map((action) => rebind(action) as Action);
      const assessment = live.assessments.get(slide),
        semantics = live.semantics.get(slide);
      if (assessment) live.assessments.set(copy, rebind(assessment));
      if (semantics) live.semantics.set(copy, rebind(semantics));
      const details = live.checkpointDetails.get(slide);
      if (details) live.checkpointDetails.set(copy, rebind(details));
      live.copies.set(
        copy,
        structuredClone({
          page: copied,
          assessment: assessment && rebind(assessment),
          semantics: semantics && rebind(semantics),
          details: details && rebind(details),
        }),
      );
      live.pages.set(copy, copied.pageId);
      live.teaching.set(copy, copied);
    });
  });
  if (Object.keys(assessments).length) next.assessments = assessments;
  if (Object.keys(semantics).length) next.objectSemantics = semantics;
  if (Object.keys(checkpointDetails).length) next.checkpointDetails = checkpointDetails;
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
  return tree.map((shape) => {
    const object = live!.objects.get(shape)!;
    return {
      objectId: object.logicalId,
      name: shape.getUniNvProps?.()?.cNvPr?.name ?? object.logicalId,
      sourceShapeId: object.sourceShapeId,
      sourceObjectName: object.sourceObjectName,
      status: 'mapped',
      effectTargetEligible: true,
      nativeShape: shape,
    };
  });
}

export function nativeTeachingTarget(objectId: string, tree?: LiveShape[]): unknown {
  const api = getEditorApi() as any;
  if (!live || live.api !== api) return;
  const shapes = tree
    ? slideShapes({ cSld: { spTree: tree } })
    : (api?.WordControl?.m_oLogicDocument?.Slides ?? []).flatMap(slideShapes);
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
    identity.pages[index].objects = slideShapes(slide).map((shape) => ({
      ...live!.objects.get(shape)!,
      sourceShapeId: String(shape.getUniNvProps?.()?.cNvPr?.id ?? ''),
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
  if (!Array.isArray(slides) || !slides.length || slides.length > 100)
    throw new Error('Native teaching requires 1–100 slides');
  const teaching: NativeTeaching = {
    schemaVersion: 'deckwise-native-teaching/1',
    documentId: crypto.randomUUID(),
    revision: crypto.randomUUID(),
    pages: slides.map((slide) => {
      const text = (slide.notes?.cSld?.spTree ?? []).map(readPresentationShapeText).filter(Boolean).join('\n').trim();
      return {
        pageId: crypto.randomUUID(),
        locked: false,
        reviewState: 'review',
        actions: text ? [{ id: crypto.randomUUID(), type: 'speech', text, source: 'notes' }] : [],
      };
    }),
    media: [],
    identity: { pages: [] },
  };
  return commitNativeTeaching(teaching);
}

export function confirmNativeTeachingPage(input: { documentId: string; expectedRevision: string; pageId: string }) {
  const current = readNativeTeaching();
  if (
    getReadonlyMode() ||
    !current ||
    current.documentId !== input.documentId ||
    current.revision !== input.expectedRevision
  )
    throw new Error('Teaching revision changed or is read-only');
  const next = { ...current, pages: structuredClone(current.pages), media: [...current.media] },
    page = next.pages.find((p) => p.pageId === input.pageId);
  if (!page) throw new Error('Native teaching page unavailable');
  page.reviewState = 'ready';
  return commitNativeTeaching(next);
}
