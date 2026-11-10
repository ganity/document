import { readPresentationShapeText } from './agent-plugin/slide-text-read';
import { getReadonlyMode } from './onlyoffice/readonly';
import { getEditorApi } from './agent-plugin/editor-bridge';
import { deckwisePageLineage } from './deckwise-page-lineage';
import { assertNativeTeachingAnimations, readNativeTeaching, replaceNativeTeachingPage, nativeTeachingAudio, setNativeTeachingLock, nativeTeachingObjects, adoptNativeTeachingAudio, confirmNativeTeachingPage } from './deckwise-native-teaching';

/** Return undefined only for a legacy document. Native documents never write through the legacy Host. */
export async function routeNativeTeaching(operation: string, args: Record<string, unknown>, pageNumber: number): Promise<{ value: unknown } | undefined> {
  const teaching = readNativeTeaching();
  if (!teaching) return;
  const sourceSha256 = deckwisePageLineage.context()?.sourceSha256 ?? '';
  for (const current of teaching.pages) for (const action of current.actions) if (action.type === 'native_animation') action.sourceSha256 = sourceSha256;
  const page = teaching.pages[pageNumber - 1];
  if (!page) throw new Error('Native teaching page is unavailable');
  const view = () => {
    return { nativeTeaching: true, viewId: teaching.documentId + '.' + page.pageId, pageId: page.pageId, pageNumber,
      project: { schemaVersion: 'deckwise-courseware/1', projectId: teaching.documentId, taskType: 'plain', title: 'Native teaching', teachingRevision: teaching.revision,
        source: { revision: teaching.revision, sha256: deckwisePageLineage.context()?.sourceSha256 ?? '' }, pages: teaching.pages },
      teachingObjects: nativeTeachingObjects(pageNumber).map(({ nativeShape, ...object }) => ({ ...object, aliases: [] })) };
  };
  if (operation === 'open') return { value: view() };
  if (operation === 'notes-context') {
    const api = getEditorApi() as any;
    const slide = api?.WordControl?.m_oLogicDocument?.Slides?.[pageNumber - 1];
    return { value: { view: view(), text: (slide?.notes?.cSld?.spTree ?? []).map(readPresentationShapeText).filter(Boolean).join('\n').trim() } };
  }
  if (operation === 'narration-context') {
    const api = getEditorApi() as any;
    const slide = api?.WordControl?.m_oLogicDocument?.Slides?.[pageNumber - 1];
    return { value: { view: view(), text: (slide?.cSld?.spTree ?? []).map(readPresentationShapeText).filter(Boolean).join('\n') } };
  }
  if (['history', 'draft:get', 'draft:set'].includes(operation) && args.viewId !== teaching.documentId + '.' + page.pageId) throw new Error('Native teaching view changed');
  if (operation === 'history') return { value: { entries: [], skippedEntryCount: 0 } };
  if (operation !== 'open' && (args.viewId !== teaching.documentId + '.' + page.pageId || args.expectedTeachingRevision !== teaching.revision) && !['draft:get', 'draft:set', 'undo', 'redo'].includes(operation)) throw new Error('Native teaching view changed');
  if (operation === 'validate-animations') { assertNativeTeachingAnimations(pageNumber); return { value: { valid: true } }; }
  if (operation === 'adopt-audio') {
    await adoptNativeTeachingAudio({ ...(args.command as Parameters<typeof adoptNativeTeachingAudio>[0]), documentId: teaching.documentId, expectedRevision: teaching.revision, pageId: page.pageId });
    return routeNativeTeaching('open', {}, pageNumber);
  }
  if (operation === 'confirm') {
    confirmNativeTeachingPage({ documentId: teaching.documentId, expectedRevision: teaching.revision, pageId: page.pageId });
    return routeNativeTeaching('open', {}, pageNumber);
  }
  if (operation === 'draft:set' && args.draft != null) throw new Error('Native teaching edits require write with the current teaching revision');
  if (operation === 'draft:get' || operation === 'draft:set') return { value: { draft: null } };
  if (operation === 'lock') {
    setNativeTeachingLock({ documentId: teaching.documentId, expectedRevision: String(args.expectedTeachingRevision), pageId: page.pageId, locked: args.locked as boolean });
    return routeNativeTeaching('open', {}, pageNumber);
  }
  if (operation === 'undo' || operation === 'redo') {
    nativeTeachingHistory(operation === 'undo');
    return routeNativeTeaching('open', {}, pageNumber);
  }
  if (operation === 'write') {
    if (args.viewId !== teaching.documentId + '.' + page.pageId) throw new Error('Native teaching view changed');
    const next = replaceNativeTeachingPage({ documentId: teaching.documentId, expectedRevision: String(args.expectedTeachingRevision), pageId: page.pageId,
      actions: args.actions as Parameters<typeof replaceNativeTeachingPage>[0]['actions'] });
    return routeNativeTeaching('open', {}, next.pages.findIndex(p => p.pageId === page.pageId) + 1);
  }
  if (operation === 'tts:preview' || operation === 'media:preview') {
    const action = page.actions.find(a => a.id === args.actionId);
    const id = action?.audioAssetId ?? action?.videoAssetId;
    if (typeof id !== 'string') throw new Error('Native teaching action has no adopted media');
    const blob = await nativeTeachingAudio(id);
    return { value: { bytes: await blob.arrayBuffer(), mimeType: blob.type } };
  }
  if (operation === 'media:adopt') {
    const asset = teaching.media.find(m => m.id === args.intentId);
    if (!asset) return { value: { nativeService: true } };
    if (asset.mimeType.startsWith('video/') !== (args.kind === 'video') || typeof args.actionId !== 'string' || !['music', 'video'].includes(String(args.kind))) throw new Error('Embedded media unavailable');
    replaceNativeTeachingPage({ documentId: teaching.documentId, expectedRevision: teaching.revision, pageId: page.pageId,
      actions: [...page.actions, { id: args.actionId, type: args.kind === 'video' ? 'video_clip' : 'audio_clip',
        [args.kind === 'video' ? 'videoAssetId' : 'audioAssetId']: asset.id, durationMs: asset.durationMs }] });
    return routeNativeTeaching('open', {}, pageNumber);
  }
  if (operation === 'media:list') return { value: teaching.media.map(m => ({ kind: m.mimeType.startsWith('video/') ? 'video' : 'music', intentId: m.id, title: m.id, model: 'embedded', inPage: page.actions.some(a => a.audioAssetId === m.id || a.videoAssetId === m.id) })) };
  if (operation.startsWith('tts:')) return { value: { nativeService: true } };
  throw new Error('Native teaching operation requires the native media service: ' + operation);
}

export function nativeTeachingHistory(undo: boolean): void {
  if (getReadonlyMode()) throw new Error('Teaching is read-only');
  const api = getEditorApi() as unknown as { Undo(): void; Redo(): void } | null;
  if (!api) throw new Error('Native teaching history unavailable');
  if (undo) api.Undo(); else api.Redo();
}
