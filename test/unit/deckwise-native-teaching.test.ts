import { webcrypto } from 'node:crypto';
import { afterEach, expect, it, vi } from 'vitest';
import { initializeNativeTeaching, assertNativeTeachingAnimations, hasNativeTeaching, readNativeTeaching, replaceNativeTeachingPage, nativeTeachingAudio, adoptNativeTeachingAudio, NATIVE_MEDIA_NAMESPACE, NATIVE_TEACHING_NAMESPACE, nativeTeachingObjects, nativeTeachingTarget, setNativeTeachingLock, prepareNativeTeachingSave } from '../../lib/deckwise-native-teaching';
import { routeNativeTeaching } from '../../lib/deckwise-native-teaching-router';
import { attentionGeometry, selectedAttentionTargets } from '../../lib/deckwise-teaching-preview';
vi.mock('../../lib/deckwise-page-lineage', () => ({ deckwisePageLineage: { context: () => ({ sourceSha256: 'a'.repeat(64) }) } }));
const state = vi.hoisted(() => ({ api: {}, readonly: false }));
vi.mock('../../lib/agent-plugin/editor-bridge', () => ({ getEditorApi: () => state.api }));
vi.mock('../../lib/onlyoffice/guards/slide-source-identity', () => ({ importedSlideShapeId: (shape: any) => shape.sourceId }));
vi.mock('../../lib/onlyoffice/readonly', () => ({ getReadonlyMode: () => state.readonly }));
afterEach(() => { document.body.replaceChildren(); state.readonly = false; });
function mount(locked = false, slides?: any[]) {
  const data = { schemaVersion: 'deckwise-native-teaching/1', documentId: 'document', revision: 'revision', pages: [{ pageId: 'page', locked, actions: [{ id: 'speech', type: 'speech', text: 'Before', audioAssetId: 'old', durationMs: 1000 }] }], media: [], identity: { pages: [{ pageId: 'page', objects: [{ logicalId: 'object', sourceShapeId: '4', type: 'shape' }] }] } };
  const parts: Array<{ GetXml(): string; Delete(): void }> = [];
  const namespaces = new WeakMap<object, string | null>();
  const add = (xml: string) => { const part = { GetXml: () => xml, Delete: () => { parts.splice(parts.indexOf(part), 1); } }; namespaces.set(part, new DOMParser().parseFromString(xml, 'application/xml').documentElement.namespaceURI); parts.push(part); return part; };
  const root = document.implementation.createDocument(NATIVE_TEACHING_NAMESPACE, 'teaching');
  root.documentElement.textContent = JSON.stringify(data);
  add(new XMLSerializer().serializeToString(root));
  const frame = document.createElement('iframe'); document.body.append(frame);
  const logic = { ...(slides ? { Slides: slides } : {}), StartAction: vi.fn(), FinalizeAction: vi.fn() };
  state.api = { WordControl: { m_oLogicDocument: logic } };
  Object.assign(frame.contentWindow!, { editor: state.api, AscDFH: { historydescription_Document_ApiBuilder: 1 }, AscBuilder: { ApiCustomXmlParts: class { GetAll() { return [...parts]; } GetByNamespace(namespace: string) { return parts.filter(part => namespaces.get(part) === namespace); } Add(xml: string) { return add(xml); } } } });
  return { logic, data, parts, frame, add };
}
it('does not serialize embedded media during teaching reads, routing, edits or save preparation', async () => {
  const { add } = mount(false, [{ cSld: { spTree: [] } }]);
  const audioReads = Array.from({ length: 22 }, () => vi.spyOn(add(`<media xmlns="${NATIVE_MEDIA_NAMESPACE}">immutable audio payload</media>`), 'GetXml'));
  expect(hasNativeTeaching()).toBe(true);
  expect(readNativeTeaching()?.pages[0].actions[0].text).toBe('Before');
  const opened = (await routeNativeTeaching('open', {}, 1))!.value as any;
  const edited = (await routeNativeTeaching('write', { viewId: opened.viewId, expectedTeachingRevision: opened.project.teachingRevision,
    actions: [{ id: 'speech', type: 'speech', text: 'After' }] }, 1))!.value as any;
  expect(edited.project.pages[0].actions[0].text).toBe('After');
  prepareNativeTeachingSave();
  for (const read of audioReads) expect(read).not.toHaveBeenCalled();
});
it('checks native presence without serializing the metadata on every object or selection', () => {
  const { parts } = mount();
  const read = vi.spyOn(parts[0], 'GetXml');
  for (let i = 0; i < 100; i++) expect(hasNativeTeaching()).toBe(true);
  expect(read).not.toHaveBeenCalled();
  expect(readNativeTeaching()?.revision).toBe('revision');
  expect(read).toHaveBeenCalledOnce();
});
it('changes one page with CAS and invalidates audio after text changes', () => {
  const { logic } = mount();
  const next = replaceNativeTeachingPage({ documentId: 'document', expectedRevision: 'revision', pageId: 'page', actions: [{ id: 'speech', type: 'speech', text: 'After', audioAssetId: 'old', durationMs: 1000 }] });
  expect(next.pages[0].actions[0]).toEqual({ id: 'speech', type: 'speech', text: 'After' });
  expect(readNativeTeaching()?.revision).toBe(next.revision);
  expect(logic.StartAction).toHaveBeenCalledOnce();
  expect(logic.FinalizeAction).toHaveBeenCalledOnce();
  expect(() => replaceNativeTeachingPage({ documentId: 'document', expectedRevision: 'revision', pageId: 'page', actions: [] })).toThrow('revision');
});
it('rejects read-only and locked changes without starting history', async () => {
  const { logic } = mount(true);
  const input = { documentId: 'document', expectedRevision: 'revision', pageId: 'page', actions: [] };
  expect(() => replaceNativeTeachingPage(input)).toThrow('locked');
  state.readonly = true;
  expect(() => replaceNativeTeachingPage(input)).toThrow('read-only');
  expect(logic.StartAction).not.toHaveBeenCalled();
  await expect(nativeTeachingAudio('missing')).rejects.toThrow('missing');
});

it('adopts verified media atomically and rejects late results after a narration edit', async () => {
  vi.stubGlobal('crypto', webcrypto);
  try {
    mount();
    const bytes = new TextEncoder().encode('test media bytes').buffer;
    const digest = Array.from(new Uint8Array(await crypto.subtle.digest('SHA-256', bytes)), b => b.toString(16).padStart(2, '0')).join('');
    const input = { documentId: 'document', expectedRevision: 'revision', pageId: 'page', actionId: 'speech', assetId: 'tts-' + digest,
      sha256: digest, mimeType: 'audio/mpeg', durationMs: 1200, bytes };
    const adopted = await adoptNativeTeachingAudio(input);
    expect(adopted.pages[0].actions[0].audioAssetId).toBe(input.assetId);
    expect(adopted.media).toHaveLength(1);
    await expect(adoptNativeTeachingAudio(input)).rejects.toThrow('revision');
    await expect(adoptNativeTeachingAudio({ ...input, expectedRevision: adopted.revision, sha256: 'a'.repeat(64) })).rejects.toThrow('digest');
    expect(readNativeTeaching()?.revision).toBe(adopted.revision);
  } finally { vi.unstubAllGlobals(); }
});

it('keeps logical targets through rename, slide duplication, reorder and native undo references', () => {
  const shape = () => ({ sourceId: '4', getUniNvProps: () => ({ cNvPr: { id: 8, name: 'Same name' } }) });
  const firstShape = shape();
  const slide = { cSld: { spTree: [firstShape] }, createDuplicate() { return { cSld: { spTree: [shape()] } }; } };
  const slides = [slide]; mount(false, slides);
  expect(nativeTeachingObjects(1)[0].objectId).toBe('object');
  expect(nativeTeachingTarget('object')).toBe(firstShape);
  firstShape.getUniNvProps = () => ({ cNvPr: { id: 8, name: 'Renamed' } });
  expect(nativeTeachingObjects(1)[0].objectId).toBe('object');
  const copy = slide.createDuplicate(); slides.push(copy as typeof slide);
  const duplicated = readNativeTeaching()!;
  expect(duplicated.pages[1].actions[0].text).toBe('Before');
  expect(duplicated.pages[1].pageId).not.toBe(duplicated.pages[0].pageId);
  expect(duplicated.pages[1].actions[0].id).not.toBe('speech');
  expect(nativeTeachingObjects(2)[0].objectId).not.toBe('object');
  expect(nativeTeachingTarget('object')).toBe(firstShape);
  slides.reverse();
  expect(readNativeTeaching()!.pages[1].pageId).toBe('page');
  slides.pop(); slides.push(slide);
  expect(readNativeTeaching()!.pages[1].actions[0].text).toBe('Before');
  prepareNativeTeachingSave();
  expect((readNativeTeaching()!.identity as any).pages[1].objects[0].logicalId).toBe('object');
});

it('locks and unlocks the native page with revision CAS', () => {
  mount();
  const next = setNativeTeachingLock({ documentId: 'document', expectedRevision: 'revision', pageId: 'page', locked: true });
  expect(() => replaceNativeTeachingPage({ documentId: 'document', expectedRevision: next.revision, pageId: 'page', actions: [] })).toThrow('locked');
  const unlocked = setNativeTeachingLock({ documentId: 'document', expectedRevision: next.revision, pageId: 'page', locked: false });
  expect(unlocked.pages[0].locked).toBe(false);
});

it('resolves nested native teaching targets without replacing them with their parent group', () => {
  const group = { sourceId: '10', spTree: [] as any[], getUniNvProps: () => ({ cNvPr: { id: 10, name: 'Album' } }) };
  const target = { sourceId: '4', group, extX: 20, extY: 10,
    transform: { TransformPointX: (x: number) => x + 10, TransformPointY: (_: number, y: number) => y + 5 },
    getUniNvProps: () => ({ cNvPr: { id: 4, name: 'Album cover' } }) };
  group.spTree.push(target);
  const slides = [{ cSld: { spTree: [group] } }, { cSld: { spTree: [] as any[] } }];
  mount(false, slides);
  const objects = nativeTeachingObjects(1);
  expect(nativeTeachingTarget('object')).toBe(target);
  const geometry = () => attentionGeometry([group] as any, objects, ['object'], 100, 50, (X, Y) => ({ X, Y }));
  expect(geometry().targets[0].shape).toBe(target);
  expect(geometry().targets[0].rect).toEqual({ x: 9.6, y: 9.4, w: 20.8, h: 21.2 });
  expect(objects.find(object => object.objectId === 'object')?.effectTargetEligible).toBe(true);
  expect(selectedAttentionTargets([group] as any, [target], objects)).toEqual(['object']);
  group.spTree.splice(0);
  slides[1].cSld.spTree.push(target);
  expect(() => geometry()).toThrow('不在当前页');
  slides[1].cSld.spTree.splice(0);
  expect(() => geometry()).toThrow('已删除');
});

it('routes write, confirmation and local reads through native history with no Host fallback', async () => {
  mount();
  const opened = (await routeNativeTeaching('open', {}, 1))!.value as any;
  expect(opened.nativeTeaching).toBe(true);
  await expect(routeNativeTeaching('write', { viewId: opened.viewId, expectedTeachingRevision: 'old', actions: [] }, 1)).rejects.toThrow('changed');
  const saved = (await routeNativeTeaching('write', { viewId: opened.viewId, expectedTeachingRevision: opened.project.teachingRevision,
    actions: [{ id: 'delay', type: 'delay', durationMs: 1500 }] }, 1))!.value as any;
  expect(saved.project.pages[0].actions[0].type).toBe('delay');
  const confirmed = (await routeNativeTeaching('confirm', { viewId: saved.viewId, expectedTeachingRevision: saved.project.teachingRevision }, 1))!.value as any;
  expect(confirmed.project.pages[0].reviewState).toBe('ready');
  await expect(routeNativeTeaching('history', { viewId: 'other-page' }, 1)).rejects.toThrow('changed');
  await expect(routeNativeTeaching('unknown-writer', {}, 1)).rejects.toThrow();
});

it('stores adopted bytes separately so subsequent text edits do not rewrite media parts', async () => {
  vi.stubGlobal('crypto', webcrypto);
  try {
    const { logic } = mount();
    const bytes = new TextEncoder().encode('immutable media').buffer;
    const digest = Array.from(new Uint8Array(await crypto.subtle.digest('SHA-256', bytes)), b => b.toString(16).padStart(2, '0')).join('');
    const adopted = await adoptNativeTeachingAudio({ documentId: 'document', expectedRevision: 'revision', pageId: 'page', actionId: 'speech', assetId: 'tts-' + digest,
      sha256: digest, mimeType: 'audio/mpeg', durationMs: 1000, bytes });
    expect(adopted.media[0].data).toBeUndefined();
    expect((await nativeTeachingAudio('tts-' + digest)).size).toBe(bytes.byteLength);
    replaceNativeTeachingPage({ documentId: 'document', expectedRevision: adopted.revision, pageId: 'page', actions: [{ id: 'speech', type: 'speech', text: 'After' }] });
    expect(readNativeTeaching()!.media[0].data).toBeUndefined();
    expect(logic.StartAction).toHaveBeenCalledTimes(2);
  } finally { vi.unstubAllGlobals(); }
});

it('restores the old carrier when a custom XML replacement fails', () => {
  const { parts } = mount();
  const original = parts[0].GetXml();
  const originalDelete = parts[0].Delete;
  parts[0].Delete = () => { originalDelete(); throw new Error('Injected delete failure'); };
  expect(() => replaceNativeTeachingPage({ documentId: 'document', expectedRevision: 'revision', pageId: 'page', actions: [] })).toThrow('Injected delete failure');
  expect(parts).toHaveLength(1);
  expect(parts[0].GetXml()).toBe(original);
  expect(readNativeTeaching()?.revision).toBe('revision');
});

it('detects same-count native animation reorder before playing a different object', () => {
  const shape = (sourceId: string, Id: string) => ({ sourceId, Id, getUniNvProps: () => ({ cNvPr: { id: Number(sourceId), name: Id } }) });
  const effect = (id: string) => ({ getObjectId: () => id, cTn: { presetClass: 1, presetID: 10, presetSubtype: 0, nodeType: 2 }, asc_getDuration: () => 500, asc_getDelay: () => 0 });
  const effects = [effect('native1'), effect('native2')];
  mount(false, [{ cSld: { spTree: [shape('4', 'native1'), shape('5', 'native2')] }, timing: { getAllAnimEffects: () => effects } }]);
  replaceNativeTeachingPage({ documentId: 'document', expectedRevision: 'revision', pageId: 'page', actions: effects.map((_, stepIndex) => ({
    id: 'animation' + stepIndex, type: 'native_animation', version: 1, sourceSha256: 'a'.repeat(64), stepIndex,
  })) });
  expect(() => assertNativeTeachingAnimations(1)).not.toThrow();
  effects.reverse();
  expect(() => assertNativeTeachingAnimations(1)).toThrow('动画顺序');
  effects.reverse();
  expect(() => assertNativeTeachingAnimations(1)).not.toThrow();
});

it('activates a fresh document from loaded notes once without a Host or parser', () => {
  const slides = [{ cSld: { spTree: [] }, notes: { cSld: { spTree: [{ getText: () => 'Fresh lecture notes' }] } } }, { cSld: { spTree: [] } }];
  const { parts, logic } = mount(false, slides);
  parts.splice(0);
  const created = initializeNativeTeaching();
  expect(created.pages).toHaveLength(2);
  expect(created.pages[0].actions[0]).toMatchObject({ type: 'speech', text: 'Fresh lecture notes', source: 'notes' });
  expect(created.pages[1].actions).toEqual([]);
  expect(initializeNativeTeaching().documentId).toBe(created.documentId);
  expect(logic.StartAction).toHaveBeenCalledOnce();
});
