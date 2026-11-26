import { webcrypto } from 'node:crypto';
import { afterEach, expect, it, vi } from 'vitest';
import {
  initializeNativeTeaching,
  insertNativeTeachingAssessment,
  assertNativeTeachingAnimations,
  nativeAnimationSignatures,
  hasNativeTeaching,
  readNativeTeaching,
  replaceNativeTeachingPage,
  replaceNativeTeachingRegion,
  nativeTeachingAudio,
  adoptNativeTeachingAudio,
  NATIVE_MEDIA_NAMESPACE,
  NATIVE_TEACHING_NAMESPACE,
  nativeTeachingObjects,
  nativeTeachingTarget,
  setNativeTeachingLock,
  prepareNativeTeachingSave,
  synchronizeNativeTeachingStructure,
} from '../../lib/deckwise-native-teaching';
import { routeNativeTeaching } from '../../lib/deckwise-native-teaching-router';
import { attentionGeometry, selectedAttentionTargets } from '../../lib/deckwise-teaching-preview';
import { hasUnsavedChanges, resetUnsavedChanges } from '../../lib/unsaved-guard';
vi.mock('../../lib/deckwise-page-lineage', () => ({
  deckwisePageLineage: { context: () => ({ sourceSha256: 'a'.repeat(64) }) },
}));
const state = vi.hoisted(() => ({ api: {}, readonly: false }));
vi.mock('../../lib/agent-plugin/editor-bridge', () => ({
  getEditorApi: () => state.api,
  requireEditorContext: () => ({
    api: state.api,
    AscFormat: { CTiming: Timing, ExecuteNoHistory: (fn: () => unknown) => fn() },
  }),
}));
vi.mock('../../lib/onlyoffice/guards/slide-source-identity', () => ({
  importedSlideShapeId: (shape: any) => shape.sourceId,
}));
vi.mock('../../lib/onlyoffice/readonly', () => ({ getReadonlyMode: () => state.readonly }));
afterEach(() => {
  document.body.replaceChildren();
  state.readonly = false;
  resetUnsavedChanges();
});
class Timing {
  effects: any[] = [];
  setParent() {}
  createDuplicate() {
    const copy = new Timing();
    copy.effects = [...this.effects];
    return copy;
  }
  getAllAnimEffects() {
    return this.effects;
  }
  createEffect(id: string, kind: number, preset: number, subtype: number) {
    let duration = 0,
      delay = 0;
    return {
      getObjectId: () => id,
      cTn: { presetClass: kind, presetID: preset, presetSubtype: subtype, setNodeType() {} },
      asc_getDuration: () => duration,
      asc_putDuration: (value: number) => {
        duration = value;
      },
      asc_getDelay: () => delay,
      asc_putDelay(value: number) {
        delay = value;
      },
    };
  }
  addEffectsToMainSequence(effects: any[]) {
    this.effects.push(...effects);
  }
  removeEffects(effects: any[]) {
    this.effects = this.effects.filter((effect) => !effects.includes(effect));
  }
}
function mount(
  locked = false,
  slides?: any[],
  objects: Array<{ logicalId: string; sourceShapeId: string; sourceObjectName?: string; type: string }> = [
    { logicalId: 'object', sourceShapeId: '4', type: 'shape' },
  ],
) {
  const data = {
    schemaVersion: 'deckwise-native-teaching/1',
    documentId: 'document',
    revision: 'revision',
    pages: [
      {
        pageId: 'page',
        locked,
        actions: [{ id: 'speech', type: 'speech', text: 'Before', audioAssetId: 'old', durationMs: 1000 }],
      },
    ],
    media: [],
    identity: { pages: [{ pageId: 'page', objects }] },
  };
  const parts: Array<{ GetXml(): string; Delete(): void }> = [];
  const namespaces = new WeakMap<object, string | null>();
  const add = (xml: string) => {
    const part = {
      GetXml: () => xml,
      Delete: () => {
        parts.splice(parts.indexOf(part), 1);
      },
    };
    namespaces.set(part, new DOMParser().parseFromString(xml, 'application/xml').documentElement.namespaceURI);
    parts.push(part);
    return part;
  };
  const root = document.implementation.createDocument(NATIVE_TEACHING_NAMESPACE, 'teaching');
  root.documentElement.textContent = JSON.stringify(data);
  add(new XMLSerializer().serializeToString(root));
  const frame = document.createElement('iframe');
  document.body.append(frame);
  const logic = { ...(slides ? { Slides: slides } : {}), StartAction: vi.fn(), FinalizeAction: vi.fn() };
  state.api = { WordControl: { m_oLogicDocument: logic } };
  Object.assign(frame.contentWindow!, {
    editor: state.api,
    AscDFH: { historydescription_Document_ApiBuilder: 1 },
    AscBuilder: {
      ApiCustomXmlParts: class {
        GetAll() {
          return [...parts];
        }
        GetByNamespace(namespace: string) {
          return parts.filter((part) => namespaces.get(part) === namespace);
        }
        Add(xml: string) {
          return add(xml);
        }
      },
    },
  });
  return { logic, data, parts, frame, add };
}
it('does not serialize embedded media during teaching reads, routing, edits or save preparation', async () => {
  const { add } = mount(false, [{ cSld: { spTree: [] } }]);
  const audioReads = Array.from({ length: 22 }, () =>
    vi.spyOn(add(`<media xmlns="${NATIVE_MEDIA_NAMESPACE}">immutable audio payload</media>`), 'GetXml'),
  );
  expect(hasNativeTeaching()).toBe(true);
  expect(readNativeTeaching()?.pages[0].actions[0].text).toBe('Before');
  const opened = (await routeNativeTeaching('open', {}, 1))!.value as any;
  const edited = (await routeNativeTeaching(
    'write',
    {
      viewId: opened.viewId,
      expectedTeachingRevision: opened.project.teachingRevision,
      actions: [{ id: 'speech', type: 'speech', text: 'After' }],
    },
    1,
  ))!.value as any;
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
it('records an online page reorder in the existing history action', () => {
  const first = { cSld: { spTree: [] } };
  const second = { cSld: { spTree: [] } };
  const { logic, parts } = mount(false, [first, second]);
  synchronizeNativeTeachingStructure();
  const original = readNativeTeaching()!;
  const savedPart = parts[0];
  synchronizeNativeTeachingStructure();
  expect(parts[0]).toBe(savedPart);

  logic.Slides!.reverse();
  synchronizeNativeTeachingStructure();
  const xml = new DOMParser().parseFromString(parts[0].GetXml(), 'application/xml');
  const saved = JSON.parse(xml.documentElement.textContent!);
  expect(saved.revision).not.toBe(original.revision);
  expect(saved.pages).toEqual([...original.pages].reverse());
  expect(saved.identity.pages.map((page: { pageId: string }) => page.pageId)).toEqual(
    saved.pages.map((page: { pageId: string }) => page.pageId),
  );
  expect(logic.StartAction).not.toHaveBeenCalled();
  expect(logic.FinalizeAction).not.toHaveBeenCalled();
});
it('persists renamed teaching objects so a new native session can restore their references', () => {
  let name = 'Before';
  const shape = { sourceId: '4', getUniNvProps: () => ({ cNvPr: { id: 8, name } }) };
  const slides = [{ cSld: { spTree: [shape] } }];
  const { parts } = mount(false, slides, [{ logicalId: 'object', sourceShapeId: '4', sourceObjectName: name, type: 'shape' }]);
  expect(nativeTeachingTarget('object')).toBeUndefined();
  expect(nativeTeachingObjects(1)[0].objectId).toBe('object');
  const original = parts[0];
  name = 'After';
  synchronizeNativeTeachingStructure();
  expect(parts[0]).not.toBe(original);
  expect(parts[0].GetXml()).toContain('"sourceObjectName":"After"');
});
it('marks teaching writes dirty so the host saves them before reporting success', () => {
  mount();
  resetUnsavedChanges();
  expect(hasUnsavedChanges()).toBe(false);
  replaceNativeTeachingPage({ documentId: 'document', expectedRevision: 'revision', pageId: 'page', actions: [] });
  expect(hasUnsavedChanges()).toBe(true);
});
it('changes one page with CAS and invalidates audio after text changes', () => {
  const { logic } = mount();
  const next = replaceNativeTeachingPage({
    documentId: 'document',
    expectedRevision: 'revision',
    pageId: 'page',
    actions: [{ id: 'speech', type: 'speech', text: 'After', audioAssetId: 'old', durationMs: 1000 }],
  });
  expect(next.pages[0].actions[0]).toEqual({ id: 'speech', type: 'speech', text: 'After' });
  expect(readNativeTeaching()?.revision).toBe(next.revision);
  expect(logic.StartAction).toHaveBeenCalledOnce();
  expect(logic.FinalizeAction).toHaveBeenCalledOnce();
  expect(() =>
    replaceNativeTeachingPage({ documentId: 'document', expectedRevision: 'revision', pageId: 'page', actions: [] }),
  ).toThrow('revision');
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
    const digest = Array.from(new Uint8Array(await crypto.subtle.digest('SHA-256', bytes)), (b) =>
      b.toString(16).padStart(2, '0'),
    ).join('');
    const input = {
      documentId: 'document',
      expectedRevision: 'revision',
      pageId: 'page',
      actionId: 'speech',
      assetId: 'tts-' + digest,
      sha256: digest,
      mimeType: 'audio/mpeg',
      durationMs: 1200,
      bytes,
    };
    const adopted = await adoptNativeTeachingAudio(input);
    expect(adopted.pages[0].actions[0].audioAssetId).toBe(input.assetId);
    expect(adopted.media).toHaveLength(1);
    await expect(adoptNativeTeachingAudio(input)).rejects.toThrow('revision');
    await expect(
      adoptNativeTeachingAudio({ ...input, expectedRevision: adopted.revision, sha256: 'a'.repeat(64) }),
    ).rejects.toThrow('digest');
    expect(readNativeTeaching()?.revision).toBe(adopted.revision);
  } finally {
    vi.unstubAllGlobals();
  }
});

it.each([false, true])('preserves teaching references through native slide changes (online=%s)', (online) => {
  const shape = (Id: string) => ({ Id, sourceId: '4', getUniNvProps: () => ({ cNvPr: { id: 8, name: 'Same name' } }) });
  const firstShape = shape('1_2');
  const slide = {
    cSld: { spTree: [firstShape] },
    createDuplicate() {
      return { Id: '1_3', cSld: { spTree: [shape('1_4')] } };
    },
  };
  const slides: Array<{ Id?: string; cSld: typeof slide.cSld }> = [slide];
  mount(false, slides);
  if (online) Object.assign(state.api, { CoAuthoringApi: { getDocId: () => 'shared-doc' } });
  expect(nativeTeachingObjects(1)[0].objectId).toBe('object');
  expect(nativeTeachingTarget('object')).toBe(firstShape);
  firstShape.getUniNvProps = () => ({ cNvPr: { id: 8, name: 'Renamed' } });
  expect(nativeTeachingObjects(1)[0].objectId).toBe('object');
  const copy = slide.createDuplicate();
  slides.push(copy);
  const duplicated = readNativeTeaching()!;
  expect(duplicated.pages[1].actions[0].text).toBe('Before');
  expect(duplicated.pages[1].pageId).not.toBe(duplicated.pages[0].pageId);
  expect(duplicated.pages[1].actions[0].id).not.toBe('speech');
  expect(nativeTeachingObjects(2)[0].objectId).not.toBe('object');
  if (online) {
    expect(duplicated.pages[1].pageId).toBe('page-shared-doc-1_3');
    expect(nativeTeachingObjects(2)[0].objectId).toBe('object-shared-doc-1_4');
  }
  expect(nativeTeachingTarget('object')).toBe(firstShape);
  slides.reverse();
  expect(readNativeTeaching()!.pages[1].pageId).toBe('page');
  slides.pop();
  slides.push(slide);
  expect(readNativeTeaching()!.pages[1].actions[0].text).toBe('Before');
  prepareNativeTeachingSave();
  expect((readNativeTeaching()!.identity as any).pages[1].objects[0].logicalId).toBe('object');
});

it('locks and unlocks the native page with revision CAS', () => {
  mount();
  const next = setNativeTeachingLock({
    documentId: 'document',
    expectedRevision: 'revision',
    pageId: 'page',
    locked: true,
  });
  expect(() =>
    replaceNativeTeachingPage({ documentId: 'document', expectedRevision: next.revision, pageId: 'page', actions: [] }),
  ).toThrow('locked');
  const unlocked = setNativeTeachingLock({
    documentId: 'document',
    expectedRevision: next.revision,
    pageId: 'page',
    locked: false,
  });
  expect(unlocked.pages[0].locked).toBe(false);
});

it('distinguishes native groups from leaf narration targets', async () => {
  const group = { sourceId: '10', spTree: [{ sourceId: '4', x: 1, y: 2, extX: 20, extY: 10 }], extX: 40, extY: 30 };
  mount(false, [{ cSld: { spTree: [group] } }]);
  const context = (await routeNativeTeaching('narration-context', {}, 1))!.value as any;
  expect(context.evidence.objects.find((object: any) => object.objectId === 'object')?.kind).toBe('shape');
  expect(context.evidence.objects.find((object: any) => object.objectId !== 'object')?.kind).toBe('group');
});

it('resolves nested native teaching targets without replacing them with their parent group', () => {
  const group = { sourceId: '10', spTree: [] as any[], getUniNvProps: () => ({ cNvPr: { id: 10, name: 'Album' } }) };
  const target = {
    sourceId: '4',
    group,
    extX: 20,
    extY: 10,
    transform: { TransformPointX: (x: number) => x + 10, TransformPointY: (_: number, y: number) => y + 5 },
    getUniNvProps: () => ({ cNvPr: { id: 4, name: 'Album cover' } }),
  };
  group.spTree.push(target);
  const slides = [{ cSld: { spTree: [group] } }, { cSld: { spTree: [] as any[] } }];
  mount(false, slides);
  const objects = nativeTeachingObjects(1);
  expect(nativeTeachingTarget('object')).toBe(target);
  const geometry = () => attentionGeometry([group] as any, objects, ['object'], 100, 50, (X, Y) => ({ X, Y }));
  expect(geometry().targets[0].shape).toBe(target);
  expect(geometry().targets[0].rect).toEqual({ x: 9.6, y: 9.4, w: 20.8, h: 21.2 });
  expect(objects.find((object) => object.objectId === 'object')?.effectTargetEligible).toBe(true);
  expect(selectedAttentionTargets([group] as any, [target], objects)).toEqual(['object']);
  group.spTree.splice(0);
  slides[1].cSld.spTree.push(target);
  expect(() => geometry()).toThrow('不在当前页');
  slides[1].cSld.spTree.splice(0);
  expect(() => geometry()).toThrow('已删除');
});

it('recovers a uniquely named target when the saved PPTX renumbers it', () => {
  const shape = { sourceId: '9', getUniNvProps: () => ({ cNvPr: { name: 'Album cover' } }) };
  const reused = { sourceId: '4', getUniNvProps: () => ({ cNvPr: { name: 'Photo frame' } }) };
  mount(
    false,
    [{ cSld: { spTree: [reused, shape] } }],
    [{ logicalId: 'object', sourceShapeId: '4', sourceObjectName: 'Album cover', type: 'shape' }],
  );
  expect(nativeTeachingObjects(1)[1].objectId).toBe('object');
  expect(nativeTeachingTarget('object')).toBe(shape);
});

it('leaves a renumbered target unresolved when the shape name is ambiguous', () => {
  const shape = (sourceId: string, name: string) => ({ sourceId, getUniNvProps: () => ({ cNvPr: { name } }) });
  mount(
    false,
    [{ cSld: { spTree: [shape('4', 'Frame'), shape('8', 'Photo'), shape('9', 'Photo')] } }],
    [{ logicalId: 'object', sourceShapeId: '4', sourceObjectName: 'Photo', type: 'shape' }],
  );
  nativeTeachingObjects(1);
  expect(nativeTeachingTarget('object')).toBeUndefined();
});

it('shares a new native object identity across editors while keeping it scoped to the online document', () => {
  const open = (key: string) => {
    const shape = { Id: '1_2', getUniNvProps: () => ({ cNvPr: { name: 'New shape' } }) };
    mount(false, [{ cSld: { spTree: [shape] } }]);
    Object.assign(state.api, { CoAuthoringApi: { getDocId: () => key } });
    return nativeTeachingObjects(1)[0].objectId;
  };
  const objectId = open('session-1');
  expect(open('session-1')).toBe(objectId);
  expect(open('session-2')).not.toBe(objectId);
});

it('routes narration on an inserted slide to the same page in connected editors', () => {
  const open = (key: string) => {
    const slides = [{ Id: '42', cSld: { spTree: [] } }];
    const { add, parts } = mount(false, slides);
    Object.assign(state.api, { CoAuthoringApi: { getDocId: () => key } });
    readNativeTeaching();
    slides.unshift({ Id: '1_8', cSld: { spTree: [] } });
    const teaching = readNativeTeaching()!;
    expect(teaching.pages[1].pageId).toBe('page');
    expect(teaching.pages[1].actions[0].text).toBe('Before');
    return { add, parts, pageId: teaching.pages[0].pageId };
  };
  const writer = open('session-1');
  const saved = replaceNativeTeachingPage({
    documentId: 'document',
    expectedRevision: 'revision',
    pageId: writer.pageId,
    actions: [{ id: 'inserted-speech', type: 'speech', text: 'New slide narration' }],
  });
  const peer = open('session-1');
  peer.parts[0].Delete();
  const xml = document.implementation.createDocument(NATIVE_TEACHING_NAMESPACE, 'teaching');
  xml.documentElement.textContent = JSON.stringify(saved);
  peer.add(new XMLSerializer().serializeToString(xml));
  expect(readNativeTeaching()!.pages[0].actions[0]?.text).toBe('New slide narration');
  expect(open('session-2').pageId).not.toBe(writer.pageId);
});

it('routes write, confirmation and local reads through native history with no Host fallback', async () => {
  mount();
  const opened = (await routeNativeTeaching('open', {}, 1))!.value as any;
  expect(opened.nativeTeaching).toBe(true);
  await expect(
    routeNativeTeaching('write', { viewId: opened.viewId, expectedTeachingRevision: 'old', actions: [] }, 1),
  ).rejects.toThrow('changed');
  const saved = (await routeNativeTeaching(
    'write',
    {
      viewId: opened.viewId,
      expectedTeachingRevision: opened.project.teachingRevision,
      actions: [{ id: 'delay', type: 'delay', durationMs: 1500 }],
    },
    1,
  ))!.value as any;
  expect(saved.project.pages[0].actions[0].type).toBe('delay');
  const confirmed = (await routeNativeTeaching(
    'confirm',
    { viewId: saved.viewId, expectedTeachingRevision: saved.project.teachingRevision },
    1,
  ))!.value as any;
  expect(confirmed.project.pages[0].reviewState).toBe('ready');
  await expect(routeNativeTeaching('history', { viewId: 'other-page' }, 1)).rejects.toThrow('changed');
  await expect(routeNativeTeaching('unknown-writer', {}, 1)).rejects.toThrow();
});

it('stores adopted bytes separately so subsequent text edits do not rewrite media parts', async () => {
  vi.stubGlobal('crypto', webcrypto);
  try {
    const { logic } = mount();
    const bytes = new TextEncoder().encode('immutable media').buffer;
    const digest = Array.from(new Uint8Array(await crypto.subtle.digest('SHA-256', bytes)), (b) =>
      b.toString(16).padStart(2, '0'),
    ).join('');
    const adopted = await adoptNativeTeachingAudio({
      documentId: 'document',
      expectedRevision: 'revision',
      pageId: 'page',
      actionId: 'speech',
      assetId: 'tts-' + digest,
      sha256: digest,
      mimeType: 'audio/mpeg',
      durationMs: 1000,
      bytes,
    });
    expect(adopted.media[0].data).toBeUndefined();
    expect((await nativeTeachingAudio('tts-' + digest)).size).toBe(bytes.byteLength);
    replaceNativeTeachingPage({
      documentId: 'document',
      expectedRevision: adopted.revision,
      pageId: 'page',
      actions: [{ id: 'speech', type: 'speech', text: 'After' }],
    });
    expect(readNativeTeaching()!.media[0].data).toBeUndefined();
    expect(logic.StartAction).toHaveBeenCalledTimes(2);
  } finally {
    vi.unstubAllGlobals();
  }
});

it('restores the old carrier when a custom XML replacement fails', () => {
  const { parts } = mount();
  const original = parts[0].GetXml();
  const originalDelete = parts[0].Delete;
  parts[0].Delete = () => {
    originalDelete();
    throw new Error('Injected delete failure');
  };
  expect(() =>
    replaceNativeTeachingPage({ documentId: 'document', expectedRevision: 'revision', pageId: 'page', actions: [] }),
  ).toThrow('Injected delete failure');
  expect(parts).toHaveLength(1);
  expect(parts[0].GetXml()).toBe(original);
  expect(readNativeTeaching()?.revision).toBe('revision');
});

it('detects same-count native animation reorder before playing a different object', () => {
  const shape = (sourceId: string, Id: string) => ({
    sourceId,
    Id,
    getUniNvProps: () => ({ cNvPr: { id: Number(sourceId), name: Id } }),
  });
  const effect = (id: string) => ({
    getObjectId: () => id,
    cTn: { presetClass: 1, presetID: 10, presetSubtype: 0, nodeType: 2 },
    asc_getDuration: () => 500,
    asc_getDelay: () => 0,
  });
  const effects = [effect('native1'), effect('native2')];
  mount(false, [
    { cSld: { spTree: [shape('4', 'native1'), shape('5', 'native2')] }, timing: { getAllAnimEffects: () => effects } },
  ]);
  replaceNativeTeachingPage({
    documentId: 'document',
    expectedRevision: 'revision',
    pageId: 'page',
    actions: effects.map((_, stepIndex) => ({
      id: 'animation' + stepIndex,
      type: 'native_animation',
      version: 1,
      sourceSha256: 'a'.repeat(64),
      stepIndex,
    })),
  });
  expect(() => assertNativeTeachingAnimations(1)).not.toThrow();
  effects.reverse();
  expect(() => assertNativeTeachingAnimations(1)).toThrow('动画顺序');
  effects.reverse();
  expect(() => assertNativeTeachingAnimations(1)).not.toThrow();
});

it('rejects an online binding when a peer reordered the animation steps after selection', async () => {
  vi.stubGlobal('crypto', webcrypto);
  try {
    const shapes = ['4', '5'].map((sourceId) => ({
      sourceId,
      Id: sourceId,
      getUniNvProps: () => ({ cNvPr: { id: sourceId, name: 'Shape ' + sourceId } }),
    }));
    const effects = shapes.map((shape) => ({
      getObjectId: () => shape.Id,
      cTn: { presetClass: 1, presetID: 10, presetSubtype: 0, nodeType: 2 },
      asc_getDuration: () => 500,
      asc_getDelay: () => 0,
    }));
    const { logic } = mount(false, [{ cSld: { spTree: shapes }, timing: { getAllAnimEffects: () => effects } }]);
    readNativeTeaching();
    const payload = {
      viewId: 'document.page',
      expectedTeachingRevision: 'revision',
      actions: nativeAnimationSignatures(1, true).map((expectedAnimation, stepIndex) => ({
        id: 'step-' + stepIndex,
        type: 'native_animation',
        version: 1,
        stepIndex,
        expectedAnimation,
      })),
    };
    effects.reverse();
    await expect(routeNativeTeaching('write', payload, 1, true)).rejects.toThrow('在绑定后发生变化');
    expect(logic.StartAction).not.toHaveBeenCalled();
    effects.reverse();
    await routeNativeTeaching('write', payload, 1, true);
    expect(readNativeTeaching()?.pages[0].actions[0]).toMatchObject({ sourceSha256: /^[a-f0-9]{64}$/ });
    expect(readNativeTeaching()?.pages[0].actions[0]).not.toHaveProperty('expectedAnimation');
    expect(() => assertNativeTeachingAnimations(1)).not.toThrow();
  } finally {
    vi.unstubAllGlobals();
  }
});

it('leaves the original timing tree alone on a teaching-only edit even if an existing cue has no native mirror', () => {
  const timing = new Timing();
  const slide = {
    cSld: { spTree: [{ sourceId: '4', Id: 'shape-4' }] },
    timing,
    setTiming: vi.fn(),
  };
  const { data, parts, add } = mount(false, [slide]);
  const speech = {
    id: 'speech',
    type: 'speech',
    text: 'Before',
    visualCues: [
      { id: 'fade', atMs: 0, endMs: 500, effect: { type: 'element_animation', effect: 'fadeIn', elementId: 'object' } },
    ],
  };
  const root = document.implementation.createDocument(NATIVE_TEACHING_NAMESPACE, 'teaching');
  root.documentElement.textContent = JSON.stringify({
    ...data,
    pages: [{ ...data.pages[0], actions: [speech] }],
  });
  parts[0].Delete();
  add(new XMLSerializer().serializeToString(root));
  const updated = replaceNativeTeachingPage({
    documentId: 'document',
    expectedRevision: 'revision',
    pageId: 'page',
    actions: [{ ...speech, text: 'After' }],
  });
  expect(updated.pages[0].actions[0].text).toBe('After');
  expect(slide.setTiming).not.toHaveBeenCalled();
});

it('activates a fresh document from loaded notes once without a Host or parser', () => {
  const slides = [
    { cSld: { spTree: [] }, notes: { cSld: { spTree: [{ getText: () => 'Fresh lecture notes' }] } } },
    { cSld: { spTree: [] } },
  ];
  const { parts, logic } = mount(false, slides);
  parts.splice(0);
  const created = initializeNativeTeaching();
  expect(created.pages).toHaveLength(2);
  expect(created.pages[0].actions[0]).toMatchObject({ type: 'speech', text: 'Fresh lecture notes', source: 'notes' });
  expect(created.pages[1].actions).toEqual([]);
  expect(initializeNativeTeaching().documentId).toBe(created.documentId);
  expect(logic.StartAction).toHaveBeenCalledOnce();
});

it('retains generated logical requirements and live object evidence after page projection', async () => {
  const { parts, add, data } = mount(false, [{ cSld: { spTree: [{ sourceId: '4', x: 1, y: 2, extX: 3, extY: 4 }] } }]);
  const teaching = {
    pageType: 'slide',
    phase: 'opening',
    objective: 'Greet',
    durationSeconds: 30,
    requiredObjects: [],
  };
  const root = document.implementation.createDocument(NATIVE_TEACHING_NAMESPACE, 'teaching');
  root.documentElement.textContent = JSON.stringify({
    ...data,
    title: 'Generated lesson',
    generation: {
      pages: [{ pageId: 'page', teaching, narration: 'Planned narration evidence' }],
      educationContext: { schoolStage: 'KINDERGARTEN', subject: 'ENGLISH' },
    },
  });
  parts[0].Delete();
  add(new XMLSerializer().serializeToString(root));
  const opened = (await routeNativeTeaching('open', {}, 1))!.value as any;
  expect(opened.teaching).toEqual(teaching);
  expect(opened.project.title).toBe('Generated lesson');
  expect(opened.project.educationContext.schoolStage).toBe('KINDERGARTEN');
  const context = (await routeNativeTeaching(
    'narration-context',
    { viewId: opened.viewId, expectedTeachingRevision: opened.project.teachingRevision },
    1,
  ))!.value as any;
  expect(context.evidence.objects[0]).toMatchObject({ objectId: 'object', bbox: [1, 2, 3, 4] });
  expect(context.evidence.notes).toBe('Planned narration evidence');
});

it.each(['highlight', 'spotlight'])(
  'resolves generated %s and animation against adopted audio in the same history action',
  async (type) => {
    vi.stubGlobal('crypto', webcrypto);
    try {
      const slide = {
        cSld: { spTree: [{ sourceId: '4', Id: 'shape-4' }] },
        timing: null as Timing | null,
        setTiming(value: Timing | null) {
          this.timing = value;
        },
      };
      const { logic } = mount(false, [slide]);
      const teaching = replaceNativeTeachingPage({
        documentId: 'document',
        expectedRevision: 'revision',
        pageId: 'page',
        actions: [
          {
            id: 'speech',
            type: 'speech',
            text: 'Look here.',
            attention: { type, objectIds: ['object'], animation: 'fadeIn' },
          },
        ],
      });
      const bytes = new TextEncoder().encode('generated audio bytes').buffer;
      const sha256 = Array.from(new Uint8Array(await crypto.subtle.digest('SHA-256', bytes)), (b) =>
        b.toString(16).padStart(2, '0'),
      ).join('');
      const input = {
        documentId: 'document',
        expectedRevision: teaching.revision,
        pageId: 'page',
        actionId: 'speech',
        assetId: 'tts-generated',
        sha256,
        mimeType: 'audio/mpeg',
        durationMs: 1000,
        bytes,
      };
      const adopted = await adoptNativeTeachingAudio(input),
        action = adopted.pages[0].actions[0];
      expect(action.attention).toBeUndefined();
      expect(slide.timing?.getAllAnimEffects()).toHaveLength(1);
      const timing = slide.timing;
      const repeated = await adoptNativeTeachingAudio({ ...input, expectedRevision: adopted.revision });
      expect(slide.timing).toBe(timing);
      expect(repeated.pages[0].actions[0].visualCues).toEqual(action.visualCues);
      expect(action.visualCues).toMatchObject([
        { atMs: 0, endMs: 600, effect: { type: 'element_animation', elementId: 'object', effect: 'fadeIn' } },
        { atMs: 600, endMs: 1000, effect: { type } },
      ]);
      expect(logic.StartAction).toHaveBeenCalledTimes(3);
      const changed = replaceNativeTeachingPage({
        documentId: 'document',
        expectedRevision: repeated.revision,
        pageId: 'page',
        actions: [{ id: 'speech', type: 'speech', text: 'Look here.', attention: { type, objectIds: ['deleted'] } }],
      });
      await expect(adoptNativeTeachingAudio({ ...input, expectedRevision: changed.revision })).rejects.toThrow(
        'target',
      );
      expect(readNativeTeaching()!.revision).toBe(changed.revision);
    } finally {
      vi.unstubAllGlobals();
    }
  },
);

it('resolves newly repaired attention using retained embedded audio without re-adoption', async () => {
  vi.stubGlobal('crypto', webcrypto);
  try {
    const slide = {
      cSld: { spTree: [{ sourceId: '4', Id: 'shape-4' }] },
      timing: null as Timing | null,
      setTiming(value: Timing | null) {
        this.timing = value;
      },
    };
    mount(false, [slide]);
    const bytes = new TextEncoder().encode('existing audio').buffer;
    const sha256 = Array.from(new Uint8Array(await crypto.subtle.digest('SHA-256', bytes)), (b) =>
      b.toString(16).padStart(2, '0'),
    ).join('');
    const adopted = await adoptNativeTeachingAudio({
      documentId: 'document',
      expectedRevision: 'revision',
      pageId: 'page',
      actionId: 'speech',
      assetId: 'tts-existing',
      sha256,
      mimeType: 'audio/mpeg',
      durationMs: 1200,
      bytes,
    });
    const repaired = replaceNativeTeachingPage({
      documentId: 'document',
      expectedRevision: adopted.revision,
      pageId: 'page',
      actions: [
        {
          ...adopted.pages[0].actions[0],
          visualCues: [{ id: 'old', atMs: 0, endMs: 1200, effect: { type: 'highlight', elementIds: ['object'] } }],
          attention: { type: 'spotlight', objectIds: ['object'], animation: 'pulse' },
        },
      ],
    });
    expect(repaired.pages[0].actions[0].attention).toBeUndefined();
    expect(repaired.pages[0].actions[0]).toMatchObject({
      text: 'Before',
      audioAssetId: 'tts-existing',
      durationMs: 1200,
    });
    expect(repaired.pages[0].actions[0].visualCues).toMatchObject([
      { atMs: 0, endMs: 600, effect: { type: 'element_animation', effect: 'pulse' } },
      { atMs: 600, endMs: 1200, effect: { type: 'spotlight' } },
    ]);
    expect(repaired.media).toEqual(adopted.media);
  } finally {
    vi.unstubAllGlobals();
  }
});

it('persists manual speech animation cues with unrelated native effects and rolls back a failed carrier write', () => {
  const previous = new Timing();
  previous.effects.push(previous.createEffect('unrelated', 1, 10, 0));
  const slide = {
    cSld: { spTree: [{ sourceId: '4', Id: 'shape-4' }] },
    timing: previous,
    setTiming(value: Timing) {
      this.timing = value;
    },
  };
  const { parts } = mount(false, [slide]);
  const actions = [
    {
      id: 'speech',
      type: 'speech',
      text: 'Before',
      visualCues: [
        {
          id: 'cue',
          atMs: 250,
          endMs: 850,
          trigger: { version: 1, kind: 'end' },
          effect: { type: 'element_animation', elementId: 'object', effect: 'pulse' },
        },
      ],
    },
  ];
  const oldXml = parts[0].GetXml(),
    originalDelete = parts[0].Delete;
  parts[0].Delete = () => {
    originalDelete();
    throw Error('Injected failure');
  };
  const input = { documentId: 'document', expectedRevision: 'revision', pageId: 'page', actions };
  expect(() => replaceNativeTeachingPage(input)).toThrow('Injected failure');
  expect(slide.timing).toBe(previous);
  expect(parts[0].GetXml()).toBe(oldXml);
  const saved = replaceNativeTeachingPage(input);
  expect(slide.timing.getAllAnimEffects().map((effect) => effect.getObjectId())).toEqual(['unrelated', 'shape-4']);
  const committed = slide.timing;
  replaceNativeTeachingPage({ ...input, expectedRevision: saved.revision });
  expect(slide.timing).toBe(committed);
  const revision = readNativeTeaching()!.revision,
    before = parts[0].GetXml();
  let fail = true;
  slide.setTiming = (value: Timing) => {
    slide.timing = value;
    if (fail) {
      fail = false;
      throw Error('Injected timing failure');
    }
  };
  const changed = structuredClone(actions);
  changed[0].visualCues[0].endMs = 950;
  expect(() => replaceNativeTeachingPage({ ...input, expectedRevision: revision, actions: changed })).toThrow(
    'Injected timing failure',
  );
  expect(slide.timing).toBe(committed);
  expect(parts[0].GetXml()).toBe(before);
});

it('opens the existing assessment editor from native data and saves question images with answers using CAS', async () => {
  vi.stubGlobal('crypto', webcrypto);
  try {
    mount(false, [{ cSld: { spTree: [] } }]);
    replaceNativeTeachingPage({
      documentId: 'document',
      expectedRevision: 'revision',
      pageId: 'page',
      actions: [
        { id: 'reading', type: 'speech', text: 'Choose\nA. One\nB. Two', audioAssetId: 'old', durationMs: 1000 },
        {
          id: 'q',
          type: 'interaction_checkpoint',
          checkpointKind: 'chapter_summary',
          checkpointId: 'q',
          chapterId: 'page',
          chapterTitle: 'Count',
          question: 'Choose',
          response: { kind: 'single', options: ['One', 'Two'], answerIndices: [0] },
          referenceAnswer: { summary: 'One', keyPoints: [] },
          context: { chapterSummary: 'Count', sourceSceneIds: ['page'] },
        },
      ],
    });
    const opened = (await routeNativeTeaching('open', {}, 1))!.value as any;
    expect(opened.assessment.classroom.scenes[0].content.questions[0].answer).toEqual(['A']);
    const classroom = structuredClone(opened.assessment.classroom);
    const question = classroom.scenes[0].content.questions[0];
    question.question = 'Choose two';
    question.answer = ['B'];
    question.referenceAnswer = 'Two';
    question.options[1].image = { src: 'data:image/png;base64,aW1hZ2U=', mimeType: 'image/png' };
    const args = {
      projectId: 'document',
      expectedRevision: opened.project.teachingRevision,
      expectedSourceRevision: opened.project.teachingRevision,
      classroom,
      changedPageIds: ['page'],
    };
    const saved = (await routeNativeTeaching('assessment-write', args, 1))!.value as any;
    expect(saved.classroom.scenes[0].content.questions[0]).toEqual(question);
    const carrier = readNativeTeaching()!;
    expect(carrier.pages[0].actions.find((a) => a.id === 'q')?.response).toEqual({
      kind: 'single',
      options: ['One', 'Two'],
      answerIndices: [1],
    });
    expect(carrier.pages[0].actions[0].audioAssetId).toBeUndefined();
    expect(carrier.pages[0].actions[0].text).toBe('Choose two\nA. One\nB. Two');
    await expect(routeNativeTeaching('assessment-write', args, 1)).rejects.toThrow('changed');
    const reopened = (await routeNativeTeaching('open', {}, 1))!.value as any;
    expect(reopened.assessment.classroom.scenes[0].content.questions[0].options[1].image).toEqual(
      question.options[1].image,
    );
  } finally {
    vi.unstubAllGlobals();
  }
});

it('keeps discussion reference prompts outside the strict playback action through slide duplication', async () => {
  const slide = {
    cSld: { spTree: [] },
    createDuplicate() {
      return { cSld: { spTree: [] } };
    },
  };
  const slides = [slide];
  mount(false, slides);
  const next = replaceNativeTeachingPage({
    documentId: 'document',
    expectedRevision: 'revision',
    pageId: 'page',
    actions: [
      {
        id: 'q',
        type: 'interaction_checkpoint',
        checkpointKind: 'chapter_summary',
        checkpointId: 'q',
        chapterId: 'page',
        chapterTitle: 'Discuss',
        question: 'Why?',
        referenceAnswer: { summary: 'Because', keyPoints: [] },
        context: { chapterSummary: 'Discuss', sourceSceneIds: ['page'] },
      },
    ],
  });
  const view = (await routeNativeTeaching('open', {}, 1))!.value as any;
  expect(view.embeddedAssessment.classroom.scenes[0].content.type).toBe('slide');
  const classroom = structuredClone(view.embeddedAssessment.classroom);
  classroom.scenes[0].actions[0].referenceAnswer = { summary: 'Evidence', keyPoints: ['Reason'], label: 'Sample' };
  classroom.scenes[0].actions[0].constraintPrompt = 'Give one reason';
  await routeNativeTeaching(
    'assessment-write',
    {
      projectId: 'document',
      expectedRevision: next.revision,
      expectedSourceRevision: next.revision,
      classroom,
      changedPageIds: ['page'],
    },
    1,
  );
  const saved = readNativeTeaching()!;
  expect(saved.pages[0].actions[0].referenceAnswer).toEqual({ summary: 'Evidence', keyPoints: ['Reason'] });
  expect(saved.pages[0].actions[0]).not.toHaveProperty('constraintPrompt');
  const reopened = (await routeNativeTeaching('assessment-context', {}, 1))!.value as any;
  expect(reopened.classroom.scenes[0].actions[0].constraintPrompt).toBe('Give one reason');
  slides.push(slide.createDuplicate() as typeof slide);
  const duplicated = (await routeNativeTeaching('assessment-context', {}, 2))!.value as any;
  expect(duplicated.classroom.scenes[1].actions[0].id).not.toBe('q');
  expect(duplicated.classroom.scenes[1].actions[0].constraintPrompt).toBe('Give one reason');
  expect(duplicated.classroom.scenes[1].actions[0].referenceAnswer.label).toBe('Sample');
});
it('restores the immutable copy seed when Undo precedes the first copied-page carrier write', () => {
  const slide = {
    cSld: { spTree: [] },
    createDuplicate() {
      return { cSld: { spTree: [] } };
    },
  };
  const slides = [slide],
    { parts, add } = mount(false, slides);
  readNativeTeaching();
  slides.push(slide.createDuplicate() as typeof slide);
  const copied = readNativeTeaching()!,
    originalXml = parts[0].GetXml();
  replaceNativeTeachingPage({
    documentId: copied.documentId,
    expectedRevision: copied.revision,
    pageId: copied.pages[1].pageId,
    actions: [{ id: copied.pages[1].actions[0].id, type: 'speech', text: 'CHANGED' }],
  });
  parts.splice(0);
  add(originalXml);
  expect(readNativeTeaching()!.pages[1].actions[0].text).toBe('Before');
});
it.each(['practice', 'pbl', 'discussion'])(
  'projects %s as a bound interaction rather than an invalid Quiz',
  async (kind) => {
    mount(false, [
      {
        cSld: {
          spTree: [
            {
              sourceId: '4',
              getUniNvProps: () => ({ cNvPr: { name: 'Question', id: 8 } }),
              getDocContent: () => ({ GetText: () => 'Why?' }),
            },
          ],
        },
      },
    ]);
    replaceNativeTeachingPage({
      documentId: 'document',
      expectedRevision: 'revision',
      pageId: 'page',
      actions: [
        {
          id: 'q',
          type: 'interaction_checkpoint',
          checkpointKind: 'chapter_summary',
          checkpointId: 'cp',
          chapterId: 'page',
          chapterTitle: 'Discuss',
          question: 'Why?',
          response: { kind, options: [], answerIndices: [] },
          referenceAnswer: { summary: '', keyPoints: [] },
          context: { chapterSummary: 'Discuss', sourceSceneIds: ['page'] },
        },
      ],
    });
    const value = (await routeNativeTeaching('assessment-context', {}, 1))!.value as any;
    expect(value.classroom.scenes[0].content.type).toBe('slide');
    expect(value.classroom.scenes[0].actions[0]).not.toHaveProperty('response');
    const response = structuredClone(readNativeTeaching()!.pages[0].actions[0].response);
    expect(response).toMatchObject({ kind });
    const action = value.classroom.scenes[0].actions[0];
    action.referenceAnswer.summary = 'Updated reference';
    await routeNativeTeaching(
      'assessment-write',
      {
        projectId: value.projectId,
        expectedRevision: value.revision,
        expectedSourceRevision: value.sourceRevision,
        classroom: value.classroom,
        changedPageIds: ['page'],
      },
      1,
    );
    expect(readNativeTeaching()!.pages[0].actions[0].response).toEqual(response);
    expect(readNativeTeaching()!.pages[0].actions[0].referenceAnswer).toEqual({
      summary: 'Updated reference',
      keyPoints: [],
    });
  },
);

it('carries accepted script provenance atomically with speech and preserves it through older page writers', async () => {
  const { logic } = mount(false, [{ cSld: { spTree: [] } }]);
  const enrichment = {
    schemaVersion: 'deckwise-teaching-enrichment/1' as const,
    originalPages: [{ pageId: 'page', pageNumber: 1 }],
    script: {
      sha256: 'a'.repeat(64),
      policy: 'preserve' as const,
      silentPageIds: [],
      segments: [{ id: 'segment-1', text: 'Exact. 1.25 kg.', sourceLine: 1, pageId: 'page', actionId: 'speech' }],
    },
  };
  const opened = (await routeNativeTeaching('open', {}, 1))!.value as any;
  expect(opened.enrichmentCapabilities).toEqual({ version: 1 });
  const written = (await routeNativeTeaching(
    'write',
    {
      viewId: opened.viewId,
      expectedTeachingRevision: opened.project.teachingRevision,
      actions: [{ id: 'speech', type: 'speech', text: 'Exact. 1.25 kg.' }],
      enrichment,
    },
    1,
  ))!.value as any;
  expect(written.enrichment).toEqual(enrichment);
  expect(logic.StartAction).toHaveBeenCalledOnce();
  const oldWriter = replaceNativeTeachingPage({
    documentId: 'document',
    expectedRevision: written.project.teachingRevision,
    pageId: 'page',
    actions: [{ id: 'speech', type: 'speech', text: 'Teacher edit' }],
  });
  expect(oldWriter.enrichment).toEqual(enrichment);
  expect(readNativeTeaching()?.pages[0].actions[0].text).toBe('Teacher edit');
  prepareNativeTeachingSave();
  expect(readNativeTeaching()?.enrichment).toEqual(enrichment);
  expect(() =>
    replaceNativeTeachingPage({
      documentId: 'document',
      expectedRevision: oldWriter.revision,
      pageId: 'page',
      actions: [],
      enrichment: { ...enrichment, originalPages: [{ pageId: 'foreign', pageNumber: 1 }] },
    }),
  ).toThrow('identity');
});

it('rejects unsupported source metadata before changing native history or speech', async () => {
  const { logic } = mount();
  const opened = (await routeNativeTeaching('open', {}, 1))!.value as any;
  await expect(
    routeNativeTeaching(
      'write',
      {
        viewId: opened.viewId,
        expectedTeachingRevision: opened.project.teachingRevision,
        actions: [],
        enrichment: { schemaVersion: 'future', originalPages: [] },
      },
      1,
    ),
  ).rejects.toThrow('enrichment');
  expect(readNativeTeaching()?.pages[0].actions[0].text).toBe('Before');
  expect(logic.StartAction).not.toHaveBeenCalled();
});

it('rejects an assessment insertion for another view before preparing visible changes', async () => {
  const { logic } = mount(false, [{ cSld: { spTree: [] } }]);
  await expect(
    routeNativeTeaching('assessment-insert', { viewId: 'other.page', expectedTeachingRevision: 'revision' }, 1),
  ).rejects.toThrow('view changed');
  expect(readNativeTeaching()?.pages).toHaveLength(1);
  expect(logic.StartAction).not.toHaveBeenCalled();
});

it.each(['import', 'ppt-master', 'scene'])(
  'keeps the %s original body editor when old narration contains a question',
  async (origin) => {
    mount(false, [{ cSld: { spTree: [] } }]);
    const enrichment = {
      schemaVersion: 'deckwise-teaching-enrichment/1' as const,
      originalPages: [{ pageId: 'page', pageNumber: 1 }],
    };
    const actions = [
      { id: 'lecture', type: 'speech', text: `Original ${origin} lecture`, audioAssetId: 'old', durationMs: 1000 },
      {
        id: 'q',
        type: 'interaction_checkpoint',
        checkpointKind: 'chapter_summary',
        checkpointId: 'q',
        chapterId: 'page',
        chapterTitle: 'Practice',
        question: 'Why?',
        response: { kind: 'short_answer', options: [], answerIndices: [] },
        referenceAnswer: { summary: 'Because', keyPoints: [] },
        context: { chapterSummary: 'Practice', sourceSceneIds: ['page'] },
      },
    ];
    replaceNativeTeachingPage({
      documentId: 'document',
      expectedRevision: 'revision',
      pageId: 'page',
      actions,
      enrichment,
    });
    const before = structuredClone(readNativeTeaching());
    const opened = (await routeNativeTeaching('open', {}, 1))!.value as any;
    expect(opened.assessment).toBeUndefined();
    expect(opened.embeddedAssessment).toBeUndefined();
    expect(readNativeTeaching()).toEqual(before);
    expect(opened.project.pages[0].actions[0].text).toBe(`Original ${origin} lecture`);
  },
);

it.each(['quiz', 'checkpoint'] as const)(
  'uses the accepted %s page purpose for short-answer activities after reopen',
  async (kind) => {
    mount(false, [{ cSld: { spTree: [] } }]);
    const enrichment = {
      schemaVersion: 'deckwise-teaching-enrichment/1' as const,
      originalPages: [{ pageId: 'page', pageNumber: 1 }],
      plan: {
        sha256: 'c'.repeat(64),
        units: [{ id: 'review', title: 'Review', evidence: 'Because', kind, pageIds: ['page'], activityId: 'q' }],
      },
    };
    replaceNativeTeachingPage({
      documentId: 'document',
      expectedRevision: 'revision',
      pageId: 'page',
      enrichment,
      actions: [
        { id: 'reading', type: 'speech', text: 'Why?' },
        {
          id: 'q',
          type: 'interaction_checkpoint',
          checkpointKind: 'chapter_summary',
          checkpointId: 'q',
          chapterId: 'page',
          chapterTitle: 'Review',
          question: 'Why?',
          response: { kind: 'short_answer', options: [], answerIndices: [] },
          referenceAnswer: { summary: 'Because', keyPoints: [] },
          context: { chapterSummary: 'Review', sourceSceneIds: ['page'] },
        },
      ],
    });
    for (let reopen = 0; reopen < 2; reopen++) {
      const view = (await routeNativeTeaching('open', {}, 1))!.value as any;
      const projected = kind === 'quiz' ? view.assessment : view.embeddedAssessment;
      expect(projected.classroom.scenes[0].type).toBe(kind === 'quiz' ? 'quiz' : 'slide');
      expect(projected.classroom.scenes[0].actions[0].text).toBe('Why?');
      expect(readNativeTeaching()!.pages).toHaveLength(1);
    }
  },
);

it('inserts an accepted activity as a new physical slide and retains the original lecture and audio', async () => {
  const original = { cSld: { spTree: [] } },
    added = { cSld: { spTree: [] } },
    slides = [original];
  mount(false, slides);
  const enrichment = {
    schemaVersion: 'deckwise-teaching-enrichment/1' as const,
    originalPages: [{ pageId: 'page', pageNumber: 1 }],
    plan: {
      sha256: 'c'.repeat(64),
      units: [
        { id: 'review', title: 'Review', evidence: 'Because', kind: 'checkpoint' as const, pageIds: [] as string[] },
      ],
    },
  };
  const before = replaceNativeTeachingPage({
    documentId: 'document',
    expectedRevision: 'revision',
    pageId: 'page',
    actions: readNativeTeaching()!.pages[0].actions,
    enrichment,
  });
  const originalPage = structuredClone(before.pages[0]);
  const mapped = {
    ...enrichment,
    plan: { ...enrichment.plan, units: [{ ...enrichment.plan.units[0], pageIds: ['activity'], activityId: 'q' }] },
  };
  insertNativeTeachingAssessment({
    documentId: 'document',
    expectedRevision: before.revision,
    afterPageId: 'page',
    pageId: 'activity',
    enrichment: mapped,
    actions: [
      { id: 'reading', type: 'speech', text: 'Why?' },
      {
        id: 'q',
        type: 'interaction_checkpoint',
        checkpointKind: 'chapter_summary',
        checkpointId: 'q',
        chapterId: 'activity',
        chapterTitle: 'Review',
        question: 'Why?',
        response: { kind: 'short_answer', options: [], answerIndices: [] },
        referenceAnswer: { summary: 'Because', keyPoints: [] },
        context: { chapterSummary: 'Review', sourceSceneIds: ['page'] },
      },
    ],
    change: {
      assertCurrent() {},
      apply() {
        slides.splice(1, 0, added);
        return added;
      },
      restore() {
        slides.splice(1, 1);
      },
    },
  });
  expect(slides).toHaveLength(2);
  const saved = readNativeTeaching()!;
  expect(saved.pages.map((p) => p.pageId)).toEqual(['page', 'activity']);
  expect(saved.pages[0]).toEqual(originalPage);
  const opened = (await routeNativeTeaching('open', {}, 2))!.value as any;
  expect(opened.embeddedAssessment.classroom.scenes[1].actions[1].question).toBe('Why?');
});

it('replaces and removes mirrored speech motions without accumulating old native effects or losing manual effects', () => {
  const previous = new Timing();
  previous.effects.push(previous.createEffect('unrelated', 1, 10, 0));
  const slide = {
    cSld: { spTree: [{ sourceId: '4', Id: 'shape-4' }] },
    timing: previous,
    setTiming(value: Timing) {
      this.timing = value;
    },
  };
  mount(false, [slide]);
  const actions = [
    {
      id: 'speech',
      type: 'speech',
      text: 'Before',
      visualCues: [
        {
          id: 'cue',
          atMs: 0,
          endMs: 600,
          effect: { type: 'element_animation', elementId: 'object', effect: 'zoomIn' },
        },
      ],
    },
  ];
  const first = replaceNativeTeachingPage({
    documentId: 'document',
    expectedRevision: 'revision',
    pageId: 'page',
    actions,
  });
  const repaired = structuredClone(actions);
  repaired[0].visualCues[0].effect.effect = 'fadeIn';
  repaired[0].visualCues[0].atMs = 300;
  repaired[0].visualCues[0].endMs = 900;
  const second = replaceNativeTeachingPage({
    documentId: 'document',
    expectedRevision: first.revision,
    pageId: 'page',
    actions: repaired,
  });
  expect(slide.timing.getAllAnimEffects().at(-1)?.asc_getDelay()).toBe(300);
  expect(slide.timing.getAllAnimEffects().map((effect) => [effect.getObjectId(), effect.cTn.presetID])).toEqual([
    ['unrelated', 10],
    ['shape-4', 10],
  ]);
  const third = replaceNativeTeachingPage({
    documentId: 'document',
    expectedRevision: second.revision,
    pageId: 'page',
    actions: [{ id: 'speech', type: 'speech', text: 'Before' }],
  });
  expect(slide.timing.getAllAnimEffects().map((effect) => effect.getObjectId())).toEqual(['unrelated']);
  expect(third.pages[0].actions[0].text).toBe('Before');
});

it('adopts end-of-speech audio after an earlier speech already owns a motion on the same page', async () => {
  vi.stubGlobal('crypto', webcrypto);
  class NativeTiming extends Timing {
    createDuplicate() {
      const copy = new NativeTiming();
      copy.effects = [...this.effects];
      return copy;
    }
    addEffectsToMainSequence(effects: any[]) {
      // The SDK rebuilds its tree and duplicates effects that already have a parent.
      this.effects = [...this.effects.map((e) => ({ ...e, cTn: { ...e.cTn } })), ...effects];
    }
  }
  try {
    const slide = {
      cSld: { spTree: [{ sourceId: '4', Id: 'shape-4' }] },
      timing: new NativeTiming(),
      setTiming(value: NativeTiming) {
        this.timing = value;
      },
    };
    mount(false, [slide]);
    const first = {
      id: 'first',
      type: 'speech',
      text: 'Show',
      visualCues: [
        {
          id: 'reveal',
          atMs: 0,
          endMs: 600,
          effect: { type: 'element_animation', elementId: 'object', effect: 'fadeIn' },
        },
      ],
    };
    const last = {
      id: 'last',
      type: 'speech',
      text: 'Finish',
      attention: { type: 'highlight', objectIds: ['object'], animation: 'fadeOut' },
    };
    const plan = replaceNativeTeachingPage({
      documentId: 'document',
      expectedRevision: 'revision',
      pageId: 'page',
      actions: [first, last],
    });
    const bytes = new TextEncoder().encode('original speech audio').buffer;
    const sha256 = Array.from(new Uint8Array(await crypto.subtle.digest('SHA-256', bytes)), (b) =>
      b.toString(16).padStart(2, '0'),
    ).join('');
    const adopted = await adoptNativeTeachingAudio({
      documentId: 'document',
      pageId: 'page',
      expectedRevision: plan.revision,
      actionId: 'last',
      assetId: 'original',
      sha256,
      durationMs: 1800,
      mimeType: 'audio/mpeg',
      bytes,
    });
    expect(slide.timing.effects.map((e) => [e.cTn.presetClass, e.asc_getDelay()])).toEqual([
      [1, 0],
      [2, 1800],
    ]);
    expect(adopted.pages[0].actions[1]).toMatchObject({ audioAssetId: 'original', durationMs: 1800 });
    replaceNativeTeachingPage({
      documentId: 'document',
      pageId: 'page',
      expectedRevision: adopted.revision,
      actions: [first],
    });
    expect(slide.timing.effects.map((e) => e.cTn.presetClass)).toEqual([1]);
  } finally {
    vi.unstubAllGlobals();
  }
});

it('reuses a matching teacher effect without acquiring it when the speech cue is removed', () => {
  const timing = new Timing();
  const manual = timing.createEffect('shape-4', 1, 10, 0);
  manual.asc_putDuration(600);
  manual.asc_putDelay(300);
  timing.effects.push(manual);
  const slide = {
    cSld: { spTree: [{ sourceId: '4', Id: 'shape-4' }] },
    timing,
    setTiming(value: Timing) {
      this.timing = value;
    },
  };
  mount(false, [slide]);
  const speech = {
    id: 'speech',
    type: 'speech',
    text: 'Before',
    visualCues: [
      {
        id: 'cue',
        atMs: 300,
        endMs: 900,
        effect: { type: 'element_animation', elementId: 'object', effect: 'fadeIn' },
      },
    ],
  };
  const update = (expectedRevision: string, actions: (typeof speech)[]) =>
    replaceNativeTeachingPage({
      documentId: 'document',
      expectedRevision,
      pageId: 'page',
      actions,
    });
  const first = update('revision', [speech]);
  update(first.revision, [{ ...speech, visualCues: [] }]);
  expect(slide.timing.effects).toEqual([manual]);
});

it('adopts existing audio with an eleven-part diagram focus and persists all entrance effects', async () => {
  vi.stubGlobal('crypto', webcrypto);
  try {
    const objects = Array.from({ length: 11 }, (_, i) => ({
      logicalId: `part-${i}`,
      sourceShapeId: String(i + 4),
      type: 'shape',
    }));
    const slide = {
      cSld: { spTree: objects.map((o) => ({ sourceId: o.sourceShapeId, Id: `shape-${o.sourceShapeId}` })) },
      timing: null as Timing | null,
      setTiming(value: Timing | null) {
        this.timing = value;
      },
    };
    mount(false, [slide], objects);
    const teaching = replaceNativeTeachingPage({
      documentId: 'document',
      expectedRevision: 'revision',
      pageId: 'page',
      actions: [
        {
          id: 'speech',
          type: 'speech',
          text: 'Observe the whole chain.',
          attention: { type: 'spotlight', objectIds: objects.map((o) => o.logicalId), animation: 'fadeIn' },
        },
      ],
    });
    const bytes = new TextEncoder().encode('already generated audio').buffer;
    const sha256 = Array.from(new Uint8Array(await crypto.subtle.digest('SHA-256', bytes)), (b) =>
      b.toString(16).padStart(2, '0'),
    ).join('');
    const input = {
      documentId: 'document',
      expectedRevision: teaching.revision,
      pageId: 'page',
      actionId: 'speech',
      assetId: 'existing-audio',
      sha256,
      mimeType: 'audio/mpeg',
      durationMs: 4000,
      bytes,
    };
    const adopted = await adoptNativeTeachingAudio(input);
    expect(adopted.pages[0].actions[0].attention).toBeUndefined();
    expect(adopted.pages[0].actions[0].visualCues).toHaveLength(12);
    expect(slide.timing?.getAllAnimEffects()).toHaveLength(11);
    expect((await adoptNativeTeachingAudio({ ...input, expectedRevision: adopted.revision })).pages[0].actions).toEqual(
      adopted.pages[0].actions,
    );
  } finally {
    vi.unstubAllGlobals();
  }
});

it('saves Quiz title and correct answer edits without invalidating matching speech', async () => {
  mount(false, [{ cSld: { spTree: [] } }]);
  replaceNativeTeachingPage({
    documentId: 'document',
    expectedRevision: 'revision',
    pageId: 'page',
    actions: [
      { id: 'reading', type: 'speech', text: 'Choose\nA. One\nB. Two', audioAssetId: 'old', durationMs: 1000 },
      {
        id: 'q',
        type: 'interaction_checkpoint',
        checkpointKind: 'chapter_summary',
        checkpointId: 'q',
        chapterId: 'page',
        chapterTitle: 'Production instructions',
        question: 'Choose',
        response: { answerIndices: [0], kind: 'single', options: ['One', 'Two'] },
        referenceAnswer: { summary: 'One', keyPoints: ['One'] },
        context: { chapterSummary: 'Count', sourceSceneIds: ['page'] },
      },
    ],
  });
  vi.stubGlobal('crypto', webcrypto);
  const bytes = new TextEncoder().encode('retained question audio').buffer;
  const sha256 = Array.from(new Uint8Array(await crypto.subtle.digest('SHA-256', bytes)), (b) =>
    b.toString(16).padStart(2, '0'),
  ).join('');
  await adoptNativeTeachingAudio({
    documentId: 'document',
    expectedRevision: readNativeTeaching()!.revision,
    pageId: 'page',
    actionId: 'reading',
    assetId: 'old',
    sha256,
    mimeType: 'audio/mpeg',
    durationMs: 1000,
    bytes,
  });
  vi.unstubAllGlobals();
  const opened = (await routeNativeTeaching('assessment-context', {}, 1))!.value as any;
  const classroom = structuredClone(opened.classroom);
  classroom.scenes[0].title = 'Count';
  delete classroom.scenes[0].content.questions[0].keyPoints;
  await routeNativeTeaching(
    'assessment-write',
    {
      projectId: 'document',
      expectedRevision: opened.revision,
      expectedSourceRevision: opened.sourceRevision,
      classroom,
      changedPageIds: ['page'],
    },
    1,
  );
  const reopened = (await routeNativeTeaching('assessment-context', {}, 1))!.value as any;
  expect(reopened.classroom.scenes[0].title).toBe('Count');
  expect((readNativeTeaching()!.pages[0].actions[1].referenceAnswer as any).keyPoints).toEqual(['One']);
  expect(readNativeTeaching()!.pages[0].actions[0]).toMatchObject({
    text: 'Choose\nA. One\nB. Two',
    audioAssetId: 'old',
    durationMs: 1000,
  });
  const answered = structuredClone(reopened.classroom);
  answered.scenes[0].content.questions[0].answer = ['B'];
  await routeNativeTeaching(
    'assessment-write',
    {
      projectId: reopened.projectId,
      expectedRevision: reopened.revision,
      expectedSourceRevision: reopened.sourceRevision,
      classroom: answered,
      changedPageIds: ['page'],
    },
    1,
    true,
  );
  expect(readNativeTeaching()!.pages[0].actions[1].response).toMatchObject({ answerIndices: [1] });
  expect(readNativeTeaching()!.pages[0].actions[0]).toMatchObject({ audioAssetId: 'old', durationMs: 1000 });
});

it('saves a checkpoint page-reference correction without changing accepted speech, audio or reference keys', async () => {
  mount();
  const speech = structuredClone(readNativeTeaching()!.pages[0].actions[0]);
  replaceNativeTeachingPage({
    documentId: 'document',
    expectedRevision: 'revision',
    pageId: 'page',
    actions: [
      speech,
      {
        id: 'q',
        type: 'interaction_checkpoint',
        checkpointKind: 'chapter_summary',
        checkpointId: 'q',
        chapterId: 'page',
        chapterTitle: 'Runoff',
        question: speech.text,
        referenceAnswer: { summary: 'Surface runoff', keyPoints: ['River flows to the sea'] },
        context: { chapterSummary: 'Runoff', sourceSceneIds: ['page'] },
      },
    ],
  });
  for (const hint of ['Return to page 2.', 'Return to page 3.']) {
    const current = (await routeNativeTeaching('assessment-context', {}, 1))!.value as any;
    const classroom = structuredClone(current.classroom);
    const checkpoint = classroom.scenes[0].actions.find((a: any) => a.id === 'q');
    checkpoint.constraintPrompt = hint;
    await routeNativeTeaching(
      'assessment-write',
      {
        projectId: current.projectId,
        expectedRevision: current.revision,
        expectedSourceRevision: current.sourceRevision,
        classroom,
        changedPageIds: ['page'],
      },
      1,
    );
    const saved = readNativeTeaching()!;
    expect(saved.pages[0].actions[0]).toEqual(speech);
    expect(saved.pages[0].actions[1].referenceAnswer).toEqual({
      summary: 'Surface runoff',
      keyPoints: ['River flows to the sea'],
    });
    expect(saved.pages[0].actions[1]).not.toHaveProperty('constraintPrompt');
    const reread = (await routeNativeTeaching('assessment-context', {}, 1))!.value as any;
    expect(reread.classroom.scenes[0].actions.find((a: any) => a.id === 'q').constraintPrompt).toBe(hint);
  }
});

function regionFixture() {
  const original = { Id: 'base', sourceId: '4', getUniNvProps: () => ({ cNvPr: { id: 4, name: 'arrow' } }) };
  const crop = { Id: 'crop', getUniNvProps: () => ({ cNvPr: { id: 5, name: 'region-new' } }) };
  const slide = {
    cSld: { spTree: [original] },
    timing: null as any,
    setTiming(value: any) {
      this.timing = value;
    },
  };
  const f = mount(false, [slide]);
  const data: any = f.data;
  data.generation = {
    pages: [{ pageId: 'page', teaching: { pageType: 'slide', requiredEffects: ['animation', 'spotlight'] } }],
  };
  data.pages[0].actions[0].visualCues = [
    { id: 'fade', atMs: 0, endMs: 600, effect: { type: 'element_animation', effect: 'fadeIn', elementId: 'object' } },
    { id: 'spot', atMs: 600, endMs: 1000, effect: { type: 'spotlight', elementId: 'object', elementIds: [] } },
  ];
  f.parts[0].Delete();
  const xml = document.implementation.createDocument(NATIVE_TEACHING_NAMESPACE, 'teaching');
  xml.documentElement.textContent = JSON.stringify(data);
  f.add(new XMLSerializer().serializeToString(xml));
  const before = structuredClone(readNativeTeaching());
  const change = {
    objectId: 'region-new',
    nativeShape: crop,
    assertCurrent: vi.fn(),
    apply: vi.fn(() => {
      slide.cSld.spTree.unshift(crop as typeof original);
    }),
    restore: vi.fn(() => {
      slide.cSld.spTree = [original];
    }),
  };
  return {
    ...f,
    slide,
    before,
    change,
    run: () =>
      replaceNativeTeachingRegion({
        documentId: 'document',
        expectedRevision: 'revision',
        pageId: 'page',
        actionId: 'speech',
        description: 'River region',
        change,
      }),
  };
}

it('commits region, timing and carrier together while retaining audio and original foreground IDs', () => {
  const f = regionFixture();
  const saved = f.run();
  const speech = saved.pages[0].actions[0],
    previous = f.before!.pages[0].actions[0];
  const { visualCues, ...unchanged } = speech;
  const { visualCues: oldCues, ...expected } = previous;
  expect(unchanged).toEqual(expected);
  expect((visualCues as any[])[0]).toEqual((oldCues as any[])[0]);
  expect((visualCues as any[])[1].effect.elementIds).toEqual(['region-new']);
  expect((visualCues as any[])[2].effect.elementId).toBe('region-new');
  expect(nativeTeachingObjects(1).map((o) => o.objectId)).toEqual(['region-new', 'object']);
  expect(f.slide.timing.getAllAnimEffects().map((e: any) => e.getObjectId())).toContain('crop');
  expect(f.logic.StartAction).toHaveBeenCalledOnce();
  expect(f.logic.FinalizeAction).toHaveBeenCalledOnce();
});

it('rolls back the background/crop and carrier when the compound region write fails', () => {
  const f = regionFixture();
  f.change.apply.mockImplementation(() => {
    f.slide.cSld.spTree.unshift(f.change.nativeShape as any);
    throw new Error('crop insertion failed');
  });
  expect(f.run).toThrow('crop insertion failed');
  expect(f.change.restore).toHaveBeenCalledOnce();
  expect(readNativeTeaching()).toEqual(f.before);
  expect(nativeTeachingObjects(1).map((o) => o.objectId)).toEqual(['object']);
  expect(f.slide.timing).toBeNull();
});

it('rejects stale region writes before changing the carrier or native image', () => {
  const f = regionFixture();
  replaceNativeTeachingPage({
    documentId: 'document',
    expectedRevision: 'revision',
    pageId: 'page',
    actions: f.before!.pages[0].actions,
  });
  expect(f.run).toThrow('revision changed');
  expect(f.change.apply).not.toHaveBeenCalled();
});

it('blocks a second crop on the same page after a successful region repair', () => {
  const f = regionFixture();
  const saved = f.run();
  expect(() =>
    replaceNativeTeachingRegion({
      documentId: 'document',
      expectedRevision: saved.revision,
      pageId: 'page',
      actionId: 'speech',
      description: 'Repeated crop',
      change: { ...f.change, objectId: 'region-second' },
    }),
  ).toThrow('already repaired');
  expect(f.change.apply).toHaveBeenCalledOnce();
});
