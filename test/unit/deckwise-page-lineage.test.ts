import { File as NodeFile } from 'node:buffer';
import { webcrypto, createHash } from 'node:crypto';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { DeckWisePageLineage } from '../../lib/deckwise-page-lineage';

const owner = { projectId: 'workspace', documentId: 'document', sessionId: 'session' };
const file = (text: string) => new NodeFile([text], 'lesson.pptx') as unknown as File;
const hash = (text: string) => createHash('sha256').update(text).digest('hex');
function mount() {
  const frame = document.createElement('iframe');
  frame.name = 'frameEditor';
  document.body.append(frame);
  const logic = { Slides: [{}, {}, {}] };
  const api = { isDocumentLoadComplete: true, isLoadFullApi: true, WordControl: { m_oLogicDocument: logic } };
  Object.assign(frame.contentWindow!, { editor: api });
  return { frame, api, logic };
}

describe('session-bound native page lineage', () => {
  beforeEach(() => {
    vi.stubGlobal('crypto', webcrypto);
  });
  afterEach(() => {
    document.body.replaceChildren();
    vi.unstubAllGlobals();
  });

  async function opened() {
    const lineage = new DeckWisePageLineage();
    await lineage.begin(owner, file('source'));
    const sdk = mount();
    lineage.contentReady();
    return { lineage, ...sdk };
  }

  it('binds page references at content-ready and hashes actual save bytes', async () => {
    const { lineage } = await opened();
    const context = lineage.context()!;
    expect(context.sourceSha256).toBe(hash('source'));
    expect(context.pageTokens).toEqual(context.openedPageTokens);
    expect(new Set(context.pageTokens).size).toBe(3);
    const capture = lineage.captureSave()!;
    expect(await capture(file('edited-output'), 'save-1', () => true)).toEqual({
      protocol: 'deckwise-office-page-lineage/1',
      ...owner,
      saveRequestId: 'save-1',
      ...context,
      outputSha256: hash('edited-output'),
      objects: [[], [], []], animationUnchanged: [true, true, true],
    });
    expect(lineage.context()).toEqual(context);
  });

  it('accepts the saved baseline while newer page edits remain for the next export', async () => {
    const { lineage, logic } = await opened();
    const original = lineage.context()!;
    const receipt = (await lineage.captureSave()!(file('output'), 'save', () => true))!;
    logic.Slides.reverse();
    expect(lineage.commitSave(receipt)).toBe(true);
    expect(lineage.context()!.sourceSha256).toBe(hash('output'));
    expect(lineage.context()!.openedPageTokens).toEqual(original.pageTokens);
    expect(lineage.context()!.pageTokens).toEqual([...original.pageTokens].reverse());
    const next = (await lineage.captureSave()!(file('second'), 'save-2', () => true))!;
    expect(next.sourceSha256).toBe(hash('output'));
    expect(lineage.commitSave(next)).toBe(true);
  });

  it('tracks reorder by object identity and gives copies a fresh token', async () => {
    const { lineage, logic } = await opened();
    const original = lineage.context()!;
    const [first, second, third] = logic.Slides;
    logic.Slides = [third, { ...first }, first];
    const moved = lineage.context()!;
    expect(moved.pageTokens[0]).toBe(original.pageTokens[2]);
    expect(moved.pageTokens[2]).toBe(original.pageTokens[0]);
    expect(original.pageTokens).not.toContain(moved.pageTokens[1]);
    expect(moved.openedPageTokens).toEqual(original.openedPageTokens);
    logic.Slides.push(second);
    expect(lineage.context()!.pageTokens[3]).toBe(original.pageTokens[1]);
  });

  it.each(['history', 'reorder', 'replace', 'reset', 'frame', 'document'] as const)(
    'withholds a save receipt after %s changes during export',
    async (change) => {
      const { lineage, logic, frame, api } = await opened();
      const capture = lineage.captureSave()!;
      if (change === 'reorder') logic.Slides.reverse();
      if (change === 'replace') logic.Slides[0] = {};
      if (change === 'reset') lineage.reset();
      if (change === 'frame') frame.remove();
      if (change === 'document') api.WordControl.m_oLogicDocument = { Slides: [{}] };
      expect(await capture(file('output'), 'save-1', () => change !== 'history')).toBeUndefined();
    },
  );

  it('does not bind an old editor or late content-ready after reset', async () => {
    const lineage = new DeckWisePageLineage();
    mount();
    await lineage.begin(owner, file('source'));
    lineage.contentReady();
    expect(lineage.context()).toBeUndefined();
    await lineage.begin(owner, file('source'));
    lineage.reset();
    document.body.replaceChildren();
    mount();
    lineage.contentReady();
    expect(lineage.captureSave()).toBeUndefined();
  });

  it('rejects repeated references and never reuses tokens across reopen', async () => {
    const { lineage, logic } = await opened();
    const old = lineage.context()!;
    logic.Slides.push(logic.Slides[0]);
    expect(lineage.context()).toBeUndefined();
    await lineage.begin({ ...owner, sessionId: 'reopen' }, file('output'));
    document.body.replaceChildren();
    mount();
    lineage.contentReady();
    const reopened = lineage.context()!;
    expect(reopened.sourceSha256).toBe(hash('output'));
    expect(reopened.pageTokens.some((token) => old.pageTokens.includes(token))).toBe(false);
  });

  it('ignores a superseded open callback without consuming the newer baseline', async () => {
    const lineage = new DeckWisePageLineage();
    const oldReady = await lineage.begin(owner, file('old'));
    const newReady = await lineage.begin({ ...owner, sessionId: 'new' }, file('new'));
    mount();
    oldReady();
    expect(lineage.context()).toBeUndefined();
    newReady();
    expect(lineage.context()!.sourceSha256).toBe(hash('new'));
  });

  it('captures binding values before asynchronous hashing', async () => {
    const lineage = new DeckWisePageLineage();
    const binding = { ...owner };
    const pending = lineage.begin(binding, file('source'));
    binding.sessionId = 'changed';
    const ready = await pending;
    mount();
    ready();
    const receipt = await lineage.captureSave()!(file('output'), 'save', () => true);
    expect(receipt?.sessionId).toBe(owner.sessionId);
  });

  it('copies outward arrays and rejects malformed or superseded opens', async () => {
    const { lineage } = await opened();
    lineage.context()!.pageTokens.reverse();
    expect(lineage.context()!.pageTokens).toEqual(lineage.context()!.openedPageTokens);
    await expect(lineage.begin({ ...owner, sessionId: '' }, file('source'))).rejects.toThrow(/Invalid/);
    const pending = lineage.begin(owner, file('source'));
    lineage.reset();
    await expect(pending).rejects.toThrow(/superseded/);
  });
});

it('keeps object identity across moves and renews the source only after a successful save', async () => {
  vi.stubGlobal('crypto', webcrypto);
  const lineage = new DeckWisePageLineage();
  await lineage.begin(owner, file('source'));
  const sdk = mount();
  const shape = { getUniNvProps: () => ({ cNvPr: { name: 'Panda' } }) };
  Object.assign(sdk.logic.Slides[0], { cSld: { spTree: [shape] } });
  lineage.contentReady();
  const capture = lineage.captureSave()!;
  const receipt = await capture(file('saved'), 'save-1', () => true);
  expect(receipt?.objects[0][0].sourceName).toBe('Panda');
  expect(lineage.context()!.sourceSha256).toBe(hash('source'));
  lineage.commitSave(receipt!);
  expect(lineage.context()!.sourceSha256).toBe(hash('saved'));
  document.body.replaceChildren();
  vi.unstubAllGlobals();
});
