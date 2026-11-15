import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { DeckWiseEditorBridge } from '../../lib/deckwise-editor-bridge';
import { canSetSlideFade, prepareSlideFade } from '../../lib/agent-plugin/slide-fade';
import { setReadonlyState } from '../../lib/onlyoffice/readonly';

vi.mock('../../lib/agent-plugin/slide-fade', () => ({ canSetSlideFade: vi.fn(), prepareSlideFade: vi.fn() }));
const owner = { projectId: 'project', documentId: 'document', sessionId: 'session' };
const result = {
  page: 1,
  transition: 'fade' as const,
  durationMs: 750,
  verified: true as const,
  saved: false as const,
};

function setup() {
  const history = { Index: -1, Points: [] as Array<{ Items: unknown[] }> };
  let page = 0;
  const selectedObjects: unknown[] = [];
  const logic = {
    Slides: [{ cSld: { spTree: [{ getText: () => 'Lesson' }] } }],
    GetSelectedSlides: () => [page],
    GetCurrentController: () => ({ selectedObjects }),
  };
  const api = {
    isDocumentLoadComplete: true,
    isLoadFullApi: true,
    getCurrentPage: () => page,
    WordControl: { m_oLogicDocument: logic },
  };
  const frame = document.createElement('iframe');
  frame.name = 'frameEditor';
  document.body.append(frame);
  // Keep the controller identity stable, as the actual editor does.
  const controller = { selectedObjects };
  logic.GetCurrentController = () => controller;
  Object.assign(frame.contentWindow!, { editor: api, AscCommon: { History: history } });
  const bridge = new DeckWiseEditorBridge();
  bridge.reset(owner);
  const context = bridge.getContext(owner);
  const proof = {
    editGeneration: context.editGeneration,
    page: 1,
    sourceHash: `sha256:${'a'.repeat(64)}`,
    slideHash: `sha256:${'b'.repeat(64)}`,
  };
  const command = {
    ...owner,
    requestId: 'fade-request',
    editGeneration: context.editGeneration,
    tool: 'set_slide_fade',
    input: {},
  };
  return { bridge, context, proof, command, history, selectedObjects, navigate: () => page++ };
}

describe('Host-inspected native fade bridge', () => {
  beforeEach(() => {
    setReadonlyState(false);
    vi.mocked(canSetSlideFade).mockReturnValue(true);
    vi.mocked(prepareSlideFade).mockReturnValue(Object.freeze({ page: 1, apply: vi.fn(() => ({ ...result })) }));
  });
  afterEach(() => {
    setReadonlyState(false);
    document.body.replaceChildren();
    vi.clearAllMocks();
  });

  it('withholds SDK-ready fade until exact source-bound Host review and keeps proof private', () => {
    const { bridge, context, proof } = setup();
    expect(context.capabilities.some((c) => c.name === 'set_slide_fade')).toBe(false);
    const approved = bridge.reviewFade(owner, proof);
    expect(approved.editGeneration).toBe(context.editGeneration);
    expect(approved.capabilities.find((c) => c.name === 'set_slide_fade')).toMatchObject({
      readOnly: false,
      inputSchema: { additionalProperties: false, properties: {} },
    });
    expect(JSON.stringify(approved)).not.toContain(proof.sourceHash);
    expect(JSON.stringify(approved)).not.toContain(proof.slideHash);
  });

  it('dispatches the fixed native edit once through the dedicated Host channel', async () => {
    const { bridge, proof, command } = setup();
    bridge.reviewFade(owner, proof);
    const answer = await bridge.executeFade(command, proof, command.requestId);
    expect(answer).toEqual({
      ...owner,
      requestId: command.requestId,
      editGeneration: 1,
      status: 'verified',
      tool: 'set_slide_fade',
      result,
    });
    expect(prepareSlideFade).toHaveBeenCalledOnce();
    await expect(bridge.executeFade(command, proof, command.requestId)).rejects.toThrow();
    expect(bridge.getContext(owner).capabilities.some((c) => c.name === 'set_slide_fade')).toBe(false);
  });

  it('rejects ordinary execution even after review and consumes that attempted grant', async () => {
    const { bridge, proof, command } = setup();
    bridge.reviewFade(owner, proof);
    await expect(bridge.execute(command, command.requestId)).rejects.toThrow(/host fade authorization/);
    expect(prepareSlideFade).not.toHaveBeenCalled();
    await expect(bridge.executeFade(command, proof, command.requestId)).rejects.toThrow();
  });

  it.each([
    {},
    { editGeneration: 0 },
    { editGeneration: 2 },
    { page: 0 },
    { page: 2 },
    { sourceHash: 'a'.repeat(64) },
    { slideHash: 'sha256:BAD' },
    { extra: true },
  ])('rejects invalid or stale authorization %j', (patch) => {
    const { bridge, proof } = setup();
    const value = Object.keys(patch).length ? { ...proof, ...patch } : {};
    expect(() => bridge.reviewFade(owner, value)).toThrow();
    expect(bridge.getContext(owner).capabilities.some((c) => c.name === 'set_slide_fade')).toBe(false);
  });

  it.each(['page', 'history', 'selection', 'readonly', 'capability', 'revoke', 'reset'])(
    'refuses a reviewed grant after %s changes',
    async (change) => {
      const s = setup();
      s.bridge.reviewFade(owner, s.proof);
      if (change === 'page') s.navigate();
      if (change === 'history') s.history.Index++;
      if (change === 'selection') s.selectedObjects.push({});
      if (change === 'readonly') setReadonlyState(true);
      if (change === 'capability') vi.mocked(canSetSlideFade).mockReturnValue(false);
      if (change === 'revoke') s.bridge.revokeFade();
      if (change === 'reset') s.bridge.reset(owner);
      // The helper itself is mocked here, so native readiness is checked by the bridge too.

      await expect(s.bridge.executeFade(s.command, s.proof, s.command.requestId)).rejects.toThrow();
      expect(prepareSlideFade).not.toHaveBeenCalled();
    },
  );

  it.each([
    { page: 2 },
    { sourceHash: `sha256:${'c'.repeat(64)}` },
    { slideHash: `sha256:${'d'.repeat(64)}` },
    { editGeneration: 2 },
  ])('rejects a different transport proof %j', async (patch) => {
    const s = setup();
    s.bridge.reviewFade(owner, s.proof);
    await expect(s.bridge.executeFade(s.command, { ...s.proof, ...patch }, s.command.requestId)).rejects.toThrow();
    expect(prepareSlideFade).not.toHaveBeenCalled();
  });

  it.each([{ durationMs: 1000 }, { page: 2 }, { script: 'bad' }, { sourceHash: 'bad' }])(
    'rejects model effect/identity fields %j',
    async (input) => {
      const s = setup();
      s.bridge.reviewFade(owner, s.proof);
      await expect(s.bridge.executeFade({ ...s.command, input }, s.proof, s.command.requestId)).rejects.toThrow();
      expect(prepareSlideFade).not.toHaveBeenCalled();
    },
  );

  it('pins reviewed proof, allows explicit revoke, and rejects uncorrelated request IDs', async () => {
    const s = setup();
    const original = { ...s.proof };
    s.bridge.reviewFade(owner, s.proof);
    s.proof.sourceHash = `sha256:${'c'.repeat(64)}`;
    await expect(s.bridge.executeFade(s.command, original, 'other')).rejects.toThrow(/command/);
    s.bridge.reviewFade(owner, null);
    await expect(s.bridge.executeFade(s.command, original, s.command.requestId)).rejects.toThrow();
    expect(prepareSlideFade).not.toHaveBeenCalled();
  });

  it('consumes uncertain native execution and never retries or saves', async () => {
    const s = setup();
    s.bridge.reviewFade(owner, s.proof);
    const apply = vi.fn(() => {
      throw new Error('Native readback uncertain; inspect and Undo');
    });
    vi.mocked(prepareSlideFade).mockReturnValue(Object.freeze({ page: 1, apply }));
    await expect(s.bridge.executeFade(s.command, s.proof, s.command.requestId)).rejects.toThrow(/uncertain/);
    await expect(s.bridge.executeFade(s.command, s.proof, s.command.requestId)).rejects.toThrow();
    expect(apply).toHaveBeenCalledOnce();
    expect(s.bridge.isExecuting).toBe(false);
  });

  it('copies the Host proof before the caller can mutate the review object', async () => {
    const s = setup(),
      original = { ...s.proof };
    s.bridge.reviewFade(owner, s.proof);
    s.proof.sourceHash = `sha256:${'c'.repeat(64)}`;
    await expect(s.bridge.executeFade(s.command, original, s.command.requestId)).resolves.toMatchObject({ result });
  });

  it('rejects mismatched prepared native page before applying', async () => {
    const s = setup();
    s.bridge.reviewFade(owner, s.proof);
    const apply = vi.fn(() => ({ ...result }));
    vi.mocked(prepareSlideFade).mockReturnValue(Object.freeze({ page: 2, apply }));
    await expect(s.bridge.executeFade(s.command, s.proof, s.command.requestId)).rejects.toThrow(/page changed/);
    expect(apply).not.toHaveBeenCalled();
  });
});
