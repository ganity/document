import { createHash, webcrypto } from 'node:crypto';
import { afterEach, beforeEach, expect, it, vi } from 'vitest';
import { replaceSlideImageFromAsset, insertSlideImageFromAsset } from '../../lib/agent-plugin/slide-image-asset';
import { captureDocumentToolTarget } from '../../lib/agent-plugin/document-tool-action';
import { prepareSlideImageReplacement } from '../../lib/agent-plugin/slide-image-replacement';
import { prepareSlideImageInsertion } from '../../lib/agent-plugin/slide-image-insertion';
vi.mock('../../lib/agent-plugin/slide-image-insertion', () => ({ prepareSlideImageInsertion: vi.fn() }));

vi.mock('../../lib/agent-plugin/document-tool-action', () => ({ captureDocumentToolTarget: vi.fn() }));
vi.mock('../../lib/agent-plugin/slide-image-replacement', () => ({ prepareSlideImageReplacement: vi.fn() }));
const png = Uint8Array.from(
  Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=', 'base64'),
);
const sha256 = createHash('sha256').update(png).digest('hex');
function setup() {
  const binding = { projectId: 'project', documentId: 'document', sessionId: 'session', editGeneration: 4 };
  const current = { ...binding, readOnly: false, saving: false, imageReplacementInstalled: true };
  const asset = { assetId: 'asset', sha256, mime: 'image/png' as const, byteLength: png.length, width: 1, height: 1 };
  const isCurrent = vi.fn(() => true);
  vi.mocked(captureDocumentToolTarget).mockReturnValue({
    context: { kind: 'slide' },
    label: 'page 2',
    selectedText: '',
    isCurrent,
  });
  const apply = vi.fn(() => ({ page: 2, imageId: 'image', verified: true as const }));
  vi.mocked(prepareSlideImageReplacement).mockResolvedValue({ page: 2, imageId: 'image', apply });
  const loadAsset = vi.fn(async () => png.slice());
  return { binding, asset, current, apply, isCurrent, loadAsset, readCurrent: () => current };
}
beforeEach(() => vi.stubGlobal('crypto', webcrypto));
afterEach(() => {
  vi.clearAllMocks();
  vi.unstubAllGlobals();
});

it('loads only the Host-selected asset, verifies actual bytes, and applies once without saving', async () => {
  const s = setup();
  const signal = new AbortController().signal;
  await expect(replaceSlideImageFromAsset({ ...s, signal })).resolves.toEqual({
    page: 2,
    imageId: 'image',
    verified: true,
    assetId: 'asset',
    sha256,
    saved: false,
  });
  expect(s.loadAsset).toHaveBeenCalledExactlyOnceWith('asset', signal);
  expect(prepareSlideImageReplacement).toHaveBeenCalledExactlyOnceWith(png, signal);
  expect(s.apply).toHaveBeenCalledTimes(1);
});

it.each(['binding', 'generation', 'mime', 'hash', 'length', 'dimensions'])(
  'rejects invalid %s before loading or preloading',
  async (change) => {
    const s = setup();
    if (change === 'binding') s.binding.projectId = '';
    if (change === 'generation') s.binding.editGeneration = 0;
    if (change === 'mime') Reflect.set(s.asset, 'mime', 'image/jpeg');
    if (change === 'hash') s.asset.sha256 = 'bad';
    if (change === 'length') s.asset.byteLength = 4 * 1024 * 1024 + 1;
    if (change === 'dimensions') s.asset.width = 4097;
    await expect(replaceSlideImageFromAsset(s)).rejects.toThrow(/pinned PNG/);
    expect(s.loadAsset).not.toHaveBeenCalled();
    expect(prepareSlideImageReplacement).not.toHaveBeenCalled();
  },
);

it.each(['projectId', 'documentId', 'sessionId', 'editGeneration', 'readOnly', 'saving', 'capability', 'selection'])(
  'rejects a late %s change during asset retrieval without preload or apply',
  async (change) => {
    const s = setup();
    s.loadAsset.mockImplementation(async () => {
      if (change === 'projectId' || change === 'documentId' || change === 'sessionId') s.current[change] = 'other';
      if (change === 'editGeneration') s.current.editGeneration++;
      if (change === 'readOnly') s.current.readOnly = true;
      if (change === 'saving') s.current.saving = true;
      if (change === 'capability') s.current.imageReplacementInstalled = false;
      if (change === 'selection') s.isCurrent.mockReturnValue(false);
      return png.slice();
    });
    await expect(replaceSlideImageFromAsset(s)).rejects.toThrow(/context changed/);
    expect(prepareSlideImageReplacement).not.toHaveBeenCalled();
    expect(s.apply).not.toHaveBeenCalled();
  },
);

it.each(['length', 'bytes', 'dimensions'])('rejects downloaded %s mismatch before native preload', async (change) => {
  const s = setup();
  if (change === 'length') s.loadAsset.mockResolvedValue(png.slice(0, -1));
  if (change === 'bytes') {
    const corrupt = png.slice();
    corrupt[30] ^= 1;
    s.loadAsset.mockResolvedValue(corrupt);
  }
  if (change === 'dimensions') s.asset.width = 2;
  await expect(replaceSlideImageFromAsset(s)).rejects.toThrow(/mismatch/);
  expect(prepareSlideImageReplacement).not.toHaveBeenCalled();
  expect(s.apply).not.toHaveBeenCalled();
});

it('captures metadata and binding before awaiting a resolver', async () => {
  const s = setup();
  s.loadAsset.mockImplementation(async () => {
    s.asset.assetId = 'other';
    s.asset.sha256 = 'bad';
    s.binding.projectId = 'other';
    return png.slice();
  });
  const result = await replaceSlideImageFromAsset(s);
  expect(result.assetId).toBe('asset');
  expect(result.sha256).toBe(sha256);
});

it('owns downloaded bytes across asynchronous hashing', async () => {
  const s = setup();
  const bytes = png.slice();
  s.loadAsset.mockResolvedValue(bytes);
  const digest = webcrypto.subtle.digest.bind(webcrypto.subtle);
  const hash = vi.fn(async (algorithm: string, value: Parameters<typeof digest>[1]) => {
    bytes.fill(0);
    return digest(algorithm, value);
  });
  vi.stubGlobal('crypto', { subtle: { digest: hash } });
  await replaceSlideImageFromAsset(s);
  expect(prepareSlideImageReplacement).toHaveBeenCalledWith(png, undefined);
});

it.each(['hash', 'preload'])('rejects selection changes during %s before applying', async (stage) => {
  const s = setup();
  if (stage === 'hash') {
    const digest = webcrypto.subtle.digest.bind(webcrypto.subtle);
    vi.stubGlobal('crypto', {
      subtle: {
        digest: async (algorithm: string, value: Parameters<typeof digest>[1]) => {
          s.isCurrent.mockReturnValue(false);
          return digest(algorithm, value);
        },
      },
    });
  } else
    vi.mocked(prepareSlideImageReplacement).mockImplementation(async () => {
      s.isCurrent.mockReturnValue(false);
      return { page: 2, imageId: 'image', apply: s.apply };
    });
  await expect(replaceSlideImageFromAsset(s)).rejects.toThrow(/context changed/);
  expect(s.apply).not.toHaveBeenCalled();
});

it('does not retry failed authorization, cancellation or an uncertain native command', async () => {
  const s = setup();
  s.loadAsset.mockRejectedValueOnce(new Error('TEAM_PERMISSION_DENIED'));
  await expect(replaceSlideImageFromAsset(s)).rejects.toThrow('TEAM_PERMISSION_DENIED');
  expect(s.loadAsset).toHaveBeenCalledTimes(1);
  expect(prepareSlideImageReplacement).not.toHaveBeenCalled();
  const abort = new AbortController();
  s.loadAsset.mockImplementationOnce(async () => {
    abort.abort();
    return png;
  });
  await expect(replaceSlideImageFromAsset({ ...s, signal: abort.signal })).rejects.toThrow();
  expect(s.apply).not.toHaveBeenCalled();
  s.apply.mockImplementationOnce(() => {
    throw new Error('unverified');
  });
  await expect(replaceSlideImageFromAsset(s)).rejects.toThrow('unverified');
  expect(s.apply).toHaveBeenCalledTimes(1);
});

const box = { xMm: 25, yMm: 35, widthMm: 40, heightMm: 30 };
function insertion() {
  const s = setup();
  const current = { ...s.current, imageInsertionInstalled: true };
  const apply = vi.fn(() => ({ page: 2, imageId: 'inserted', box, verified: true as const, saved: false as const }));
  vi.mocked(prepareSlideImageInsertion).mockResolvedValue({ page: 2, box, apply });
  return { ...s, apply, box: { ...box }, current, readCurrent: () => current };
}
it('pins authorized insertion bytes, explicit geometry and receipt without using replacement', async () => {
  const s = insertion();
  const result = await insertSlideImageFromAsset(s);
  expect(result).toMatchObject({ ...box, assetId: 'asset', sha256, page: 2, imageId: 'inserted', saved: false });
  expect(prepareSlideImageInsertion).toHaveBeenCalledExactlyOnceWith(png, box, undefined);
  expect(prepareSlideImageReplacement).not.toHaveBeenCalled();
  expect(s.apply).toHaveBeenCalledTimes(1);
});
it.each(['asset', 'hash', 'geometry', 'reader', 'loader'])('pins insertion %s across awaits', async (change) => {
  const s = insertion();
  let finish!: (value: ReturnType<typeof png.slice>) => void;
  s.loadAsset.mockImplementation(
    () =>
      new Promise((resolve) => {
        finish = resolve;
      }),
  );
  const pending = insertSlideImageFromAsset(s);
  if (change === 'asset') s.asset.assetId = 'different';
  if (change === 'hash') s.asset.sha256 = 'b'.repeat(64);
  if (change === 'geometry') s.box.xMm = 100;
  if (change === 'reader')
    s.readCurrent = () => {
      throw new Error('substituted reader');
    };
  if (change === 'loader')
    s.loadAsset = vi.fn(async () => {
      throw new Error('substituted loader');
    });
  finish(png.slice());
  await expect(pending).resolves.toMatchObject({ ...box, assetId: 'asset', sha256 });
});
it.each(['generation', 'readonly', 'saving', 'capability', 'selection', 'cancel'])(
  'refuses stale insertion %s',
  async (change) => {
    const s = insertion(),
      abort = new AbortController();
    s.loadAsset.mockImplementation(async () => {
      if (change === 'generation') s.current.editGeneration++;
      if (change === 'readonly') s.current.readOnly = true;
      if (change === 'saving') s.current.saving = true;
      if (change === 'capability') s.current.imageInsertionInstalled = false;
      if (change === 'selection') s.isCurrent.mockReturnValue(false);
      if (change === 'cancel') abort.abort();
      return png.slice();
    });
    await expect(insertSlideImageFromAsset({ ...s, signal: abort.signal })).rejects.toThrow();
    expect(prepareSlideImageInsertion).not.toHaveBeenCalled();
    expect(s.apply).not.toHaveBeenCalled();
  },
);
it.each([
  { ...box, xMm: -1 },
  { ...box, heightMm: 0 },
  { ...box, widthMm: NaN },
  { ...box, yMm: 1001 },
  { ...box, url: 'bad' },
])('refuses invalid explicit box %j before asset loading', async (value) => {
  const s = insertion();
  await expect(insertSlideImageFromAsset({ ...s, box: value })).rejects.toThrow(/geometry/);
  expect(s.loadAsset).not.toHaveBeenCalled();
});
it('rechecks insertion generation after native preload', async () => {
  const s = insertion();
  vi.mocked(prepareSlideImageInsertion).mockImplementation(async () => {
    s.current.editGeneration++;
    return { page: 2, box, apply: s.apply };
  });
  await expect(insertSlideImageFromAsset(s)).rejects.toThrow(/changed/);
  expect(s.apply).not.toHaveBeenCalled();
});
