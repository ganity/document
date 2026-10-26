import { afterEach, beforeEach, expect, it, vi } from 'vitest';
import { canReplaceSlideImage, prepareSlideImageReplacement } from '../../lib/agent-plugin/slide-image-replacement';
import { setReadonlyState } from '../../lib/onlyoffice/readonly';
import { DeckWiseEditorBridge } from '../../lib/deckwise-editor-bridge';
const png = Uint8Array.from(
  Buffer.from(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR4nGP45Zr+HwAGiQKmz1SUAQAAAABJRU5ErkJggg==',
    'base64',
  ),
);
function mount() {
  const history = { Index: -1, Points: [] as Array<{ Items: unknown[] }> };
  const image = {
    Id: 'image-1',
    group: null,
    isPlaceholder: () => false,
    getObjectType: () => 72548352,
    blipFill: { RasterImageId: 'old.png', srcRect: null as unknown },
    spPr: {
      geometry: { preset: 'rect' },
      xfrm: { offX: 10, offY: 10, extX: 20, extY: 20, rot: null as number | null, flipH: false, flipV: false },
    },
  };
  const slide = { cSld: { spTree: [image] } };
  const controller = { selectedObjects: [image], getTargetDocContent: () => undefined };
  const logic = {
    Slides: [slide],
    CanEdit: () => true,
    IsMasterMode: () => false,
    GetWidthMM: () => 320,
    GetHeightMM: () => 200,
    GetCurrentController: () => controller,
    GetSelectedSlides: () => [0],
  };
  const api = {
    isDocumentLoadComplete: true,
    isLoadFullApi: true,
    isViewMode: false,
    canEdit: () => true,
    isSlideShow: () => false,
    isLongAction: () => false,
    getCurrentPage: () => 0,
    WordControl: { m_oLogicDocument: logic },
    ImageLoader: {
      LoadImagesWithCallback: vi.fn((_urls: string[], callback: () => void) => callback()),
      LoadImage: vi.fn(() => ({ Image: { naturalWidth: 1, naturalHeight: 1 } })),
    },
    ImgApply: vi.fn((_props: { ImageUrl: string }) => {
      image.blipFill.RasterImageId = 'new.png';
      history.Index++;
      history.Points.push({ Items: [{}] });
    }),
  };
  const common = {
    History: history,
    CollaborativeEditing: { Get_GlobalLock: () => false },
    sendImgUrls: vi.fn(
      (_api: unknown, _urls: string[], callback: (value: Array<{ path: string; url: string }>) => void) =>
        callback([{ path: 'media/new.png', url: 'blob:local' }]),
    ),
    getFullImageSrc2: () => 'blob:local',
    g_oDocumentUrls: { imagePath2Local: () => 'new.png', getImageLocal: () => 'new.png' },
  };
  const frame = document.createElement('iframe');
  document.body.append(frame);
  Object.assign(frame.contentWindow!, {
    editor: api,
    Asc: {},
    AscCommon: common,
    AscDFH: { historyitem_type_ImageShape: 72548352 },
  });
  return { api, image, slide, logic, controller, common, history, frame };
}
beforeEach(() => setReadonlyState(false));
afterEach(() => {
  setReadonlyState(false);
  document.body.replaceChildren();
  vi.restoreAllMocks();
  vi.useRealTimers();
});
it('preloads local PNG without mutation and applies once through the native command', async () => {
  const state = mount();
  expect(canReplaceSlideImage()).toBe(true);
  expect(state.common.sendImgUrls).not.toHaveBeenCalled();
  const proposal = await prepareSlideImageReplacement(png);
  expect(state.api.ImgApply).not.toHaveBeenCalled();
  expect(proposal.apply()).toEqual({ page: 1, imageId: 'image-1', verified: true });
  expect(state.api.ImgApply).toHaveBeenCalledExactlyOnceWith({ ImageUrl: 'blob:local' });
  expect(() => proposal.apply()).toThrow(/consumed/);
});
it.each([
  [Math.PI / 4, false, false],
  [Math.PI / 2, true, false],
  [Math.PI, false, true],
  [Math.PI * 2, true, true],
  [0, true, false],
  [0, false, true],
  [0, true, true],
] as const)('preserves rotation %s and flips %s/%s', async (rotation, horizontal, vertical) => {
  const s = mount();
  Object.assign(s.image.spPr.xfrm, { rot: rotation, flipH: horizontal, flipV: vertical });
  const geometry = { ...s.image.spPr.xfrm };
  const prepared = await prepareSlideImageReplacement(png);
  expect(prepared.apply().verified).toBe(true);
  expect(s.image.spPr.xfrm).toEqual(geometry);
  expect(s.api.ImgApply).toHaveBeenCalledExactlyOnceWith({ ImageUrl: 'blob:local' });
});
it('accepts a right-angle image whose rotated bounds touch the slide edge', async () => {
  const s = mount();
  Object.assign(s.image.spPr.xfrm, { offX: 0, rot: Math.PI / 2 });
  const prepared = await prepareSlideImageReplacement(png);
  expect(prepared.apply().verified).toBe(true);
});
it.each([
  'readonly',
  'view',
  'crop',
  'invalid-rotation',
  'invalid-flip',
  'rotated-offpage',
  'group',
  'placeholder',
  'wrong-type',
  'multiple',
  'foreign',
  'offpage',
  'nonfinite',
  'mask',
  'tile',
  'transparency',
  'busy',
  'slideshow',
  'master',
  'locked',
  'unloaded',
  'untracked',
])('rejects unsupported %s without dispatch', async (change) => {
  const s = mount();
  if (change === 'readonly') setReadonlyState(true);
  if (change === 'view') s.api.isViewMode = true;
  if (change === 'crop') s.image.blipFill.srcRect = { l: 10, t: 0, r: 90, b: 100 };
  if (change === 'invalid-rotation') s.image.spPr.xfrm.rot = NaN;
  if (change === 'invalid-flip') Reflect.set(s.image.spPr.xfrm, 'flipH', 1);
  if (change === 'rotated-offpage') {
    s.image.spPr.xfrm.rot = Math.PI / 4;
    s.image.spPr.xfrm.offX = 0;
  }
  if (change === 'group') Reflect.set(s.image, 'group', {});
  if (change === 'placeholder') s.image.isPlaceholder = () => true;
  if (change === 'wrong-type') s.image.getObjectType = () => 0;
  if (change === 'multiple') s.controller.selectedObjects.push(s.image);
  if (change === 'foreign') s.slide.cSld.spTree = [];
  if (change === 'mask') s.image.spPr.geometry.preset = 'ellipse';
  if (change === 'tile') Reflect.set(s.image.blipFill, 'tile', {});
  if (change === 'transparency') Reflect.set(s.image.blipFill, 'transparent', 127);
  if (change === 'offpage') s.image.spPr.xfrm.offX = 319;
  if (change === 'nonfinite') s.image.spPr.xfrm.extX = NaN;
  if (change === 'busy') s.api.isLongAction = () => true;
  if (change === 'slideshow') s.api.isSlideShow = () => true;
  if (change === 'master') s.logic.IsMasterMode = () => true;
  if (change === 'locked') s.common.CollaborativeEditing.Get_GlobalLock = () => true;
  if (change === 'unloaded') s.api.isDocumentLoadComplete = false;
  if (change === 'untracked') Reflect.deleteProperty(s.common, 'History');
  expect(canReplaceSlideImage()).toBe(false);
  await expect(prepareSlideImageReplacement(png)).rejects.toThrow();
  expect(s.api.ImgApply).not.toHaveBeenCalled();
  expect(s.common.sendImgUrls).not.toHaveBeenCalled();
});
it.each(['signature', 'truncated', 'too-large', 'zero-width', 'huge-height'])(
  'rejects invalid PNG %s',
  async (change) => {
    const s = mount();
    let value = png.slice();
    if (change === 'signature') value[0] = 0;
    if (change === 'truncated') value = value.slice(0, 24);
    if (change === 'too-large') value = new Uint8Array(4 * 1024 * 1024 + 1);
    if (change === 'zero-width') value.fill(0, 16, 20);
    if (change === 'huge-height') new DataView(value.buffer).setUint32(20, 4097);
    await expect(prepareSlideImageReplacement(value)).rejects.toThrow();
    expect(s.common.sendImgUrls).not.toHaveBeenCalled();
  },
);
it.each(['history', 'selection', 'geometry', 'rotation', 'flip', 'document', 'pointer', 'readonly', 'cancel'])(
  'rejects late %s while preloading',
  async (change) => {
    const s = mount();
    const abort = new AbortController();
    let loaded!: () => void;
    s.api.ImageLoader.LoadImagesWithCallback.mockImplementation((_urls, callback) => {
      loaded = callback;
    });
    const pending = prepareSlideImageReplacement(png, abort.signal);
    if (change === 'history') {
      s.history.Index++;
      s.history.Points.push({ Items: [{}] });
    }
    if (change === 'selection') s.controller.selectedObjects = [];
    if (change === 'geometry') s.image.spPr.xfrm.offX++;
    if (change === 'rotation') s.image.spPr.xfrm.rot = Math.PI / 4;
    if (change === 'flip') s.image.spPr.xfrm.flipH = true;
    if (change === 'document') s.api.WordControl.m_oLogicDocument = { ...s.logic };
    if (change === 'pointer') s.frame.contentDocument!.dispatchEvent(new Event('pointerdown'));
    if (change === 'readonly') setReadonlyState(true);
    if (change === 'cancel') abort.abort();
    loaded();
    await expect(pending).rejects.toThrow();
    expect(s.api.ImgApply).not.toHaveBeenCalled();
  },
);
it('rechecks the generation immediately before apply and never retries', async () => {
  const s = mount();
  const proposal = await prepareSlideImageReplacement(png);
  s.history.Points.push({ Items: [] });
  expect(() => proposal.apply()).toThrow(/changed/);
  expect(s.api.ImgApply).not.toHaveBeenCalled();
});
it.each(['no-op', 'resize', 'rotate', 'flip', 'crop', 'throw', 'cancel'])(
  'reports unverified %s after dispatch',
  async (change) => {
    const s = mount();
    const abort = new AbortController();
    const original = s.api.ImgApply.getMockImplementation()!;
    s.api.ImgApply.mockImplementation((value) => {
      if (change !== 'no-op') original(value);
      if (change === 'resize') s.image.spPr.xfrm.extX++;
      if (change === 'rotate') s.image.spPr.xfrm.rot = Math.PI / 4;
      if (change === 'flip') s.image.spPr.xfrm.flipV = true;
      if (change === 'crop') s.image.blipFill.srcRect = {};
      if (change === 'throw') throw new Error('native failure');
      if (change === 'cancel') abort.abort();
    });
    const proposal = await prepareSlideImageReplacement(png, abort.signal);
    expect(() => proposal.apply()).toThrow();
    expect(() => proposal.apply()).toThrow(/consumed/);
    expect(s.api.ImgApply).toHaveBeenCalledTimes(1);
  },
);
it('rejects decoder failures and times out without a document command, including a late callback', async () => {
  const s = mount();
  s.api.ImageLoader.LoadImage.mockReturnValue({ Image: { naturalWidth: 0, naturalHeight: 0 } });
  await expect(prepareSlideImageReplacement(png)).rejects.toThrow(/decode/);
  vi.useFakeTimers();
  let loaded!: () => void;
  s.api.ImageLoader.LoadImagesWithCallback.mockImplementation((_urls, callback) => {
    loaded = callback;
  });
  const pending = expect(prepareSlideImageReplacement(png)).rejects.toThrow(/timed out/);
  await vi.advanceTimersByTimeAsync(10000);
  await pending;
  loaded();
  expect(s.api.ImgApply).not.toHaveBeenCalled();
});

it.each([
  'header-crc',
  'data-crc',
  'end-crc',
  'truncated',
  'oversized-chunk',
  'trailing-bytes',
  'missing-end',
  'missing-data',
  'duplicate-header',
])('rejects PNG %s before native registration or mutation', async (change) => {
  const state = mount();
  let bad = png.slice();
  if (change === 'header-crc') bad[29] ^= 1;
  if (change === 'data-crc') bad[54] ^= 1;
  if (change === 'end-crc') bad[66] ^= 1;
  if (change === 'truncated') bad = bad.slice(0, -1);
  if (change === 'oversized-chunk') new DataView(bad.buffer).setUint32(33, 0x7fffffff);
  if (change === 'trailing-bytes') bad = Uint8Array.from([...bad, 0]);
  if (change === 'duplicate-header') bad = Uint8Array.from([...bad.slice(0, 33), ...bad.slice(8)]);
  if (change === 'missing-end') bad = bad.slice(0, -12);
  if (change === 'missing-data') bad = Uint8Array.from([...bad.slice(0, 33), ...bad.slice(-12)]);
  await expect(prepareSlideImageReplacement(bad)).rejects.toThrow(/PNG/);
  expect(state.common.sendImgUrls).not.toHaveBeenCalled();
  expect(state.api.ImgApply).not.toHaveBeenCalled();
});

it('rejects the formerly accepted IDAT CRC error even when the native cache reports a decoded image', async () => {
  const state = mount();
  const corrupt = Uint8Array.from(
    Buffer.from(
      'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+j4z0AAAAASUVORK5CYII=',
      'base64',
    ),
  );
  await expect(prepareSlideImageReplacement(corrupt)).rejects.toThrow(/PNG/);
  expect(state.common.sendImgUrls).not.toHaveBeenCalled();
  expect(state.api.ImgApply).not.toHaveBeenCalled();
});

it('validates complete PNG bytes through a nonzero-offset view', async () => {
  const state = mount();
  const storage = new Uint8Array(png.length + 8);
  storage.set(png, 4);
  const prepared = await prepareSlideImageReplacement(storage.subarray(4, -4));
  expect(prepared.apply().verified).toBe(true);
  expect(state.api.ImgApply).toHaveBeenCalledTimes(1);
});

const sizeOwner = { projectId: 'project', documentId: 'doc', sessionId: 'size-session' };
function resizeSetup() {
  const state = mount();
  state.api.ImgApply.mockImplementation((props) => {
    Object.assign(state.image.spPr.xfrm, { extX: Reflect.get(props, 'Width'), extY: Reflect.get(props, 'Height') });
    state.history.Index++;
    state.history.Points.push({ Items: [{}] });
  });
  const bridge = new DeckWiseEditorBridge();
  bridge.reset(sizeOwner);
  const command = (input: unknown = { widthMm: 80.5, heightMm: 30 }) => ({
    ...sizeOwner,
    requestId: 'resize-image',
    editGeneration: bridge.getContext(sizeOwner).editGeneration,
    tool: 'set_slide_image_size',
    input,
  });
  return { ...state, bridge, command };
}

it('advertises and resizes one ordinary image with native readback, preserving asset, position, style and peers', async () => {
  const state = resizeSetup();
  const peer = { ...state.image, Id: 'peer-image', spPr: structuredClone(state.image.spPr) };
  state.slide.cSld.spTree.push(peer);
  Reflect.set(state.image.spPr, 'ln', { width: 1, color: 'blue' });
  const original = JSON.parse(JSON.stringify(state.image));
  const peerBefore = JSON.parse(JSON.stringify(peer));
  const context = state.bridge.getContext(sizeOwner);
  expect(context.selectionCharacters).toBe(0);
  expect(context.capabilities.map((capability) => capability.name)).toContain('set_slide_image_size');
  const command = state.command();
  await expect(state.bridge.execute(command, command.requestId)).resolves.toMatchObject({
    ...sizeOwner,
    tool: 'set_slide_image_size',
    status: 'verified',
    result: { widthMm: 80.5, heightMm: 30, page: 1, verified: true },
  });
  expect(state.api.ImgApply).toHaveBeenCalledExactlyOnceWith({ Width: 80.5, Height: 30 });
  expect(JSON.parse(JSON.stringify(state.image))).toEqual({
    ...original,
    spPr: { ...original.spPr, xfrm: { ...original.spPr.xfrm, extX: 80.5, extY: 30 } },
  });
  expect(JSON.parse(JSON.stringify(peer))).toEqual(peerBefore);
  expect(state.common.sendImgUrls).not.toHaveBeenCalled();
  await expect(state.bridge.execute({ ...command, requestId: 'stale-size' }, 'stale-size')).rejects.toThrow(/changed/);
  expect(state.api.ImgApply).toHaveBeenCalledTimes(1);
});

it.each([
  'readonly',
  'rotation',
  'flip',
  'crop',
  'group',
  'placeholder',
  'mask',
  'multiple',
  'offpage',
  'busy',
  'master',
  'locked',
  'untracked',
])('withholds image resize and rejects %s before dispatch', async (change) => {
  const state = resizeSetup();
  if (change === 'readonly') setReadonlyState(true);
  if (change === 'rotation') state.image.spPr.xfrm.rot = Math.PI / 4;
  if (change === 'flip') state.image.spPr.xfrm.flipH = true;
  if (change === 'crop') state.image.blipFill.srcRect = {};
  if (change === 'group') Reflect.set(state.image, 'group', {});
  if (change === 'placeholder') state.image.isPlaceholder = () => true;
  if (change === 'mask') state.image.spPr.geometry.preset = 'ellipse';
  if (change === 'multiple') state.controller.selectedObjects.push(state.image);
  if (change === 'offpage') state.image.spPr.xfrm.offX = 319;
  if (change === 'busy') state.api.isLongAction = () => true;
  if (change === 'master') state.logic.IsMasterMode = () => true;
  if (change === 'locked') state.common.CollaborativeEditing.Get_GlobalLock = () => true;
  if (change === 'untracked') Reflect.deleteProperty(state.common, 'History');
  expect(state.bridge.getContext(sizeOwner).capabilities.map((capability) => capability.name)).not.toContain(
    'set_slide_image_size',
  );
  const command = state.command();
  await expect(state.bridge.execute(command, command.requestId)).rejects.toThrow();
  expect(state.api.ImgApply).not.toHaveBeenCalled();
});

it.each([
  { widthMm: 0, heightMm: 30 },
  { widthMm: 80, heightMm: NaN },
  { widthMm: Infinity, heightMm: 30 },
  { widthMm: 1001, heightMm: 30 },
  { widthMm: 311, heightMm: 30 },
  { widthMm: 80, heightMm: 191 },
  { widthMm: 80, heightMm: 30, script: 'bad' },
  { widthMm: 80 },
])('rejects unsupported image dimensions/input %s before native dispatch', async (input) => {
  const state = resizeSetup(),
    command = state.command(input);
  await expect(state.bridge.execute(command, command.requestId)).rejects.toThrow();
  expect(state.api.ImgApply).not.toHaveBeenCalled();
});

it.each(['no-op', 'move', 'asset', 'style', 'peer', 'rotation', 'selection', 'page', 'throw', 'cancel'])(
  'reports uncertain image resize %s without replay',
  async (change) => {
    const state = resizeSetup();
    const peer = { ...state.image, Id: 'peer-image', spPr: structuredClone(state.image.spPr) };
    state.slide.cSld.spTree.push(peer);
    const apply = state.api.ImgApply.getMockImplementation()!;
    state.api.ImgApply.mockImplementation((props) => {
      if (change !== 'no-op') apply(props);
      if (change === 'move') state.image.spPr.xfrm.offX++;
      if (change === 'asset') state.image.blipFill.RasterImageId = 'changed.png';
      if (change === 'style') Reflect.set(state.image.spPr, 'ln', { width: 99 });
      if (change === 'peer') peer.spPr.xfrm.extX++;
      if (change === 'rotation') state.image.spPr.xfrm.rot = 1;
      if (change === 'selection') state.controller.selectedObjects = [peer];
      if (change === 'page') state.api.getCurrentPage = () => 1;
      if (change === 'throw') throw new Error('Native failure');
      if (change === 'cancel') state.bridge.setImageAssetSource(undefined);
    });
    const command = state.command();
    await expect(state.bridge.execute(command, command.requestId)).rejects.toThrow(/unverified|Inspect|inspect/);
    await expect(state.bridge.execute(command, command.requestId)).rejects.toThrow(/consumed/);
    expect(state.api.ImgApply).toHaveBeenCalledTimes(1);
  },
);

it('rejects an intervening user interaction before image resize and requires fresh context', async () => {
  const state = resizeSetup(),
    command = state.command();
  state.frame.contentDocument!.dispatchEvent(new Event('pointerdown'));
  await expect(state.bridge.execute(command, command.requestId)).rejects.toThrow(/changed/);
  expect(state.api.ImgApply).not.toHaveBeenCalled();
});

it('withholds resize for the rotated flipped B49 target but supports its ordinary same-page peer', async () => {
  const state = resizeSetup();
  // B49 source.pptx slide 2: TARGET is 45 degrees and flipped on both axes.
  Object.assign(state.image.spPr.xfrm, {
    offX: 25.4,
    offY: 50.8,
    extX: 76.2,
    extY: 50.8,
    rot: Math.PI / 4,
    flipH: true,
    flipV: true,
  });
  const peer = {
    ...state.image,
    Id: 'ordinary-same-page-peer',
    blipFill: { ...state.image.blipFill, RasterImageId: 'peer.png' },
    spPr: {
      ...state.image.spPr,
      xfrm: { offX: 152.4, offY: 50.8, extX: 76.2, extY: 50.8, rot: null, flipH: false, flipV: false },
    },
  };
  state.slide.cSld.spTree.push(peer);
  expect(state.bridge.getContext(sizeOwner).capabilities.map((item) => item.name)).not.toContain(
    'set_slide_image_size',
  );
  const unsupported = state.command();
  await expect(state.bridge.execute(unsupported, unsupported.requestId)).rejects.toThrow();
  expect(state.api.ImgApply).not.toHaveBeenCalled();

  const targetBefore = JSON.parse(JSON.stringify(state.image));
  state.controller.selectedObjects = [peer];
  state.api.ImgApply.mockImplementation((props) => {
    Object.assign(peer.spPr.xfrm, { extX: Reflect.get(props, 'Width'), extY: Reflect.get(props, 'Height') });
    state.history.Index++;
    state.history.Points.push({ Items: [{}] });
  });
  expect(state.bridge.getContext(sizeOwner).capabilities.map((item) => item.name)).toContain('set_slide_image_size');
  const command = { ...state.command(), requestId: 'resize-supported-peer' };
  await expect(state.bridge.execute(command, command.requestId)).resolves.toMatchObject({
    tool: 'set_slide_image_size',
    result: { widthMm: 80.5, heightMm: 30, verified: true },
  });
  expect(peer.spPr.xfrm).toMatchObject({ offX: 152.4, offY: 50.8, extX: 80.5, extY: 30 });
  expect(peer.blipFill.RasterImageId).toBe('peer.png');
  expect(JSON.parse(JSON.stringify(state.image))).toEqual(targetBefore);
});

function positionSetup() {
  const state = mount();
  state.api.ImgApply.mockImplementation((props) => {
    const position = Reflect.get(props, 'Position');
    Object.assign(state.image.spPr.xfrm, { offX: position.X, offY: position.Y });
    state.history.Index++;
    state.history.Points.push({ Items: [{}] });
  });
  const bridge = new DeckWiseEditorBridge();
  bridge.reset(sizeOwner);
  const command = (input: unknown = { xMm: 25.5, yMm: 35 }) => ({
    ...sizeOwner,
    requestId: 'position-image',
    editGeneration: bridge.getContext(sizeOwner).editGeneration,
    tool: 'set_slide_image_position',
    input,
  });
  return { ...state, bridge, command };
}

it('moves one selected ordinary image using native millimeter coordinates, preserving dimensions, asset and peers', async () => {
  const state = positionSetup();
  const peer = { ...state.image, Id: 'peer-image', spPr: structuredClone(state.image.spPr) };
  state.slide.cSld.spTree.push(peer);
  Reflect.set(state.image.spPr, 'ln', { width: 1, color: 'blue' });
  const original = JSON.parse(JSON.stringify(state.image));
  const peerBefore = JSON.parse(JSON.stringify(peer));
  expect(state.bridge.getContext(sizeOwner).capabilities.map((item) => item.name)).toContain(
    'set_slide_image_position',
  );
  const command = state.command();
  await expect(state.bridge.execute(command, command.requestId)).resolves.toMatchObject({
    tool: 'set_slide_image_position',
    status: 'verified',
    result: { xMm: 25.5, yMm: 35, page: 1, verified: true },
  });
  expect(state.api.ImgApply).toHaveBeenCalledExactlyOnceWith({ Position: { X: 25.5, Y: 35 } });
  expect(JSON.parse(JSON.stringify(state.image))).toEqual({
    ...original,
    spPr: { ...original.spPr, xfrm: { ...original.spPr.xfrm, offX: 25.5, offY: 35 } },
  });
  expect(JSON.parse(JSON.stringify(peer))).toEqual(peerBefore);
  expect(state.common.sendImgUrls).not.toHaveBeenCalled();
  await expect(state.bridge.execute({ ...command, requestId: 'stale-position' }, 'stale-position')).rejects.toThrow(
    /changed/,
  );
  expect(state.api.ImgApply).toHaveBeenCalledTimes(1);
});

it.each([
  'readonly',
  'view',
  'rotation',
  'flip',
  'crop',
  'group',
  'placeholder',
  'mask',
  'multiple',
  'offpage',
  'busy',
  'slideshow',
  'master',
  'locked',
  'untracked',
  'tile',
  'transparency',
])('withholds ordinary image positioning for unsupported %s', async (change) => {
  const state = positionSetup();
  if (change === 'readonly') setReadonlyState(true);
  if (change === 'view') state.api.isViewMode = true;
  if (change === 'rotation') state.image.spPr.xfrm.rot = Math.PI / 4;
  if (change === 'flip') state.image.spPr.xfrm.flipV = true;
  if (change === 'crop') state.image.blipFill.srcRect = {};
  if (change === 'group') Reflect.set(state.image, 'group', {});
  if (change === 'placeholder') state.image.isPlaceholder = () => true;
  if (change === 'mask') state.image.spPr.geometry.preset = 'ellipse';
  if (change === 'multiple') state.controller.selectedObjects.push(state.image);
  if (change === 'offpage') state.image.spPr.xfrm.offX = 319;
  if (change === 'busy') state.api.isLongAction = () => true;
  if (change === 'slideshow') state.api.isSlideShow = () => true;
  if (change === 'master') state.logic.IsMasterMode = () => true;
  if (change === 'locked') state.common.CollaborativeEditing.Get_GlobalLock = () => true;
  if (change === 'untracked') Reflect.deleteProperty(state.common, 'History');
  if (change === 'tile') Reflect.set(state.image.blipFill, 'tile', {});
  if (change === 'transparency') Reflect.set(state.image.blipFill, 'transparent', 127);
  expect(state.bridge.getContext(sizeOwner).capabilities.map((item) => item.name)).not.toContain(
    'set_slide_image_position',
  );
  const command = state.command();
  await expect(state.bridge.execute(command, command.requestId)).rejects.toThrow();
  expect(state.api.ImgApply).not.toHaveBeenCalled();
});

it.each([
  { xMm: -1, yMm: 35 },
  { xMm: 25, yMm: NaN },
  { xMm: Infinity, yMm: 35 },
  { xMm: 1001, yMm: 35 },
  { xMm: 301, yMm: 35 },
  { xMm: 25, yMm: 181 },
  { xMm: 25, yMm: 35, script: 'bad' },
  { xMm: 25 },
])('rejects invalid or out-of-slide image position %s before dispatch', async (input) => {
  const state = positionSetup(),
    command = state.command(input);
  await expect(state.bridge.execute(command, command.requestId)).rejects.toThrow();
  expect(state.api.ImgApply).not.toHaveBeenCalled();
});

it.each([
  { xMm: 0, yMm: 0 },
  { xMm: 300, yMm: 180 },
])('accepts a complete image touching the slide boundary %s', async (input) => {
  const state = positionSetup(),
    command = state.command(input);
  await expect(state.bridge.execute(command, command.requestId)).resolves.toMatchObject({
    result: { ...input, page: 1, verified: true },
  });
});

it.each(['no-op', 'resize', 'asset', 'style', 'peer', 'rotation', 'selection', 'page', 'throw', 'cancel'])(
  'reports uncertain native image positioning %s without replay',
  async (change) => {
    const state = positionSetup();
    const peer = { ...state.image, Id: 'peer-image', spPr: structuredClone(state.image.spPr) };
    state.slide.cSld.spTree.push(peer);
    const apply = state.api.ImgApply.getMockImplementation()!;
    state.api.ImgApply.mockImplementation((props) => {
      if (change !== 'no-op') apply(props);
      if (change === 'resize') state.image.spPr.xfrm.extX++;
      if (change === 'asset') state.image.blipFill.RasterImageId = 'changed.png';
      if (change === 'style') Reflect.set(state.image.spPr, 'ln', { width: 99 });
      if (change === 'peer') peer.spPr.xfrm.offX++;
      if (change === 'rotation') state.image.spPr.xfrm.rot = 1;
      if (change === 'selection') state.controller.selectedObjects = [peer];
      if (change === 'page') state.api.getCurrentPage = () => 1;
      if (change === 'throw') throw new Error('Native failure');
      if (change === 'cancel') state.bridge.setImageAssetSource(undefined);
    });
    const command = state.command();
    await expect(state.bridge.execute(command, command.requestId)).rejects.toThrow(/unverified|Inspect|inspect/);
    await expect(state.bridge.execute(command, command.requestId)).rejects.toThrow(/consumed/);
    expect(state.api.ImgApply).toHaveBeenCalledTimes(1);
  },
);

it('rejects changed selection interaction before native image positioning', async () => {
  const state = positionSetup(),
    command = state.command();
  state.frame.contentDocument!.dispatchEvent(new Event('pointerdown'));
  await expect(state.bridge.execute(command, command.requestId)).rejects.toThrow(/changed/);
  expect(state.api.ImgApply).not.toHaveBeenCalled();
});
