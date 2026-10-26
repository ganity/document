import { afterEach, beforeEach, expect, it, vi } from 'vitest';
import { canInsertSlideImage, prepareSlideImageInsertion } from '../../lib/agent-plugin/slide-image-insertion';
import { setReadonlyState } from '../../lib/onlyoffice/readonly';
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
  const slide = { cSld: { spTree: [image] }, deleteLock: { Lock: { Type: 1 } } };
  const makeImage = vi.fn((url: string, x: number, y: number, w: number, h: number) => ({
    ...image,
    Id: 'inserted',
    blipFill: { RasterImageId: 'new.png', srcRect: null },
    spPr: {
      geometry: { preset: 'rect' },
      xfrm: { offX: x, offY: y, extX: w, extY: h, rot: null, flipH: false, flipV: false },
    },
    setParent: vi.fn(),
    addToDrawingObjects() {
      slide.cSld.spTree.push(this);
    },
  }));
  const controller = {
    selectedObjects: [] as (typeof image)[],
    getTargetDocContent: (): unknown => undefined,
    resetSelection: vi.fn(() => {
      controller.selectedObjects = [];
    }),
    createImage: makeImage,
    selectObject: vi.fn((added: typeof image) => {
      controller.selectedObjects = [added];
    }),
  };
  const logic = {
    Slides: [slide],
    FocusOnNotes: false,
    StartAction: vi.fn((_description: number) => {
      history.Index++;
      history.Points.push({ Items: [{}] });
    }),
    FinalizeAction: vi.fn(),
    Document_UpdateInterfaceState: vi.fn(),
    CheckEmptyPlaceholderNotes: vi.fn(),
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
    c_oAscLockTypes: { kLockTypeNone: 1, kLockTypeMine: 2 },
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
    AscDFH: { historyitem_type_ImageShape: 72548352, historydescription_Presentation_AddFlowImage: 1234 },
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

const box = { xMm: 25, yMm: 35, widthMm: 40, heightMm: 30 };
it('preloads without mutation, inserts once in one native action, and preserves peers', async () => {
  const s = mount(),
    original = structuredClone(s.image.spPr);
  expect(canInsertSlideImage()).toBe(true);
  const prepared = await prepareSlideImageInsertion(png, box);
  expect(s.controller.createImage).not.toHaveBeenCalled();
  expect(prepared.apply()).toEqual({ page: 1, imageId: 'inserted', box, verified: true, saved: false });
  expect(s.controller.createImage).toHaveBeenCalledExactlyOnceWith('blob:local', 25, 35, 40, 30);
  expect(s.logic.StartAction).toHaveBeenCalledExactlyOnceWith(1234);
  expect(s.logic.FinalizeAction).toHaveBeenCalledTimes(1);
  expect(s.slide.cSld.spTree).toHaveLength(2);
  expect(s.image.spPr).toEqual(original);
  expect(() => prepared.apply()).toThrow(/consumed/);
});
it('pins explicit geometry before async preload', async () => {
  const s = mount();
  let done!: () => void;
  s.api.ImageLoader.LoadImagesWithCallback.mockImplementation((_urls, callback) => {
    done = callback;
  });
  const mutable = { ...box };
  const pending = prepareSlideImageInsertion(png, mutable);
  mutable.xMm = 100;
  mutable.widthMm = 200;
  done();
  const prepared = await pending;
  expect(prepared.box).toEqual(box);
  expect(prepared.apply().box).toEqual(box);
});
it.each([
  'readonly',
  'view',
  'cannot-edit',
  'loading',
  'busy',
  'slideshow',
  'master',
  'notes',
  'locked',
  'slide-lock',
  'selection',
  'text-selection',
  'untracked',
  'missing-native',
  'missing-constant',
])('refuses %s before PNG registration or insertion', async (change) => {
  const s = mount();
  if (change === 'readonly') setReadonlyState(true);
  if (change === 'view') s.api.isViewMode = true;
  if (change === 'cannot-edit') s.logic.CanEdit = () => false;
  if (change === 'loading') s.api.isDocumentLoadComplete = false;
  if (change === 'busy') s.api.isLongAction = () => true;
  if (change === 'slideshow') s.api.isSlideShow = () => true;
  if (change === 'master') s.logic.IsMasterMode = () => true;
  if (change === 'notes') s.logic.FocusOnNotes = true;
  if (change === 'locked') s.common.CollaborativeEditing.Get_GlobalLock = () => true;
  if (change === 'slide-lock') s.slide.deleteLock.Lock.Type = 3;
  if (change === 'selection') s.controller.selectedObjects = [s.image];
  if (change === 'text-selection') s.controller.getTargetDocContent = () => ({});
  if (change === 'untracked') Reflect.deleteProperty(s.common, 'History');
  if (change === 'missing-native') Reflect.deleteProperty(s.controller, 'createImage');
  if (change === 'missing-constant')
    Reflect.deleteProperty(
      Reflect.get(s.frame.contentWindow!, 'AscDFH'),
      'historydescription_Presentation_AddFlowImage',
    );
  expect(canInsertSlideImage()).toBe(false);
  await expect(prepareSlideImageInsertion(png, box)).rejects.toThrow();
  expect(s.common.sendImgUrls).not.toHaveBeenCalled();
  expect(s.logic.StartAction).not.toHaveBeenCalled();
});
it.each([
  { ...box, xMm: -1 },
  { ...box, widthMm: 0 },
  { ...box, yMm: NaN },
  { ...box, heightMm: Infinity },
  { ...box, xMm: 300 },
  { ...box, yMm: 190 },
])('refuses invalid/off-page geometry %j', async (value) => {
  const s = mount();
  await expect(prepareSlideImageInsertion(png, value)).rejects.toThrow();
  expect(s.common.sendImgUrls).not.toHaveBeenCalled();
  expect(s.logic.StartAction).not.toHaveBeenCalled();
});
it.each([
  'history',
  'selection',
  'page-size',
  'peer-geometry',
  'peer-url',
  'peer-order',
  'pointer',
  'readonly',
  'cancel',
])('refuses changed %s after async preload', async (change) => {
  const s = mount(),
    abort = new AbortController();
  let done!: () => void;
  s.api.ImageLoader.LoadImagesWithCallback.mockImplementation((_urls, callback) => {
    done = callback;
  });
  const pending = prepareSlideImageInsertion(png, box, abort.signal);
  if (change === 'history') s.history.Points.push({ Items: [{}] });
  if (change === 'selection') s.controller.selectedObjects = [s.image];
  if (change === 'page-size') s.logic.GetWidthMM = () => 300;
  if (change === 'peer-geometry') s.image.spPr.xfrm.offX++;
  if (change === 'peer-url') s.image.blipFill.RasterImageId = 'foreign.png';
  if (change === 'peer-order') s.slide.cSld.spTree = [];
  if (change === 'pointer') s.frame.contentDocument!.dispatchEvent(new Event('pointerdown'));
  if (change === 'readonly') setReadonlyState(true);
  if (change === 'cancel') abort.abort();
  done();
  await expect(pending).rejects.toThrow();
  expect(s.logic.StartAction).not.toHaveBeenCalled();
});
it('rechecks context at apply and consumes a stale prepared insertion', async () => {
  const s = mount(),
    prepared = await prepareSlideImageInsertion(png, box);
  s.history.Points.push({ Items: [] });
  expect(() => prepared.apply()).toThrow(/changed/);
  expect(() => prepared.apply()).toThrow(/consumed/);
  expect(s.logic.StartAction).not.toHaveBeenCalled();
});
it.each(['no-op', 'wrong-size', 'wrong-url', 'peer-change', 'extra-object', 'no-selection', 'cancel', 'throw'])(
  'reports uncertain native %s without retry or rollback',
  async (change) => {
    const s = mount(),
      abort = new AbortController(),
      original = s.controller.createImage.getMockImplementation()!;
    s.controller.createImage.mockImplementation((...args) => {
      const added = original(...args);
      if (change === 'no-op') added.addToDrawingObjects = () => {};
      if (change === 'wrong-size') added.spPr.xfrm.extX++;
      if (change === 'wrong-url') added.blipFill.RasterImageId = 'other.png';
      if (change === 'peer-change') s.image.spPr.xfrm.offX++;
      if (change === 'extra-object') s.slide.cSld.spTree.push(s.image);
      if (change === 'no-selection') s.controller.selectObject.mockImplementation(() => {});
      if (change === 'cancel') abort.abort();
      if (change === 'throw') throw new Error('native');
      return added;
    });
    const prepared = await prepareSlideImageInsertion(png, box, abort.signal);
    expect(() => prepared.apply()).toThrow(/unverified.*Undo/);
    expect(() => prepared.apply()).toThrow(/consumed/);
    expect(s.controller.createImage).toHaveBeenCalledTimes(1);
    expect(s.logic.FinalizeAction).toHaveBeenCalledTimes(1);
  },
);
it('rejects malformed PNG and decode mismatch without document mutation', async () => {
  const s = mount(),
    bad = png.slice();
  bad[54] ^= 1;
  await expect(prepareSlideImageInsertion(bad, box)).rejects.toThrow(/CRC/);
  expect(s.common.sendImgUrls).not.toHaveBeenCalled();
  s.api.ImageLoader.LoadImage.mockReturnValue({ Image: { naturalWidth: 2, naturalHeight: 1 } });
  await expect(prepareSlideImageInsertion(png, box)).rejects.toThrow(/decode/);
  expect(s.logic.StartAction).not.toHaveBeenCalled();
});
it('times out and ignores late preload for mutation', async () => {
  const s = mount();
  vi.useFakeTimers();
  let done!: () => void;
  s.api.ImageLoader.LoadImagesWithCallback.mockImplementation((_urls, callback) => {
    done = callback;
  });
  const pending = expect(prepareSlideImageInsertion(png, box)).rejects.toThrow(/timed out/);
  await vi.advanceTimersByTimeAsync(10000);
  await pending;
  done();
  expect(s.logic.StartAction).not.toHaveBeenCalled();
});
