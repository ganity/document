import { createHash, webcrypto } from 'node:crypto';
import { afterEach, expect, it, vi } from 'vitest';
import { nativeLayerEvidence } from '../../lib/agent-plugin/slide-layer-evidence';

const state = vi.hoisted(() => ({ url: '', registered: true }));
vi.mock('../../lib/agent-plugin/editor-bridge', () => ({
  requireEditorContext: () => ({
    AscCommon: {
      getFullImageSrc2: () => state.url,
      g_oDocumentUrls: { getImageLocal: () => (state.registered ? 'local.png' : null) },
    },
  }),
}));
const png = Buffer.from(
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAACXBIWXMAAAPoAAAD6AG1e1JrAAAADUlEQVR4nGP4z8DwHwAFAAH/iZk9HQAAAABJRU5ErkJggg==',
  'base64',
);
function fixture() {
  vi.stubGlobal('crypto', webcrypto);
  state.registered = true;
  state.url = 'data:image/png;base64,' + png.toString('base64');
  const properties = {
    descr: JSON.stringify({
      schema: 'deckwise-layer/1',
      name: 'Evaporation',
      description: 'Upward vapor arrows and label',
      sha256: createHash('sha256').update(png).digest('hex'),
    }),
  };
  const shape = {
    isImage: () => true,
    getUniNvProps: () => ({ cNvPr: properties }),
    blipFill: { RasterImageId: 'local.png' },
  };
  return { shape, properties };
}
afterEach(() => vi.unstubAllGlobals());

it('reads persisted layer hints only with matching current local image bytes', async () => {
  const { shape } = fixture();
  expect(await nativeLayerEvidence(shape)).toEqual({
    layerMetadata: { name: 'Evaporation', description: 'Upward vapor arrows and label' },
    dataUrl: state.url,
  });
  // Fresh live objects after reopen retain the same identity evidence.
  expect(await nativeLayerEvidence({ ...shape, blipFill: { ...shape.blipFill } })).toBeDefined();
  const replacement = Buffer.from(png);
  replacement[replacement.length - 1] ^= 1;
  state.url = 'data:image/png;base64,' + replacement.toString('base64');
  expect(await nativeLayerEvidence(shape)).toBeUndefined();
});

it('does not fetch remote, unregistered, cropped, missing or malformed metadata', async () => {
  const { shape, properties } = fixture();
  state.url = 'https://example.invalid/image.png';
  const fetch = vi.fn();
  vi.stubGlobal('fetch', fetch);
  expect(await nativeLayerEvidence(shape)).toBeUndefined();
  expect(fetch).not.toHaveBeenCalled();
  state.url = 'data:image/png;base64,' + png.toString('base64');
  state.registered = false;
  expect(await nativeLayerEvidence(shape)).toBeUndefined();
  state.registered = true;
  expect(await nativeLayerEvidence({ ...shape, blipFill: { ...shape.blipFill, srcRect: {} } })).toBeUndefined();
  expect(await nativeLayerEvidence({ ...shape, spPr: { xfrm: { flipV: true } } })).toBeUndefined();
  expect(await nativeLayerEvidence({ ...shape, spPr: { xfrm: { rot: 1 } } })).toBeUndefined();
  properties.descr = '{}';
  expect(await nativeLayerEvidence(shape)).toBeUndefined();
  properties.descr = 'legacy image';
  expect(await nativeLayerEvidence(shape)).toBeUndefined();
});
