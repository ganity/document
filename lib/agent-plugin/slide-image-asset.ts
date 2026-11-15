import { captureDocumentToolTarget } from './document-tool-action';
import { prepareSlideImageReplacement } from './slide-image-replacement';
import { prepareSlideImageInsertion, type SlideImageBox } from './slide-image-insertion';

interface ImageBinding {
  projectId: string;
  documentId: string;
  sessionId: string;
  editGeneration: number;
}

/** Host-private metadata pins; this is not the canonical AssetRef or an authorization grant. */
export interface SlideImageAsset {
  assetId: string;
  sha256: string;
  mime: 'image/png';
  byteLength: number;
  width: number;
  height: number;
}

interface ImageAssetInput {
  binding: ImageBinding;
  asset: SlideImageAsset;
  loadAsset: (assetId: string, signal?: AbortSignal) => Promise<Uint8Array>;
  signal?: AbortSignal;
}
type ImageState = ImageBinding & { readOnly: boolean; saving: boolean };
type ReplacementInput = ImageAssetInput & { readCurrent: () => ImageState & { imageReplacementInstalled: boolean } };
type InsertionInput = ImageAssetInput & {
  box: SlideImageBox;
  readCurrent: () => ImageState & { imageInsertionInstalled: boolean };
};

/** Resolve bytes with the current user's authorization; never accept model URLs or credentials here. */
export function replaceSlideImageFromAsset(input: ReplacementInput) {
  return applySlideImageAsset(input);
}

export async function insertSlideImageFromAsset(input: InsertionInput) {
  const box = { ...input.box };
  if (
    !Object.values(box).every((value) => typeof value === 'number' && Number.isFinite(value)) ||
    Object.keys(box).length !== 4 ||
    !['xMm', 'yMm', 'widthMm', 'heightMm'].every((key) => key in box) ||
    box.xMm < 0 ||
    box.yMm < 0 ||
    box.widthMm < 1 ||
    box.heightMm < 1 ||
    Object.values(box).some((value) => value > 1000)
  )
    throw new TypeError('Explicit bounded image geometry is required');
  const result = await applySlideImageAsset(input, box);
  return { ...result, ...box };
}

async function applySlideImageAsset(input: ReplacementInput | InsertionInput, box?: SlideImageBox) {
  const binding = { ...input.binding };
  const asset = { ...input.asset };
  const readCurrent = input.readCurrent.bind(input);
  const loadAsset = input.loadAsset.bind(input);
  const { signal } = input;
  if (
    ![binding.projectId, binding.documentId, binding.sessionId, asset.assetId].every(
      (value) => typeof value === 'string' && value.trim().length > 0 && value.length <= 512,
    ) ||
    !Number.isSafeInteger(binding.editGeneration) ||
    binding.editGeneration < 1 ||
    asset.mime !== 'image/png' ||
    !/^[a-f0-9]{64}$/.test(asset.sha256) ||
    !Number.isSafeInteger(asset.byteLength) ||
    asset.byteLength < 33 ||
    asset.byteLength > 4 * 1024 * 1024 ||
    ![asset.width, asset.height].every((value) => Number.isSafeInteger(value) && value >= 1 && value <= 4096)
  )
    throw new TypeError('A bound editor and pinned PNG asset are required');
  const target = captureDocumentToolTarget();
  const fresh = () => {
    signal?.throwIfAborted();
    const now = readCurrent();
    if (
      now.projectId !== binding.projectId ||
      now.documentId !== binding.documentId ||
      now.sessionId !== binding.sessionId ||
      now.editGeneration !== binding.editGeneration ||
      now.readOnly !== false ||
      now.saving !== false ||
      Reflect.get(now, box ? 'imageInsertionInstalled' : 'imageReplacementInstalled') !== true ||
      target.context.kind !== 'slide' ||
      !target.isCurrent(true)
    )
      throw new Error('Image context changed or capability unavailable; read fresh context');
  };
  fresh();
  const loaded = await loadAsset(asset.assetId, signal);
  fresh();
  if (!(loaded instanceof Uint8Array) || loaded.byteLength !== asset.byteLength)
    throw new Error('Image asset length mismatch');
  // Own the bytes before hashing; an awaiting resolver must not mutate the verified payload.
  const png = new Uint8Array(loaded);
  const digest = new Uint8Array(await crypto.subtle.digest('SHA-256', png));
  fresh();
  const sha256 = Array.from(digest, (byte) => byte.toString(16).padStart(2, '0')).join('');
  const header = new DataView(png.buffer);
  if (sha256 !== asset.sha256 || header.getUint32(16) !== asset.width || header.getUint32(20) !== asset.height)
    throw new Error('Image asset hash or dimensions mismatch');
  const prepared = box
    ? await prepareSlideImageInsertion(png, box, signal)
    : await prepareSlideImageReplacement(png, signal);
  fresh();
  const result = prepared.apply();
  return { ...result, assetId: asset.assetId, sha256, saved: false as const };
}
