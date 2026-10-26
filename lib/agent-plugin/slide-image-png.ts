import { bytesToBase64 } from 'ranuts/utils';
import type { EditorApi } from './editor-bridge';

export interface PngApi extends EditorApi {
  ImageLoader?: {
    LoadImagesWithCallback(urls: string[], callback: () => void): void;
    LoadImage(url: string, type: number): { Image?: { naturalWidth: number; naturalHeight: number } } | null;
  };
}
export interface PngCommon {
  sendImgUrls(api: EditorApi, urls: string[], callback: (result: Array<{ path: string; url: string }>) => void): void;
  getFullImageSrc2(local: string): string;
  g_oDocumentUrls: { imagePath2Local(path: string): string; getImageLocal(url: string): string | null };
}

function hasValidPngChunks(png: Uint8Array): boolean {
  const view = new DataView(png.buffer, png.byteOffset, png.byteLength);
  let offset = 8;
  let hasData = false;
  while (offset + 12 <= png.length) {
    const length = view.getUint32(offset);
    const end = offset + 12 + length;
    if (end > png.length) return false;
    const type = String.fromCharCode(...png.subarray(offset + 4, offset + 8));
    if ((offset === 8 && (type !== 'IHDR' || length !== 13)) || (type === 'IHDR' && offset !== 8)) return false;
    if (type === 'IDAT') hasData = true;
    let crc = 0xffffffff;
    for (let index = offset + 4; index < end - 4; index++) {
      crc ^= png[index];
      for (let bit = 0; bit < 8; bit++) crc = (crc >>> 1) ^ (crc & 1 ? 0xedb88320 : 0);
    }
    if ((crc ^ 0xffffffff) >>> 0 !== view.getUint32(end - 4)) return false;
    offset = end;
    if (type === 'IEND') return length === 0 && hasData && offset === png.length;
  }
  return false;
}

export function validateSlidePng(png: Uint8Array, signal?: AbortSignal) {
  if (
    !(png instanceof Uint8Array) ||
    png.length < 33 ||
    png.length > 4 * 1024 * 1024 ||
    ![137, 80, 78, 71, 13, 10, 26, 10, 0, 0, 0, 13, 73, 72, 68, 82].every((byte, i) => png[i] === byte)
  )
    throw new TypeError('A local PNG asset between 33 bytes and 4 MiB is required');
  signal?.throwIfAborted();
  if (!hasValidPngChunks(png)) throw new TypeError('PNG chunks are incomplete or have invalid CRCs');
  const header = new DataView(png.buffer, png.byteOffset, png.byteLength);
  const width = header.getUint32(16),
    height = header.getUint32(20);
  if (!width || !height || width > 4096 || height > 4096)
    throw new TypeError('PNG dimensions must be between 1 and 4096 pixels');
  signal?.throwIfAborted();
  return { width, height };
}

export async function preloadSlidePng(
  png: Uint8Array,
  before: { api: PngApi; common: PngCommon },
  dimensions: { width: number; height: number },
  signal?: AbortSignal,
) {
  signal?.throwIfAborted();
  const { width, height } = dimensions;
  const dataUrl = `data:image/png;base64,${bytesToBase64(png)}`;
  const full = await new Promise<string>((resolve, reject) => {
    const finish = (error?: unknown, url?: string) => {
      clearTimeout(timer);
      signal?.removeEventListener('abort', abort);
      if (error) reject(error);
      else resolve(url!);
    };
    const abort = () => finish(signal?.reason ?? new Error('Image preparation cancelled'));
    const timer = setTimeout(() => finish(new Error('Image preload timed out; no image command sent')), 10000);
    signal?.addEventListener('abort', abort, { once: true });
    // Preload only. A late callback cannot dispatch a document mutation.
    try {
      before.common.sendImgUrls(before.api, [dataUrl], (result) => {
        try {
          const item = result?.[0];
          if (result?.length !== 1 || !item?.path || !item.url || item.url === 'error')
            throw new Error('PNG registration failed');
          const local = before.common.g_oDocumentUrls.imagePath2Local(item.path);
          const url = before.common.getFullImageSrc2(local);
          if (!local || before.common.g_oDocumentUrls.getImageLocal(url) !== local)
            throw new Error('PNG is not registered in this editor');
          before.api.ImageLoader!.LoadImagesWithCallback([url], () => finish(undefined, url));
        } catch (error) {
          finish(error);
        }
      });
    } catch (error) {
      finish(error);
    }
  });
  const ready = () => {
    const loaded = before.api.ImageLoader!.LoadImage(full, 1)?.Image;
    if (
      loaded?.naturalWidth !== width ||
      loaded.naturalHeight !== height ||
      !before.common.g_oDocumentUrls.getImageLocal(full)
    )
      throw new Error('PNG did not decode into the native cache');
  };
  return { full, ready };
}
