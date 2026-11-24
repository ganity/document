import { bytesToBase64 } from 'ranuts/utils';
import { requireEditorContext } from './editor-bridge';

type Layer = {
  isImage?(): boolean;
  getUniNvProps?(): { cNvPr?: { descr?: string } };
  blipFill?: { RasterImageId?: string; srcRect?: unknown; tile?: unknown };
  group?: unknown;
  spPr?: { xfrm?: { rot?: number; flipH?: boolean; flipV?: boolean } };
};
type Pixels = { sha256: string; dataUrl: string };
const pixels = new Map<string, { pending: Promise<Pixels | undefined>; size: number }>();
let cachedSize = 0;
function evict(url: string) {
  cachedSize -= pixels.get(url)?.size ?? 0;
  pixels.delete(url);
}

/** Bind supplier hints to the current local PNG, never to a technical name or an old shape id. */
export async function nativeLayerEvidence(nativeShape: unknown) {
  const shape = nativeShape as Layer;
  const description = shape.getUniNvProps?.()?.cNvPr?.descr;
  if (!shape.isImage?.() || typeof description !== 'string' || description.length > 5000) return;
  let meta: { schema?: string; name?: string; description?: string; sha256?: string };
  try {
    meta = JSON.parse(description);
  } catch {
    return;
  }
  if (
    !meta ||
    meta.schema !== 'deckwise-layer/1' ||
    !/^[a-f0-9]{64}$/.test(meta.sha256 ?? '') ||
    typeof meta.name !== 'string' ||
    !meta.name.trim() ||
    meta.name.length > 200 ||
    typeof meta.description !== 'string' ||
    !meta.description.trim() ||
    meta.description.length > 2000 ||
    /\p{Cc}/u.test(meta.name + meta.description)
  )
    return;
  const fill = shape.blipFill,
    local = fill?.RasterImageId;
  const transform = shape.spPr?.xfrm;
  if (
    !local ||
    fill.srcRect != null ||
    fill.tile != null ||
    shape.group ||
    transform?.rot ||
    transform?.flipH ||
    transform?.flipV
  )
    return;
  const common = requireEditorContext().AscCommon as unknown as {
    getFullImageSrc2(local: string): string;
    g_oDocumentUrls: { getImageLocal(url: string): string | null };
  };
  const url = common.getFullImageSrc2(local);
  // Descriptions cannot cause network access. Only this editor's registered immutable image is read.
  if (!/^(blob:|data:image\/png;base64,)/.test(url) || !common.g_oDocumentUrls.getImageLocal(url)) return;
  let entry = pixels.get(url);
  if (!entry) {
    if (pixels.size >= 32) evict(pixels.keys().next().value!);
    const pending = (async (): Promise<Pixels | undefined> => {
      try {
        const response = await fetch(url, { signal: AbortSignal.timeout(5000) });
        if (!response.ok) return;
        const blob = await response.blob();
        if (blob.size < 33 || blob.size > 32 * 1024 * 1024) return;
        const bytes = new Uint8Array(await blob.arrayBuffer());
        if (![137, 80, 78, 71, 13, 10, 26, 10].every((b, i) => bytes[i] === b)) return;
        const sha256 = Array.from(new Uint8Array(await crypto.subtle.digest('SHA-256', bytes)), (b) =>
          b.toString(16).padStart(2, '0'),
        ).join('');
        return { sha256, dataUrl: 'data:image/png;base64,' + bytesToBase64(bytes) };
      } catch {
        return;
      }
    })();
    entry = { pending, size: 0 };
    pixels.set(url, entry);
  }
  const result = await entry.pending;
  if (!result) evict(url);
  else if (pixels.get(url) === entry && entry.size === 0) {
    entry.size = result.dataUrl.length;
    cachedSize += entry.size;
    while (cachedSize > 64 * 1024 * 1024) evict(pixels.keys().next().value!);
  }
  if (
    !result ||
    result.sha256 !== meta.sha256 ||
    shape.blipFill?.RasterImageId !== local ||
    shape.getUniNvProps?.()?.cNvPr?.descr !== description
  )
    return;
  return { layerMetadata: { name: meta.name, description: meta.description }, dataUrl: result.dataUrl };
}
