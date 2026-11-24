import { requireSlideImage } from './agent-plugin/slide-image-replacement';
import { requireEditorContext } from './agent-plugin/editor-bridge';

const LIMIT = 8 * 1024 * 1024;
const EDIT_LIMIT = 4 * 1024 * 1024;

async function imageCanvas(blob: Blob) {
  if (blob.size > 32 * 1024 * 1024) throw new Error('图片文件过大');
  const bitmap = await createImageBitmap(blob);
  try {
    const scale = Math.min(1, 1800 / Math.max(bitmap.width, bitmap.height));
    const canvas = document.createElement('canvas');
    canvas.width = Math.max(1, Math.round(bitmap.width * scale));
    canvas.height = Math.max(1, Math.round(bitmap.height * scale));
    const ctx = canvas.getContext('2d', { willReadFrequently: true });
    if (!ctx) throw new Error('无法处理图片');
    ctx.drawImage(bitmap, 0, 0, canvas.width, canvas.height);
    return { canvas, ctx };
  } finally {
    bitmap.close();
  }
}

async function png(canvas: HTMLCanvasElement) {
  const blob = await new Promise<Blob | null>((resolve) => canvas.toBlob(resolve, 'image/png'));
  if (!blob || blob.size > EDIT_LIMIT) throw new Error('生成图片超过编辑器的 4 MiB 限制');
  return blob;
}

export function readPageStyleReference() {
  const api = requireEditorContext().api as {
    getCurrentPage?(): number;
    WordControl?: { m_oLogicDocument?: { Slides?: Array<{ getBase64Img?(): string }> } };
  };
  const slides = api.WordControl?.m_oLogicDocument?.Slides;
  const screenshot = slides?.[api.getCurrentPage?.() ?? -1]?.getBase64Img?.();
  if (!screenshot?.startsWith('data:image/png;base64,')) {
    throw new Error('无法读取当前页作为参考，请改选推荐风格');
  }
  return screenshot;
}

export async function refinementReference(blob: Blob) {
  const { canvas, ctx } = await imageCanvas(blob);
  ctx.globalCompositeOperation = 'destination-over';
  ctx.fillStyle = '#000';
  ctx.fillRect(0, 0, canvas.width, canvas.height);
  return canvas.toDataURL('image/jpeg', 0.88);
}

export async function readSelectedImageReference() {
  const { common, image } = requireSlideImage();
  const url = common.getFullImageSrc2(image.blipFill!.RasterImageId!);
  if (!/^(blob:|data:image\/(?:png|jpeg|webp);base64,)/u.test(url) || !common.g_oDocumentUrls.getImageLocal(url))
    throw new Error('无法读取所选图片，请选择已嵌入 PPTX 的图片');
  const response = await fetch(url, { signal: AbortSignal.timeout(5000) });
  if (!response.ok) throw new Error('无法读取所选图片');
  const source = await response.blob();
  const blob =
    source.size <= LIMIT && ['image/png', 'image/jpeg'].includes(source.type)
      ? source
      : await png((await imageCanvas(source)).canvas);
  return new Promise<string>((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result));
    reader.onerror = () => reject(new Error('无法读取所选图片'));
    reader.readAsDataURL(blob);
  });
}

function removeGeneratedBackground(data: Uint8ClampedArray, width: number, height: number, artText: boolean) {
  let transparent = 0;
  for (let i = 3; i < data.length; i += 4) if (data[i] < 16) transparent++;
  if (transparent < width * height * 0.01) {
    const key = artText ? [0, 0, 0] : [255, 0, 255];
    const distance = (i: number) =>
      artText ? Math.max(data[i], data[i + 1], data[i + 2]) : Math.hypot(data[i] - 255, data[i + 1], data[i + 2] - 255);
    const corners = [0, (width - 1) * 4, (height - 1) * width * 4, data.length - 4];
    if (corners.filter((i) => distance(i) < (artText ? 35 : 80)).length < 3)
      throw new Error('生成结果的背景无法分离，请重新生成');
    for (let i = 0; i < data.length; i += 4) {
      const alpha = Math.max(0, Math.min(1, (distance(i) - (artText ? 20 : 48)) / (artText ? 65 : 96)));
      data[i + 3] = Math.round(data[i + 3] * alpha);
      if (alpha > 0 && alpha < 1) {
        for (let c = 0; c < 3; c++) data[i + c] = (data[i + c] - key[c] * (1 - alpha)) / alpha;
      }
    }
  }
  let left = width,
    top = height,
    right = 0,
    bottom = 0,
    visible = 0;
  for (let y = 0; y < height; y++) {
    for (let x = 0; x < width; x++) {
      if (data[(y * width + x) * 4 + 3] <= 16) continue;
      visible++;
      left = Math.min(x, left);
      top = Math.min(y, top);
      right = Math.max(right, x + 1);
      bottom = Math.max(bottom, y + 1);
    }
  }
  const fraction = visible / (width * height);
  if (fraction < 0.005 || fraction > 0.99) throw new Error('生成结果没有清晰可用的透明主体，请重新生成');
  return { left, top, width: right - left, height: bottom - top };
}

export async function prepareEditorImage(source: string, artText: boolean, aspect?: number) {
  if (!/^data:image\/(?:png|jpeg|webp);base64,[A-Za-z0-9+/=]+$/u.test(source)) throw new Error('生成结果格式无效');
  const { canvas, ctx } = await imageCanvas(await (await fetch(source)).blob());
  const pixels = ctx.getImageData(0, 0, canvas.width, canvas.height);
  const bounds = removeGeneratedBackground(pixels.data, canvas.width, canvas.height, artText);
  const output = document.createElement('canvas');
  output.width = bounds.width + 16;
  output.height = bounds.height + 16;
  output
    .getContext('2d')!
    .putImageData(pixels, 8 - bounds.left, 8 - bounds.top, bounds.left, bounds.top, bounds.width, bounds.height);
  if (!aspect) return png(output);
  const padded = document.createElement('canvas');
  padded.width = Math.max(1, Math.round(1800 * Math.min(1, aspect)));
  padded.height = Math.max(1, Math.round(1800 * Math.min(1, 1 / aspect)));
  const scale = Math.min(padded.width / output.width, padded.height / output.height);
  const width = output.width * scale,
    height = output.height * scale;
  padded.getContext('2d')!.drawImage(output, (padded.width - width) / 2, (padded.height - height) / 2, width, height);
  return png(padded);
}
