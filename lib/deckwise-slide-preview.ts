import { captureDocumentToolTarget } from './agent-plugin/document-tool-action';
import { getEditorApi, requireEditorContext, type EditorApi } from './agent-plugin/editor-bridge';
import { deckwiseEditorBridge } from './deckwise-editor-bridge';
import { deckwisePageLineage } from './deckwise-page-lineage';
import { getReadonlyMode } from './onlyoffice/readonly';
import { hasUnsavedChanges } from './unsaved-guard';

interface PreviewSlide {
  draw?(graphics: PreviewGraphics): void;
  cSld?: { spTree?: PreviewShape[] };
  backgroundFill?: PreviewFill;
  Layout?: { cSld?: { spTree?: PreviewShape[] }; Master?: { cSld?: { spTree?: PreviewShape[] } } };
}
interface PreviewFill {
  transparent?: number | null;
}
interface PreviewShape {
  spTree?: PreviewShape[];
  getObjectType?(): number;
  brush?: PreviewFill;
  cropBrush?: PreviewFill;
  pen?: { Fill?: PreviewFill };
  getOuterShdw?(): unknown;
  txWarpStruct?: unknown;
  txWarpStructParamarks?: unknown;
  txWarpStructNoTransform?: unknown;
  txWarpStructParamarksNoTransform?: unknown;
  txWarpStruct2?: unknown;
  textBoxContent?: unknown;
  signatureLine?: unknown;
}
interface PreviewGraphics {
  init(context: CanvasRenderingContext2D, width: number, height: number, widthMM: number, heightMM: number): void;
  transform(sx: number, shy: number, shx: number, sy: number, tx: number, ty: number): void;
  m_oFontManager?: unknown;
  isSupportTextDraw?(): boolean;
}
interface PreviewRenderer {
  CGraphics?: new () => PreviewGraphics;
  IsShapeToImageConverter?: boolean;
  g_fontManager?: { m_pFont: unknown };
  g_fontManager2?: { m_pFont: unknown };
}
interface PreviewPresentation {
  Slides?: PreviewSlide[];
  GetWidthMM?(): number;
  GetHeightMM?(): number;
}
interface PreviewApi extends EditorApi {
  getCurrentPage?(): number;
  WordControl?: { m_oLogicDocument?: PreviewPresentation };
}
/** Render a saved page without navigating or mutating its opened session. */
export function renderDeckWiseCurrentSlidePreview(
  value: unknown,
  expectedGeneration: number,
  omitNativeText = false,
  pageNumber?: number,
) {
  const target = captureDocumentToolTarget();
  const api = getEditorApi() as PreviewApi | null;
  const context = deckwisePageLineage.context();
  const latest = deckwiseEditorBridge.getContext(value);
  const editor = requireEditorContext();
  const common = editor.AscCommon;
  const history = common?.History;
  if (
    !api ||
    typeof omitNativeText !== 'boolean' ||
    target.context.kind !== 'slide' ||
    !Number.isSafeInteger(expectedGeneration) ||
    expectedGeneration < 1 ||
    latest.editGeneration !== expectedGeneration ||
    !getReadonlyMode() ||
    hasUnsavedChanges() ||
    history?.Index !== -1 ||
    history.Points.length !== 0 ||
    !target.isCurrent(false) ||
    !context ||
    !/^[a-f0-9]{64}$/.test(context.sourceSha256)
  )
    throw new Error('A current bound PPTX page and source identity are required');

  const currentPage = api.getCurrentPage?.();
  if (pageNumber !== undefined && (!Number.isSafeInteger(pageNumber) || pageNumber < 1))
    throw new Error('Invalid preview page');
  const page = pageNumber === undefined ? currentPage : pageNumber - 1;
  const logic = api.WordControl?.m_oLogicDocument;
  const slides = logic?.Slides;
  const slide = Number.isSafeInteger(page) && page! >= 0 ? slides?.[page!] : undefined;
  const pageToken = context.pageTokens[page!];
  if (!slide || !pageToken || !target.isCurrent(false))
    throw new Error('Current page changed; read fresh editor context');
  if (omitNativeText) {
    const types = editor.AscDFH as
      | { historyitem_type_Shape?: number; historyitem_type_ImageShape?: number; historyitem_type_GroupShape?: number }
      | undefined;
    const trees = [slide.cSld?.spTree, slide.Layout?.cSld?.spTree, slide.Layout?.Master?.cSld?.spTree];
    const opaque = (fill: PreviewFill | undefined) => fill?.transparent == null || fill.transparent === 255;
    let count = 0;
    const safe = (shape: PreviewShape | undefined, depth = 0): boolean => {
      if (
        !shape ||
        ++count > 1000 ||
        depth > 16 ||
        !opaque(shape.brush) ||
        !opaque(shape.cropBrush) ||
        !opaque(shape.pen?.Fill) ||
        shape.getOuterShdw?.() ||
        shape.txWarpStruct ||
        shape.txWarpStructParamarks ||
        shape.txWarpStructNoTransform ||
        shape.txWarpStructParamarksNoTransform ||
        shape.txWarpStruct2 ||
        shape.textBoxContent ||
        shape.signatureLine
      )
        return false;
      const type = shape.getObjectType?.();
      if (type === types?.historyitem_type_GroupShape && Number.isSafeInteger(type))
        return Array.isArray(shape.spTree) && shape.spTree.every((child) => safe(child, depth + 1));
      return [types!.historyitem_type_Shape, types!.historyitem_type_ImageShape].includes(type);
    };
    // shortcut: warped text and translucent fills/shadows still require a separately reviewed clean plate.
    if (
      !Number.isSafeInteger(types?.historyitem_type_Shape) ||
      !Number.isSafeInteger(types?.historyitem_type_ImageShape) ||
      trees.some((tree) => !Array.isArray(tree)) ||
      !opaque(slide.backgroundFill) ||
      trees.flat().some((shape) => !safe(shape))
    )
      throw new Error('Native text omission requires ordinary opaque shapes and pictures; review another clean plate');
  }
  const renderer = common as typeof common & PreviewRenderer;
  const widthMM = logic?.GetWidthMM?.();
  const heightMM = logic?.GetHeightMM?.();
  if (
    typeof slide.draw !== 'function' ||
    typeof renderer?.CGraphics !== 'function' ||
    ![widthMM, heightMM].every(
      (dimension) => typeof dimension === 'number' && Number.isFinite(dimension) && dimension > 0,
    )
  )
    throw new Error('Native page preview is unavailable');
  const scale = 1600 / Math.max(widthMM!, heightMM!);
  const width = Math.round(widthMM! * scale);
  const height = Math.round(heightMM! * scale);
  if ([width, height].some((dimension) => !Number.isSafeInteger(dimension) || dimension < 64 || dimension > 1600))
    throw new Error('Page aspect ratio exceeds the preview budget');
  const canvas = document.createElement('canvas');
  canvas.width = width;
  canvas.height = height;
  const canvasContext = canvas.getContext('2d');
  if (!canvasContext) throw new Error('Native page preview canvas is unavailable');
  const previousRendering = renderer.IsShapeToImageConverter;
  let dataUrl: string;
  try {
    const graphics = new renderer.CGraphics();
    graphics.init(canvasContext, width, height, widthMM!, heightMM!);
    graphics.m_oFontManager = renderer.g_fontManager;
    graphics.transform(1, 0, 0, 1, 0, 0);
    if (omitNativeText) graphics.isSupportTextDraw = () => false;
    renderer.IsShapeToImageConverter = true;
    slide.draw(graphics);
    dataUrl = canvas.toDataURL('image/png');
  } finally {
    renderer.IsShapeToImageConverter = previousRendering;
    // The native thumbnail renderer also clears the shared font cache after drawing.
    if (renderer.g_fontManager) renderer.g_fontManager.m_pFont = null;
    if (renderer.g_fontManager2) renderer.g_fontManager2.m_pFont = null;
    canvas.width = canvas.height = 0;
  }
  if (typeof dataUrl !== 'string' || !dataUrl.startsWith('data:image/png;base64,'))
    throw new Error('Native page preview did not return a PNG');
  const binary = atob(dataUrl.slice('data:image/png;base64,'.length));
  if (binary.length < 33 || binary.length > 8 * 1024 * 1024)
    throw new Error('Native page preview exceeds its PNG budget');
  const png = Uint8Array.from(binary, (character) => character.charCodeAt(0));
  if (
    png[0] !== 0x89 ||
    png[1] !== 0x50 ||
    png[2] !== 0x4e ||
    png[3] !== 0x47 ||
    api !== getEditorApi() ||
    api.WordControl?.m_oLogicDocument?.Slides?.[page!] !== slide ||
    api.getCurrentPage?.() !== currentPage ||
    !target.isCurrent(false) ||
    !getReadonlyMode() ||
    hasUnsavedChanges() ||
    history.Index !== -1 ||
    history.Points.length !== 0 ||
    deckwisePageLineage.context()?.pageTokens[page!] !== pageToken
  )
    throw new Error('PPTX page changed while rendering; discard this preview');
  return {
    sourceSha256: context.sourceSha256,
    pageToken,
    page: page! + 1,
    mime: 'image/png' as const,
    png,
    ...(omitNativeText ? { nativeTextOmitted: true as const } : {}),
  };
}
