import type { EditorApi } from './editor-bridge';
import { readSlideTextSelection, type SlideTextShape } from './slide-text-read';
import { getReadonlyMode } from '../onlyoffice/readonly';

export interface SlideTextApi extends EditorApi {
  canEdit?(): boolean;
  isSlideShow?(): boolean;
  isLongAction?(): boolean;
  getCurrentPage?(): number;
  WordControl?: {
    m_oLogicDocument?: {
      Slides?: Array<{ cSld: { spTree: SlideTextShape[] } }>;
      CanEdit?(): boolean;
      IsMasterMode?(): boolean;
      GetCurrentController?(): {
        selectedObjects?: SlideTextShape[];
        getTargetDocContent?: SlideTextShape['getDocContent'];
      };
    };
  };
}

export function requireSlideTextSelection<T extends SlideTextApi>(api: T) {
  const logic = api.WordControl?.m_oLogicDocument;
  const page = api.getCurrentPage?.();
  const slide = page === undefined ? undefined : logic?.Slides?.[page];
  const controller = logic?.GetCurrentController?.();
  const shape = controller?.selectedObjects?.[0];
  const content = shape?.getDocContent?.();
  const selected = shape ? readSlideTextSelection(shape) : undefined;
  if (
    getReadonlyMode() ||
    api.isViewMode ||
    api.canEdit?.() === false ||
    logic?.CanEdit?.() === false ||
    !api.isDocumentLoadComplete ||
    !api.isLoadFullApi ||
    api.isSlideShow?.() ||
    api.isLongAction?.() ||
    logic?.IsMasterMode?.() ||
    !slide ||
    !shape ||
    controller?.selectedObjects?.length !== 1 ||
    !slide.cSld.spTree.includes(shape) ||
    controller.getTargetDocContent?.() !== content ||
    !selected?.selectedText.length
  )
    throw new Error('Formatting requires a writable plain-text selection in one current-slide text object');
  return { api, slide, shape, content, selected, page: page! };
}
