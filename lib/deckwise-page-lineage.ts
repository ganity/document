import { getEditorApi, type EditorApi } from './agent-plugin/editor-bridge';

interface Binding {
  projectId: string;
  documentId: string;
  sessionId: string;
}
interface Presentation {
  Slides: object[];
}
type SlideApi = EditorApi & { WordControl?: { m_oLogicDocument?: Presentation } };

async function sha256(file: File): Promise<string> {
  const digest = await crypto.subtle.digest('SHA-256', await file.arrayBuffer());
  return Array.from(new Uint8Array(digest), (byte) => byte.toString(16).padStart(2, '0')).join('');
}

/** Object identity belongs to one open Office session, never a reopened file. */
export class DeckWisePageLineage {
  private pending: { owner: Binding; sourceSha256: string; previousApi: EditorApi | null } | undefined;
  private opened:
    | {
        owner: Binding;
        sourceSha256: string;
        api: SlideApi;
        logic: Presentation;
        tokens: WeakMap<object, string>;
        openedPageTokens: string[];
      }
    | undefined;
  private epoch = 0;

  reset(): void {
    this.epoch++;
    this.pending = undefined;
    this.opened = undefined;
  }

  async begin(value: unknown, file: File): Promise<() => void> {
    this.reset();
    if (
      !value ||
      typeof value !== 'object' ||
      Array.isArray(value) ||
      !['projectId', 'documentId', 'sessionId'].every((key) => {
        const id = Reflect.get(value, key);
        return typeof id === 'string' && id.length > 0 && id.length <= 128;
      })
    )
      throw new Error('Invalid page lineage binding');
    const owner: Binding = {
      projectId: Reflect.get(value, 'projectId'),
      documentId: Reflect.get(value, 'documentId'),
      sessionId: Reflect.get(value, 'sessionId'),
    };
    const previousApi = getEditorApi();
    const epoch = this.epoch;
    const sourceSha256 = await sha256(file);
    if (this.epoch !== epoch) throw new Error('Page lineage open superseded');
    this.pending = {
      owner: { projectId: owner.projectId, documentId: owner.documentId, sessionId: owner.sessionId },
      sourceSha256,
      previousApi,
    };
    return () => {
      if (this.epoch === epoch) this.contentReady();
    };
  }

  /** Capture synchronously from this open operation, before the first edit. */
  contentReady(): void {
    const pending = this.pending;
    this.pending = undefined;
    if (!pending) return;
    const api = getEditorApi() as SlideApi | null;
    const logic = api?.WordControl?.m_oLogicDocument;
    if (
      !api ||
      api === pending.previousApi ||
      !api.isDocumentLoadComplete ||
      !api.isLoadFullApi ||
      !logic ||
      !this.validSlides(logic.Slides)
    )
      return;
    const tokens = new WeakMap<object, string>();
    const openedPageTokens = logic.Slides.map((slide) => {
      const token = crypto.randomUUID();
      tokens.set(slide, token);
      return token;
    });
    this.opened = { ...pending, api, logic, tokens, openedPageTokens };
  }

  private validSlides(slides: unknown): slides is object[] {
    return (
      Array.isArray(slides) &&
      slides.length > 0 &&
      slides.length <= 5000 &&
      slides.every((slide) => slide !== null && typeof slide === 'object') &&
      new Set(slides).size === slides.length
    );
  }

  context() {
    const opened = this.opened;
    if (
      !opened ||
      getEditorApi() !== opened.api ||
      opened.api.WordControl?.m_oLogicDocument !== opened.logic ||
      !this.validSlides(opened.logic.Slides)
    )
      return undefined;
    const pageTokens = opened.logic.Slides.map((slide) => {
      let token = opened.tokens.get(slide);
      if (!token) {
        token = crypto.randomUUID();
        opened.tokens.set(slide, token);
      }
      return token;
    });
    return { sourceSha256: opened.sourceSha256, openedPageTokens: [...opened.openedPageTokens], pageTokens };
  }

  captureSave() {
    const opened = this.opened;
    const context = this.context();
    if (!opened || !context) return undefined;
    const slides = [...opened.logic.Slides];
    return async (file: File, saveRequestId: string, isCurrent: () => boolean) => {
      const outputSha256 = await sha256(file);
      if (
        this.opened !== opened ||
        !this.context() ||
        !isCurrent() ||
        opened.logic.Slides.length !== slides.length ||
        slides.some((slide, index) => opened.logic.Slides[index] !== slide)
      )
        return undefined;
      return {
        protocol: 'deckwise-office-page-lineage/1' as const,
        ...opened.owner,
        saveRequestId,
        ...context,
        outputSha256,
      };
    };
  }
}

export const deckwisePageLineage = new DeckWisePageLineage();
