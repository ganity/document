import { getEditorApi, type EditorApi } from './agent-plugin/editor-bridge';

interface Binding {
  projectId: string;
  documentId: string;
  sessionId: string;
}
interface Shape {
  getUniNvProps?(): { cNvPr?: { name?: string } };
}
interface Slide {
  cSld?: { spTree?: Shape[] };
  timing?: { getAllAnimEffects?(): any[] };
}
interface Presentation {
  Slides: Slide[];
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
        shapes: WeakMap<Shape, { sourcePageToken: string; sourceName: string }>;
        animations: WeakMap<Slide, { effects: any[]; signature: string }>;
      }
    | undefined;
  private exported: { outputSha256: string; shapes: WeakMap<Shape, { sourcePageToken: string; sourceName: string }>;
    animations: WeakMap<Slide, { effects: any[]; signature: string }> } | undefined;
  private epoch = 0;

  reset(): void {
    this.epoch++;
    this.pending = undefined;
    this.opened = undefined;
    this.exported = undefined;
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
    this.opened = { ...pending, api, logic, tokens, openedPageTokens, shapes: new WeakMap(), animations: new WeakMap() };
    this.recordObjects();
  }

  private validSlides(slides: unknown): slides is Slide[] {
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

  private animation(slide: Slide) {
    const effects = slide.timing?.getAllAnimEffects?.() ?? [];
    const signature = JSON.stringify(effects.map(effect => [effect.getObjectId?.(), effect.cTn?.presetClass,
      effect.cTn?.presetID, effect.cTn?.presetSubtype, effect.cTn?.nodeType, effect.asc_getDuration?.(), effect.asc_getDelay?.()]));
    return { effects: [...effects], signature };
  }

  private recordObjects() {
    const opened = this.opened!;
    opened.shapes = new WeakMap();
    opened.animations = new WeakMap();
    for (const slide of opened.logic.Slides) {
      const sourcePageToken = opened.tokens.get(slide)!;
      for (const shape of slide.cSld?.spTree ?? []) {
        const name = shape.getUniNvProps?.()?.cNvPr?.name;
        if (name) opened.shapes.set(shape, { sourcePageToken, sourceName: name });
      }
      opened.animations.set(slide, this.animation(slide));
    }
  }

  commitSave(receipt: { sourceSha256: string; outputSha256: string; pageTokens: string[] }) {
    const opened = this.opened;
    const current = this.context();
    if (!opened || !current || current.sourceSha256 !== receipt.sourceSha256 ||
      this.exported?.outputSha256 !== receipt.outputSha256) return false;
    opened.sourceSha256 = receipt.outputSha256;
    opened.openedPageTokens = [...receipt.pageTokens];
    // Edits made while Host writes the file remain dirty and belong to the next save.
    opened.shapes = this.exported.shapes;
    opened.animations = this.exported.animations;
    this.exported = undefined;
    return true;
  }

  captureSave() {
    const opened = this.opened;
    const context = this.context();
    if (!opened || !context) return undefined;
    const slides = [...opened.logic.Slides];
    const objects = slides.map(slide => (slide.cSld?.spTree ?? []).map(shape => ({
      ...(opened.shapes.get(shape) ?? {}), name: shape.getUniNvProps?.()?.cNvPr?.name ?? '',
    })));
    const savedShapes = new WeakMap<Shape, { sourcePageToken: string; sourceName: string }>();
    const savedAnimations = new WeakMap<Slide, { effects: any[]; signature: string }>();
    slides.forEach((slide, index) => {
      for (const shape of slide.cSld?.spTree ?? []) {
        const name = shape.getUniNvProps?.()?.cNvPr?.name;
        if (name) savedShapes.set(shape, { sourcePageToken: context.pageTokens[index], sourceName: name });
      }
      savedAnimations.set(slide, this.animation(slide));
    });
    const animationUnchanged = slides.map(slide => {
      const previous = opened.animations.get(slide), current = this.animation(slide);
      return !!previous && previous.signature === current.signature && previous.effects.length === current.effects.length &&
        previous.effects.every((effect, index) => effect === current.effects[index]);
    });
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
      this.exported = { outputSha256, shapes: savedShapes, animations: savedAnimations };
      return {
        protocol: 'deckwise-office-page-lineage/1' as const,
        ...opened.owner,
        saveRequestId,
        ...context,
        objects, animationUnchanged,
        outputSha256,
      };
    };
  }
}

export const deckwisePageLineage = new DeckWisePageLineage();
