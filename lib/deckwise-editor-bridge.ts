import {
  captureDocumentToolTarget,
  DocumentToolAction,
  type DocumentToolTarget,
} from './agent-plugin/document-tool-action';
import { documentTools, parseDocumentToolPlan } from './agent-plugin/document-tool-plan';
import { getReadonlyMode } from './onlyoffice/readonly';
import {
  replaceSlideImageFromAsset,
  insertSlideImageFromAsset,
  type SlideImageAsset,
} from './agent-plugin/slide-image-asset';
import { canReplaceSlideImage } from './agent-plugin/slide-image-replacement';
import { canInsertSlideImage } from './agent-plugin/slide-image-insertion';
import { canSetSlideFade, prepareSlideFade } from './agent-plugin/slide-fade';
import { canSetSlideImageSize, setSlideImageSize } from './agent-plugin/slide-image-size';
import { canSetSlideImagePosition, setSlideImagePosition } from './agent-plugin/slide-image-position';
import { canAddSlideObjectAnimation, addSlideObjectAnimation } from './agent-plugin/slide-object-animation';
import { readDeckWiseSelectedObjects } from './deckwise-object-selection';

export interface DeckWiseEditorBinding {
  projectId: string;
  documentId: string;
  sessionId: string;
}

type Binding = DeckWiseEditorBinding;

export interface DeckWiseImageAssetSource {
  list(binding: DeckWiseEditorBinding): readonly SlideImageAsset[];
  load(assetId: string, signal?: AbortSignal): Promise<Uint8Array>;
}

function object(value: unknown): value is Record<string, unknown> {
  return !!value && typeof value === 'object' && !Array.isArray(value);
}

function binding(value: unknown): Binding {
  if (
    !object(value) ||
    !['projectId', 'documentId', 'sessionId'].every(
      (key) => typeof value[key] === 'string' && value[key].length > 0 && value[key].length <= 128,
    )
  )
    throw new Error('Invalid editor binding');
  return {
    projectId: String(value.projectId),
    documentId: String(value.documentId),
    sessionId: String(value.sessionId),
  };
}

function validImageAsset(value: unknown): value is SlideImageAsset {
  return (
    object(value) &&
    Object.keys(value).every((key) => ['assetId', 'sha256', 'mime', 'byteLength', 'width', 'height'].includes(key)) &&
    typeof value.assetId === 'string' &&
    /^[A-Za-z0-9_-]{1,128}$/.test(value.assetId) &&
    typeof value.sha256 === 'string' &&
    /^[a-f0-9]{64}$/.test(value.sha256) &&
    value.mime === 'image/png' &&
    Number.isSafeInteger(value.byteLength) &&
    Number(value.byteLength) >= 33 &&
    Number(value.byteLength) <= 4 * 1024 * 1024 &&
    Number.isSafeInteger(value.width) &&
    Number(value.width) >= 1 &&
    Number(value.width) <= 4096 &&
    Number.isSafeInteger(value.height) &&
    Number(value.height) >= 1 &&
    Number(value.height) <= 4096
  );
}

function fadeAuthorization(value: unknown) {
  if (
    !object(value) ||
    Object.keys(value).length !== 4 ||
    Object.keys(value).some((key) => !['editGeneration', 'page', 'sourceHash', 'slideHash'].includes(key)) ||
    !Number.isSafeInteger(value.editGeneration) ||
    Number(value.editGeneration) < 1 ||
    !Number.isSafeInteger(value.page) ||
    Number(value.page) < 1 ||
    ![value.sourceHash, value.slideHash].every((hash) => typeof hash === 'string' && /^sha256:[a-f0-9]{64}$/.test(hash))
  )
    throw new Error('Invalid host fade authorization');
  return {
    editGeneration: Number(value.editGeneration),
    page: Number(value.page),
    sourceHash: String(value.sourceHash),
    slideHash: String(value.slideHash),
  };
}

/** Binds native tool snapshots to one host-owned document session. No model runs here. */
export class DeckWiseEditorBridge {
  private owner: Binding | undefined;
  private generation = 0;
  private snapshot: { target: DocumentToolTarget; readOnly: boolean; assets: SlideImageAsset[] } | undefined;
  private imageAssetSource: DeckWiseImageAssetSource | undefined;
  private reviewedImage: SlideImageAsset | undefined;
  private reviewedBytes: Uint8Array | undefined;
  private fadeApproval:
    | { snapshot: NonNullable<DeckWiseEditorBridge['snapshot']>; proof: ReturnType<typeof fadeAuthorization> }
    | undefined;
  private fadeDispatch: DeckWiseEditorBridge['fadeApproval'];
  private busy = false;
  private activeCommand: AbortController | undefined;
  private readonly usedRequests = new Set<string>();

  reset(value?: unknown): void {
    this.activeCommand?.abort(new Error('Editor session has expired'));
    this.owner = undefined;
    this.snapshot = undefined;
    this.generation = 0;
    this.usedRequests.clear();
    this.imageAssetSource = undefined;
    this.reviewedImage = undefined;
    this.reviewedBytes = undefined;
    this.fadeApproval = undefined;
    this.fadeDispatch = undefined;
    if (value !== undefined) this.owner = binding(value);
  }

  setImageAssetSource(source?: DeckWiseImageAssetSource): void {
    this.activeCommand?.abort(new Error('Image asset source has changed; read fresh context'));
    this.imageAssetSource = source;
    this.snapshot = undefined;
  }

  reviewImage(value: unknown, asset: unknown) {
    const owner = this.assertOwner(value);
    if (this.busy) throw new Error('Editor command is in progress');
    if (asset !== null && !validImageAsset(asset)) throw new Error('Invalid reviewed PNG metadata');
    this.reviewedImage = asset === null ? undefined : { ...asset };
    this.reviewedBytes = undefined;
    this.setImageAssetSource(
      this.reviewedImage
        ? {
            list: () => (this.reviewedImage ? [{ ...this.reviewedImage }] : []),
            load: async (assetId) => {
              if (!this.reviewedBytes || assetId !== this.reviewedImage?.assetId)
                throw new Error('Fresh authorized image bytes are required');
              return new Uint8Array(this.reviewedBytes);
            },
          }
        : undefined,
    );
    return this.getContext(owner);
  }

  reviewImageTarget(value: unknown, asset: unknown, target: unknown) {
    const owner = this.assertOwner(value);
    if (
      !object(target) ||
      Object.keys(target).sort().join(',') !== 'editGeneration,operation,page' ||
      !Number.isSafeInteger(target.page) ||
      Number(target.page) < 1 ||
      !Number.isSafeInteger(target.editGeneration) ||
      Number(target.editGeneration) < 1 ||
      (target.operation !== 'replace' && target.operation !== 'insert')
    )
      throw new Error('Invalid image target');
    const before = this.getContext(owner),
      originalTarget = this.snapshot!.target;
    if (before.readOnly || before.page !== target.page || before.editGeneration !== target.editGeneration)
      throw new Error('Image review requires fresh writable native context');
    if (!validImageAsset(asset)) throw new Error('Invalid reviewed PNG metadata');
    try {
      const context = this.reviewImage(owner, asset);
      if (
        context.readOnly ||
        context.page !== target.page ||
        !originalTarget.isCurrent(true) ||
        !context.capabilities.some((capability) => capability.name === `${target.operation}_slide_image_from_asset`)
      )
        throw new Error('Image review requires a supported image target');
      return context;
    } catch (error) {
      this.reviewedImage = undefined;
      this.reviewedBytes = undefined;
      this.setImageAssetSource();
      throw error;
    }
  }

  async executeImage(value: unknown, bytes: unknown, messageId?: string) {
    const image = this.reviewedImage;
    if (
      !object(value) ||
      !['replace_slide_image_from_asset', 'insert_slide_image_from_asset'].includes(String(value.tool)) ||
      !object(value.input) ||
      !image ||
      value.input.assetId !== image.assetId ||
      !(bytes instanceof Uint8Array) ||
      bytes.byteLength !== image.byteLength ||
      this.busy
    )
      throw new Error('Invalid reviewed image execution');
    this.reviewedBytes = new Uint8Array(bytes);
    try {
      return await this.execute(value, messageId);
    } finally {
      this.reviewedBytes = undefined;
    }
  }

  /** Trusted Host only: retain B81 inspection privately and bind its exact saved source to this live context. */
  revokeFade(): void {
    this.fadeApproval = undefined;
    this.fadeDispatch = undefined;
  }

  reviewFade(value: unknown, authorization: unknown) {
    const owner = this.assertOwner(value);
    if (this.busy) throw new Error('Editor command is in progress');
    this.fadeApproval = undefined;
    const context = this.getContext(owner);
    if (authorization === null) return context;
    const proof = fadeAuthorization(authorization);
    if (
      proof.editGeneration !== context.editGeneration ||
      proof.page !== context.page ||
      context.readOnly ||
      !canSetSlideFade()
    )
      throw new Error('Fade authorization requires fresh writable native context');
    this.fadeApproval = { snapshot: this.snapshot!, proof };
    return this.getContext(owner);
  }

  /** Authorization is transport-only; ordinary model commands cannot supply or reuse it. */
  async executeFade(value: unknown, authorization: unknown, messageId?: string) {
    this.assertOwner(value);
    if (this.busy) throw new Error('Editor command is in progress');
    const approval = this.fadeApproval;
    this.fadeApproval = undefined;
    const proof = fadeAuthorization(authorization);
    if (
      !object(value) ||
      value.tool !== 'set_slide_fade' ||
      !approval ||
      JSON.stringify(proof) !== JSON.stringify(approval.proof)
    )
      throw new Error('Fresh host fade authorization is required');
    this.fadeDispatch = approval;
    try {
      return await this.execute(value, messageId);
    } finally {
      this.fadeApproval = undefined;
      this.fadeDispatch = undefined;
    }
  }

  get isExecuting(): boolean {
    return this.busy;
  }

  private currentImageAssets(owner: Binding, target: DocumentToolTarget, readOnly: boolean): SlideImageAsset[] {
    if (
      readOnly ||
      !this.imageAssetSource ||
      !target.isCurrent(true) ||
      (!canReplaceSlideImage() && !canInsertSlideImage())
    )
      return [];
    const values = this.imageAssetSource.list({ ...owner });
    if (!Array.isArray(values) || values.length > 100 || values.some((asset) => !validImageAsset(asset)))
      throw new Error('Invalid authorized image asset list');
    const assets = values
      .map(({ assetId, sha256, mime, byteLength, width, height }) => ({
        assetId,
        sha256,
        mime,
        byteLength,
        width,
        height,
      }))
      .sort((a, b) => a.assetId.localeCompare(b.assetId));
    if (new Set(assets.map((asset) => asset.assetId)).size !== assets.length)
      throw new Error('Duplicate authorized image asset ID');
    return assets;
  }

  private assertOwner(value: unknown): Binding {
    const expected = binding(value);
    if (
      !this.owner ||
      this.owner.projectId !== expected.projectId ||
      this.owner.documentId !== expected.documentId ||
      this.owner.sessionId !== expected.sessionId
    ) {
      throw new Error('Editor session has expired; read fresh context');
    }
    return expected;
  }

  getContext(value: unknown) {
    const owner = this.assertOwner(value);
    if (this.busy) throw new Error('Editor command is in progress');
    const readOnly = getReadonlyMode();
    if (!this.snapshot || this.snapshot.readOnly !== readOnly || !this.snapshot.target.isCurrent(false)) {
      const target = captureDocumentToolTarget();
      this.snapshot = { target, readOnly, assets: this.currentImageAssets(owner, target, readOnly) };
      this.generation++;
    } else {
      const assets = this.currentImageAssets(owner, this.snapshot.target, readOnly);
      if (JSON.stringify(assets) !== JSON.stringify(this.snapshot.assets)) {
        this.snapshot = { ...this.snapshot, assets };
        this.generation++;
      }
    }
    const { target } = this.snapshot;
    if (target.context.kind !== 'slide') throw new Error('DeckWise bridge requires a presentation');
    const capabilities = documentTools(target.context)
      .filter(
        (tool) =>
          [
            'get_presentation_text',
            'replace_selection',
            'add_slide_text',
            'set_slide_font_family',
            'set_slide_font_size',
            'set_slide_text_bold',
            'set_slide_text_italic',
            'set_slide_text_color',
            'set_slide_text_box_position',
            'set_slide_text_box_size',
          ].includes(tool.name) &&
          (tool.readOnlyHint || (!readOnly && target.isCurrent(true))),
      )
      .map((tool) => ({
        name: tool.name,
        description: tool.description,
        inputSchema: tool.inputSchema,
        readOnly: tool.readOnlyHint,
      }));
    if (!readOnly && target.isCurrent(true) && canSetSlideImageSize())
      capabilities.push({
        name: 'set_slide_image_size',
        description:
          'Resize one selected ordinary unrotated, unflipped, uncropped image using explicit on-page millimeter dimensions. Preserves its position and asset; does not save.',
        inputSchema: {
          type: 'object',
          properties: {
            widthMm: { type: 'number', minimum: 1, maximum: 1000 },
            heightMm: { type: 'number', minimum: 1, maximum: 1000 },
          },
          required: ['widthMm', 'heightMm'],
          additionalProperties: false,
        },
        readOnly: false,
      });
    if (!readOnly && target.isCurrent(true) && canSetSlideImagePosition())
      capabilities.push({
        name: 'set_slide_image_position',
        description:
          'Move one selected ordinary unrotated, unflipped, uncropped image to explicit on-page top-left millimeter coordinates. Preserves dimensions and asset; does not save.',
        inputSchema: {
          type: 'object',
          properties: {
            xMm: { type: 'number', minimum: 0, maximum: 1000 },
            yMm: { type: 'number', minimum: 0, maximum: 1000 },
          },
          required: ['xMm', 'yMm'],
          additionalProperties: false,
        },
        readOnly: false,
      });
    if (!readOnly && target.isCurrent(true) && canAddSlideObjectAnimation())
      capabilities.push({
        name: 'add_slide_object_animation',
        description:
          'Append one native on-click appear or fade entrance to the selected ordinary object, preserving existing supported effects. Does not preview or save.',
        inputSchema: {
          type: 'object',
          properties: { effect: { type: 'string', enum: ['appear', 'fade'] } },
          required: ['effect'],
          additionalProperties: false,
        },
        readOnly: false,
      });
    if (this.snapshot.assets.length && canReplaceSlideImage())
      capabilities.push({
        name: 'replace_slide_image_from_asset',
        description: 'Replace the selected ordinary slide image with one of the currently authorized PNG assets.',
        inputSchema: {
          type: 'object',
          properties: { assetId: { type: 'string', enum: this.snapshot.assets.map((asset) => asset.assetId) } },
          required: ['assetId'],
          additionalProperties: false,
        },
        readOnly: false,
      });
    if (this.snapshot.assets.length && canInsertSlideImage())
      capabilities.push({
        name: 'insert_slide_image_from_asset',
        description:
          'Insert one reviewed PNG asset on the current slide at explicit on-page millimeter geometry. Empty canvas selection required; review stretch and overlap. Does not save.',
        inputSchema: {
          type: 'object',
          properties: {
            assetId: { type: 'string', enum: this.snapshot.assets.map((asset) => asset.assetId) },
            xMm: { type: 'number', minimum: 0, maximum: 1000 },
            yMm: { type: 'number', minimum: 0, maximum: 1000 },
            widthMm: { type: 'number', minimum: 1, maximum: 1000 },
            heightMm: { type: 'number', minimum: 1, maximum: 1000 },
          },
          required: ['assetId', 'xMm', 'yMm', 'widthMm', 'heightMm'],
          additionalProperties: false,
        },
        readOnly: false,
      });
    if (
      this.fadeApproval &&
      this.fadeApproval.snapshot === this.snapshot &&
      this.fadeApproval.proof.editGeneration === this.generation &&
      !readOnly &&
      canSetSlideFade()
    )
      capabilities.push({
        name: 'set_slide_fade',
        description:
          'Set only the current slide to a 750 ms smooth click-advanced fade after Host source inspection. Does not save.',
        inputSchema: { type: 'object', properties: {}, additionalProperties: false },
        readOnly: false,
      });
    else this.fadeApproval = undefined;
    return {
      ...owner,
      editGeneration: this.generation,
      kind: 'slide' as const,
      page: target.context.page,
      selectedText: target.selectedText.slice(0, 8000),
      selectionCharacters: target.selectedText.length,
      selectedObjects: readDeckWiseSelectedObjects(),
      readOnly,
      capabilities,
    };
  }

  async execute(value: unknown, messageId?: string) {
    const owner = this.assertOwner(value);
    if (
      !object(value) ||
      typeof value.requestId !== 'string' ||
      !/^[\w-]{1,128}$/.test(value.requestId) ||
      value.requestId !== messageId ||
      !Number.isSafeInteger(value.editGeneration) ||
      Object.keys(value).some(
        (key) =>
          !['projectId', 'documentId', 'sessionId', 'requestId', 'editGeneration', 'tool', 'input'].includes(key),
      )
    ) {
      throw new Error('Invalid editor command');
    }
    if (this.busy) throw new Error('Editor command is in progress');
    if (this.usedRequests.has(value.requestId)) throw new Error('Editor request was already consumed');
    if (this.usedRequests.size >= 4096) throw new Error('Editor request limit reached; reopen the document');
    const snapshot = this.snapshot;
    if (
      !snapshot ||
      value.editGeneration !== this.generation ||
      snapshot.readOnly !== getReadonlyMode() ||
      !snapshot.target.isCurrent(false)
    ) {
      throw new Error('Editor context has changed; read fresh context');
    }
    if (
      ![
        'get_presentation_text',
        'replace_selection',
        'add_slide_text',
        'set_slide_font_family',
        'set_slide_font_size',
        'set_slide_text_bold',
        'set_slide_text_italic',
        'set_slide_text_color',
        'set_slide_text_box_position',
        'set_slide_text_box_size',
        'set_slide_image_size',
        'set_slide_image_position',
        'add_slide_object_animation',
        'replace_slide_image_from_asset',
        'insert_slide_image_from_asset',
        'set_slide_fade',
      ].includes(String(value.tool))
    ) {
      throw new Error('This tool is not enabled in the DeckWise bridge');
    }
    this.usedRequests.add(value.requestId);
    this.busy = true;
    this.snapshot = undefined;
    this.fadeApproval = undefined;
    const abort = new AbortController();
    this.activeCommand = abort;
    try {
      if (value.tool === 'set_slide_image_size') {
        if (getReadonlyMode() || !snapshot.target.isCurrent(true) || !canSetSlideImageSize())
          throw new Error('Image resize requires fresh writable native context');
        const result = setSlideImageSize(value.input, abort.signal);
        this.assertOwner(owner);
        return {
          ...owner,
          requestId: value.requestId,
          editGeneration: value.editGeneration,
          status: 'verified' as const,
          tool: 'set_slide_image_size' as const,
          result,
        };
      }
      if (value.tool === 'set_slide_image_position') {
        if (getReadonlyMode() || !snapshot.target.isCurrent(true) || !canSetSlideImagePosition())
          throw new Error('Image positioning requires fresh writable native context');
        const result = setSlideImagePosition(value.input, abort.signal);
        this.assertOwner(owner);
        return {
          ...owner,
          requestId: value.requestId,
          editGeneration: value.editGeneration,
          status: 'verified' as const,
          tool: 'set_slide_image_position' as const,
          result,
        };
      }
      if (value.tool === 'add_slide_object_animation') {
        if (getReadonlyMode() || !snapshot.target.isCurrent(true) || !canAddSlideObjectAnimation())
          throw new Error('Object animation requires fresh writable native context');
        const result = addSlideObjectAnimation(value.input, abort.signal);
        this.assertOwner(owner);
        return {
          ...owner,
          requestId: value.requestId,
          editGeneration: value.editGeneration,
          status: 'verified' as const,
          tool: 'add_slide_object_animation' as const,
          result,
        };
      }
      if (value.tool === 'set_slide_fade') {
        const approval = this.fadeDispatch;
        if (
          !approval ||
          approval.snapshot !== snapshot ||
          approval.proof.editGeneration !== value.editGeneration ||
          approval.proof.page !== snapshot.target.context.page ||
          !canSetSlideFade() ||
          !object(value.input) ||
          Object.keys(value.input).length !== 0
        )
          throw new Error('Fresh host fade authorization and empty input are required');
        const prepared = prepareSlideFade(abort.signal);
        if (prepared.page !== approval.proof.page) throw new Error('Fade page changed; read fresh context');
        const result = prepared.apply();
        this.assertOwner(owner);
        return {
          ...owner,
          requestId: value.requestId,
          editGeneration: value.editGeneration,
          status: 'verified' as const,
          tool: 'set_slide_fade' as const,
          result,
        };
      }
      if (value.tool === 'replace_slide_image_from_asset' || value.tool === 'insert_slide_image_from_asset') {
        const insertion = value.tool === 'insert_slide_image_from_asset';
        const tool = value.tool,
          requestId = value.requestId,
          editGeneration = Number(value.editGeneration);
        if (
          Number(value.editGeneration) < 1 ||
          !object(value.input) ||
          Object.keys(value.input).length !== (insertion ? 5 : 1) ||
          Object.keys(value.input).some(
            (key) => !(insertion ? ['assetId', 'xMm', 'yMm', 'widthMm', 'heightMm'] : ['assetId']).includes(key),
          ) ||
          typeof value.input.assetId !== 'string' ||
          !/^[A-Za-z0-9_-]{1,128}$/.test(value.input.assetId)
        ) {
          throw new Error('Invalid image asset command');
        }
        const assetId = value.input.assetId;
        const asset = snapshot.assets.find((item) => item.assetId === assetId);
        const source = this.imageAssetSource;
        if (!asset || !source) throw new Error('Image asset is not authorized by the latest editor context');
        const box = insertion
          ? {
              xMm: value.input.xMm,
              yMm: value.input.yMm,
              widthMm: value.input.widthMm,
              heightMm: value.input.heightMm,
            }
          : undefined;
        if (
          box &&
          (![box.xMm, box.yMm, box.widthMm, box.heightMm].every(
            (v) => typeof v === 'number' && Number.isFinite(v) && v <= 1000,
          ) ||
            Number(box.xMm) < 0 ||
            Number(box.yMm) < 0 ||
            Number(box.widthMm) < 1 ||
            Number(box.heightMm) < 1 ||
            !canInsertSlideImage())
        )
          throw new Error('Invalid or unavailable native image insertion geometry');
        const readCurrent = () => ({
          ...(this.owner ?? { projectId: '', documentId: '', sessionId: '' }),
          editGeneration: this.generation,
          readOnly: getReadonlyMode(),
          saving: false,
          imageReplacementInstalled:
            this.imageAssetSource === source &&
            JSON.stringify(this.currentImageAssets(owner, snapshot.target, getReadonlyMode())) ===
              JSON.stringify(snapshot.assets),
        });
        const common = {
          binding: { ...owner, editGeneration },
          asset,
          loadAsset: (assetId: string, signal?: AbortSignal) =>
            assetId === asset.assetId
              ? source.load(assetId, signal)
              : Promise.reject(new Error('Image asset ID changed')),
          signal: abort.signal,
        };
        const result = box
          ? await insertSlideImageFromAsset({
              ...common,
              box: {
                xMm: Number(box.xMm),
                yMm: Number(box.yMm),
                widthMm: Number(box.widthMm),
                heightMm: Number(box.heightMm),
              },
              readCurrent: () => {
                const now = readCurrent();
                return { ...now, imageInsertionInstalled: now.imageReplacementInstalled && canInsertSlideImage() };
              },
            })
          : await replaceSlideImageFromAsset({ ...common, readCurrent });
        this.assertOwner(owner);
        return {
          ...owner,
          requestId,
          editGeneration,
          status: 'verified' as const,
          tool,
          result: {
            assetId: result.assetId,
            sha256: result.sha256,
            imageId: result.imageId,
            ...(box ? { xMm: box.xMm, yMm: box.yMm, widthMm: box.widthMm, heightMm: box.heightMm } : {}),
            page: result.page,
            verified: true as const,
          },
        };
      }
      const plan = parseDocumentToolPlan(
        JSON.stringify({ tool: value.tool, input: value.input }),
        snapshot.target.context,
      );
      const action = new DocumentToolAction(snapshot.target, plan);
      const status = await action.apply(abort.signal);
      this.assertOwner(owner);
      if (plan.readOnly) {
        if (!snapshot.target.isCurrent(false) || abort.signal.aborted)
          throw new Error('Editor changed while reading; read fresh context');
        // A verified read must not expire an approval for an otherwise unchanged target.
        if (snapshot.readOnly === getReadonlyMode() && this.generation === value.editGeneration)
          this.snapshot = snapshot;
      } else if (status !== 'verified' || !action.continuationTarget?.isCurrent(false)) {
        throw new Error('Editor change could not be verified; inspect the document before retrying');
      }
      return {
        ...owner,
        requestId: value.requestId,
        editGeneration: value.editGeneration,
        status,
        tool: plan.tool,
        result: action.result,
      };
    } finally {
      this.busy = false;
      if (this.activeCommand === abort) this.activeCommand = undefined;
    }
  }
}

export const deckwiseEditorBridge = new DeckWiseEditorBridge();
