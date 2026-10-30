import { mountPanelContext } from './agent-plugin/ui/panel-context';
import { mountDeckwiseTeachingSidebar } from './deckwise-teaching-sidebar';
import { clearActiveSource } from './active-document-source';
import { clearDocumentSession } from './history/autosave';
import { beginDocumentOpen, commitDocumentOpen, isMountedDocumentOpen } from './document-open';
import { getDocmentObj, setDocmentObj } from '@ranuts/shared/store';
import { getDocumentType } from '@ranuts/shared/document-utils';
import { handleDocumentOperation, loadEditorApi } from './converter';
import { openDocumentFromUrl } from './document';
import { getReadonlyMode, requestSaveDocument, setReadonlyMode } from './onlyoffice-editor';
import { hasPendingEmbeddedSaveRequest } from './onlyoffice/save-stream';
import { captureSavePoint, commitSavePoint, invalidateSavedPoint, isSavePointCurrent } from './onlyoffice/save-point';
import { hasUnsavedChanges, markDocumentDirty, markDocumentSaved, onUnsavedChangesChange, getLastEditAt, onDocumentEdit } from './unsaved-guard';
import { DECKWISE_SAVE_REQUEST_EVENT, isDeckWiseMode } from './embed-mode';
import { deckwiseEditorBridge } from './deckwise-editor-bridge';
import { deckwisePageLineage } from './deckwise-page-lineage';
import { renderDeckWiseCurrentSlidePreview } from './deckwise-slide-preview';

type EmbedMessageType =
  | 'document:open'
  | 'document:open-url'
  | 'document:open-file'
  | 'document:open-buffer'
  | 'document:set-readonly'
  | 'document:save'
  | 'document:commit-save'
  | 'document:get-context'
  | 'document:execute'
  | 'document:review-image'
  | 'document:review-image-target'
  | 'document:execute-image'
  | 'document:review-fade'
  | 'document:execute-fade'
  | 'document:render-page-preview'
  | 'document:get-state'
  | 'document:teaching-enable'
  | 'document:teaching-response';

type EmbedMessage = {
  id?: string;
  type?: EmbedMessageType;
  payload?: Record<string, any>;
};

type EmbedResponsePayload = Record<string, any>;

const EMBED_QUERY_KEYS = ['embed', 'embedded'];

let teachingSidebar: ReturnType<typeof mountDeckwiseTeachingSidebar> | undefined;
let initialized = false;
let parentOrigin = '*';
let acceptedParentOrigin: string | null = null;
let isEmbedMode = false;
let pendingHostSave: { id: string; point: ReturnType<typeof captureSavePoint>; exported: boolean; lineage?: Parameters<typeof deckwisePageLineage.commitSave>[0] } | null = null;
let saveOutcomeUncertain = false;

function getQueryValue(key: string): string | null {
  return new URLSearchParams(window.location.search).get(key);
}

function detectEmbedMode(): boolean {
  if (window.parent !== window) {
    return true;
  }

  return EMBED_QUERY_KEYS.some((key) => {
    const value = getQueryValue(key);
    return value === '' || value === '1' || value === 'true';
  });
}

function normalizeTargetOrigin(origin: string): string {
  return origin && origin !== 'null' ? origin : '*';
}

function shouldAcceptMessage(event: MessageEvent): boolean {
  if (event.source !== window.parent) return false;
  const allowedOrigin = getQueryValue('embedOrigin');
  if (allowedOrigin && event.origin !== allowedOrigin) return false;
  return acceptedParentOrigin === null || event.origin === acceptedParentOrigin;
}

function postToParent(type: string, payload: EmbedResponsePayload = {}, id?: string): void {
  if (!isEmbedMode) {
    return;
  }

  window.parent.postMessage(
    {
      id,
      type,
      payload,
    },
    parentOrigin,
  );
}

// Some hosts (e.g. a Qt WebEngine app driving the page via
// QWebEngineView.page().runJavaScript()) can only inject JSON-serializable
// literals -- there is no way to hand over a real ArrayBuffer from the host
// language. Base64-encoding the file content is the standard workaround, so
// accept a base64 string alongside the binary forms below. Strips an optional
// data-URL prefix ("data:...;base64,") for convenience.
function decodeBase64ToUint8Array(base64: string): Uint8Array {
  const commaIndex = base64.indexOf(',');
  const raw = base64.startsWith('data:') && commaIndex !== -1 ? base64.slice(commaIndex + 1) : base64;
  const binaryString = atob(raw);
  const bytes = new Uint8Array(binaryString.length);
  for (let i = 0; i < binaryString.length; i++) {
    bytes[i] = binaryString.charCodeAt(i);
  }
  return bytes;
}

function makeFileFromPayload(payload: Record<string, any>): File {
  const fileName = payload.fileName || payload.name || 'document.xlsx';

  if (payload.file instanceof File) {
    return payload.file;
  }

  if (payload.blob instanceof Blob) {
    return new File([payload.blob], fileName, {
      type: payload.blob.type || payload.mimeType || 'application/octet-stream',
    });
  }

  const buffer = payload.buffer || payload.arrayBuffer || payload.bytes || payload.data || payload.base64;
  if (buffer instanceof ArrayBuffer) {
    return new File([buffer], fileName, {
      type: payload.mimeType || 'application/octet-stream',
    });
  }

  if (buffer instanceof Uint8Array) {
    const arrayBuffer = buffer.buffer.slice(buffer.byteOffset, buffer.byteOffset + buffer.byteLength) as ArrayBuffer;
    return new File([arrayBuffer], fileName, {
      type: payload.mimeType || 'application/octet-stream',
    });
  }

  if (typeof buffer === 'string') {
    const bytes = decodeBase64ToUint8Array(buffer);
    return new File([bytes.buffer as ArrayBuffer], fileName, {
      type: payload.mimeType || 'application/octet-stream',
    });
  }

  throw new Error('document:open requires url, file, blob, buffer, arrayBuffer, bytes, data, or base64');
}

async function openFile(file: File, readonly = false, lineageBinding?: unknown): Promise<void> {
  const opening = beginDocumentOpen();
  await loadEditorApi();
  opening.assertCurrent();
  const onContentReady =
    lineageBinding !== undefined ? await deckwisePageLineage.begin(lineageBinding, file) : undefined;
  opening.assertCurrent();
  setDocmentObj({
    fileName: file.name,
    file,
  });
  await handleDocumentOperation({
    fileName: file.name,
    file,
    isNew: false,
    readonly,
    isCurrent: opening.isCurrent,
    ...(onContentReady ? { onContentReady } : {}),
  });
  if (!commitDocumentOpen(opening.id)) throw new DOMException('Document open superseded', 'AbortError');
  clearDocumentSession();
  await clearActiveSource();
  if (!isMountedDocumentOpen(opening.id)) throw new DOMException('Document open superseded', 'AbortError');
}

async function handleOpen(payload: Record<string, any>): Promise<void> {
  const readonly = Boolean(payload.readonly);

  if (payload.url) {
    await openDocumentFromUrl(String(payload.url), payload.fileName, {
      readonly,
      fetchOptions: payload.fetchOptions,
    });
    return;
  }

  const file = makeFileFromPayload(payload);
  if (isDeckWiseMode() && !getDocumentType(file.name.split('.').pop() || '')) {
    throw new Error('Unsupported document format');
  }
  const lineageBinding =
    isDeckWiseMode() && payload.trackPageLineage === true && /\.pptx$/i.test(file.name)
      ? payload.editorBinding
      : undefined;
  await openFile(file, readonly, lineageBinding);
}

async function handleMessage(event: MessageEvent): Promise<void> {
  const message = event.data as EmbedMessage;
  if (
    !message ||
    typeof message !== 'object' ||
    typeof message.type !== 'string' ||
    !message.type.startsWith('document:')
  ) {
    return;
  }

  if (!shouldAcceptMessage(event)) {
    return;
  }

  acceptedParentOrigin ??= event.origin;
  parentOrigin = normalizeTargetOrigin(acceptedParentOrigin);
  const payload = message.payload || {};

  try {
    if (['document:save', 'document:commit-save', 'document:set-readonly'].includes(String(message.type)))
      deckwiseEditorBridge.revokeFade();
    switch (message.type) {
      case 'document:teaching-enable':
        if (!isDeckWiseMode()) throw new Error('Teaching requires DeckWise mode');
        deckwiseEditorBridge.getContext(payload.binding);
        teachingSidebar?.dispose();
        teachingSidebar = mountDeckwiseTeachingSidebar(payload.binding, postToParent);
        break;
      case 'document:teaching-response':
        teachingSidebar?.receive(message.id ?? '', payload);
        break;
      case 'document:open':
      case 'document:open-url':
      case 'document:open-file':
      case 'document:open-buffer':
        if (isDeckWiseMode() && (message.type !== 'document:open-buffer' || payload.url)) {
          throw new Error('DeckWise only accepts document:open-buffer');
        }
        if (pendingHostSave || (isDeckWiseMode() && hasPendingEmbeddedSaveRequest()))
          throw new Error('Finish the current save before opening another document');
        if (isDeckWiseMode()) {
          teachingSidebar?.dispose();
          teachingSidebar = undefined;
          deckwiseEditorBridge.reset();
          deckwisePageLineage.reset();
        }
        await handleOpen(payload);
        saveOutcomeUncertain = false;
        if (isDeckWiseMode() && payload.editorBinding !== undefined) {
          deckwiseEditorBridge.reset(payload.editorBinding);
        }
        postToParent('document:opened', { readonly: getReadonlyMode() }, message.id);
        break;

      case 'document:get-context':
      case 'document:execute':
      case 'document:review-image':
      case 'document:review-image-target':
      case 'document:execute-image':
      case 'document:review-fade':
      case 'document:execute-fade': {
        if (!isDeckWiseMode()) throw new Error('Editor commands require DeckWise mode');
        if (hasPendingEmbeddedSaveRequest() || pendingHostSave || saveOutcomeUncertain)
          throw new Error('Finish or recover the current save before using editor tools');
        if (typeof message.id !== 'string' || !/^[\w-]{1,128}$/.test(message.id))
          throw new Error('Invalid editor request ID');
        if (message.type === 'document:review-fade') {
          if (Object.keys(payload).some((key) => !['binding', 'authorization'].includes(key)))
            throw new Error('Invalid fade review payload');
          if (payload.authorization !== null && hasUnsavedChanges()) {
            deckwiseEditorBridge.revokeFade();
            throw new Error('Save the latest PPTX before fade inspection');
          }
          postToParent(
            'document:context',
            deckwiseEditorBridge.reviewFade(payload.binding, payload.authorization),
            message.id,
          );
        } else if (message.type === 'document:execute-fade') {
          if (Object.keys(payload).some((key) => !['command', 'authorization'].includes(key)))
            throw new Error('Invalid fade execution payload');
          if (hasUnsavedChanges()) {
            deckwiseEditorBridge.revokeFade();
            throw new Error('Save the latest PPTX before fade inspection');
          }
          postToParent(
            'document:command-result',
            await deckwiseEditorBridge.executeFade(payload.command, payload.authorization, message.id),
            message.id,
          );
        } else if (message.type === 'document:review-image-target') {
          if (Object.keys(payload).sort().join(',') !== 'asset,binding,target')
            throw new Error('Invalid image target review payload');
          postToParent(
            'document:image-target-reviewed',
            deckwiseEditorBridge.reviewImageTarget(payload.binding, payload.asset, payload.target),
            message.id,
          );
        } else if (message.type === 'document:review-image') {
          if (Object.keys(payload).some((key) => !['binding', 'asset'].includes(key)))
            throw new Error('Invalid image review payload');
          postToParent(
            'document:context',
            deckwiseEditorBridge.reviewImage(payload.binding, payload.asset),
            message.id,
          );
        } else if (message.type === 'document:execute-image') {
          if (Object.keys(payload).some((key) => !['command', 'bytes'].includes(key)))
            throw new Error('Invalid image execution payload');
          postToParent(
            'document:command-result',
            await deckwiseEditorBridge.executeImage(payload.command, payload.bytes, message.id),
            message.id,
          );
        } else if (message.type === 'document:get-context') {
          const context = deckwiseEditorBridge.getContext(payload);
          const pageLineage = deckwisePageLineage.context();
          postToParent('document:context', { ...context, ...(pageLineage ? { pageLineage } : {}) }, message.id);
        } else {
          postToParent('document:command-result', await deckwiseEditorBridge.execute(payload, message.id), message.id);
        }
        break;
      }

      case 'document:render-page-preview': {
        if (!isDeckWiseMode()) throw new Error('Page previews require DeckWise mode');
        if (hasPendingEmbeddedSaveRequest() || pendingHostSave || saveOutcomeUncertain)
          throw new Error('Finish or recover the current save before rendering a page preview');
        if (typeof message.id !== 'string' || !/^[\w-]{1,128}$/.test(message.id))
          throw new Error('Invalid editor request ID');
        if (
          Object.keys(payload).some((key) => !['binding', 'editGeneration', 'omitNativeText', 'page'].includes(key)) ||
          (payload.omitNativeText !== undefined && typeof payload.omitNativeText !== 'boolean') ||
          (payload.page !== undefined &&
            (!Number.isSafeInteger(payload.page) || payload.page < 1 || payload.page > 5000))
        )
          throw new Error('Invalid page preview request');
        postToParent(
          'document:page-preview',
          payload.page === undefined
            ? renderDeckWiseCurrentSlidePreview(payload.binding, payload.editGeneration, payload.omitNativeText)
            : renderDeckWiseCurrentSlidePreview(
                payload.binding,
                payload.editGeneration,
                payload.omitNativeText,
                payload.page,
              ),
          message.id,
        );
        break;
      }

      case 'document:set-readonly':
        setReadonlyMode(Boolean(payload.readonly));
        postToParent('document:readonly-changed', { readonly: getReadonlyMode() }, message.id);
        break;

      case 'document:save': {
        if (isDeckWiseMode() && deckwiseEditorBridge.isExecuting)
          throw new Error('Finish the current editor command before saving');
        // Default to the open document's own format: a bare document:save on
        // a docx/pptx used to ask x2t for XLSX, which fails (code 88) and
        // only surfaced as a save timeout.
        // Legacy binary formats (and HTML-as-xls, opened as XLSX) cannot be
        // written back by x2t; their natural target is the OOXML sibling.
        const currentExt = (getDocmentObj()?.fileName || '').split('.').pop()?.toUpperCase() || '';
        const LEGACY_TARGET: Record<string, string> = { XLS: 'XLSX', DOC: 'DOCX', PPT: 'PPTX' };
        const defaultExt = LEGACY_TARGET[currentExt] || currentExt || 'XLSX';
        const targetExt = isDeckWiseMode() ? defaultExt : payload.targetExt || defaultExt;
        if (isDeckWiseMode() && payload.targetExt && String(payload.targetExt).toUpperCase() !== defaultExt) {
          throw new Error(`DeckWise only permits ${defaultExt} save for this document`);
        }
        const hostManaged = payload.requireHostCommit === true;
        if (hostManaged && (typeof message.id !== 'string' || !/^[\w-]{1,128}$/.test(message.id))) {
          throw new Error('Invalid save request ID');
        }
        if (pendingHostSave) throw new Error('A host save is still pending');
        if (isDeckWiseMode() && hasPendingEmbeddedSaveRequest())
          throw new Error('A save request is already in progress');
        if (saveOutcomeUncertain) throw new Error('Export outcome is uncertain; reopen before saving again');
        const returnOriginalOnTimeout = hostManaged ? false : Boolean(payload.returnOriginalOnTimeout);
        // Taken before the export starts: edits made while it runs are not in
        // the bytes the host receives, so they must keep the document dirty.
        const savePoint = captureSavePoint();
        const captureLineage = hostManaged && isDeckWiseMode() ? deckwisePageLineage.captureSave() : undefined;
        const pending = hostManaged ? { id: message.id!, point: savePoint, exported: false } : null;
        pendingHostSave = pending;
        let file: File;
        let pageLineage: Awaited<ReturnType<NonNullable<typeof captureLineage>>> | undefined;
        try {
          file = await requestSaveDocument(targetExt, { returnOriginalOnTimeout });
          pageLineage = captureLineage
            ? await captureLineage(file, message.id!, () => isSavePointCurrent(savePoint))
            : undefined;
        } catch (error) {
          pendingHostSave = null;
          // An untagged SDK stream may still arrive after an export timeout.
          if (hostManaged) saveOutcomeUncertain = true;
          throw error;
        }
        if (pending) { pending.exported = true; pendingHostSave!.lineage = pageLineage; }
        // The host now holds the edited bytes. A timeout fallback may have
        // handed it the original file instead, which saves nothing -- the
        // request cannot tell the two apart, so that mode never clears the flag.
        if (!hostManaged && !returnOriginalOnTimeout && commitSavePoint(savePoint)) {
          markDocumentSaved();
        }
        postToParent(
          'document:saved',
          {
            file,
            fileName: file.name,
            mimeType: file.type,
            size: file.size,
            dirty: hasUnsavedChanges(),
            ...(pageLineage ? { pageLineage } : {}),
          },
          message.id,
        );
        break;
      }

      case 'document:commit-save': {
        if (!pendingHostSave || pendingHostSave.id !== message.id || !pendingHostSave.exported) break;
        if (typeof payload.success !== 'boolean') throw new Error('Invalid save acknowledgement');
        const pending = pendingHostSave;
        pendingHostSave = null;
        if (payload.success) {
          if (pending.lineage) deckwisePageLineage.commitSave(pending.lineage);
          if (commitSavePoint(pending.point)) markDocumentSaved();
          else {
            invalidateSavedPoint(pending.point);
            markDocumentDirty();
          }
        }
        postToParent('document:save-committed', { dirty: hasUnsavedChanges(), success: payload.success }, message.id);
        if (payload.success && !hasUnsavedChanges()) teachingSidebar?.refresh();
        break;
      }

      case 'document:get-state':
        postToParent(
          'document:state',
          { readonly: getReadonlyMode(), hasDocument: Boolean(window.editor), dirty: hasUnsavedChanges() },
          message.id,
        );
        break;

      default:
        break;
    }
  } catch (error) {
    postToParent(
      'document:error',
      {
        message: error instanceof Error ? error.message : String(error),
      },
      message.id,
    );
  }
}

export function initEmbedApi(): void {
  if (initialized) {
    return;
  }

  initialized = true;
  isEmbedMode = detectEmbedMode();
  acceptedParentOrigin = getQueryValue('embedOrigin');
  if (acceptedParentOrigin) parentOrigin = normalizeTargetOrigin(acceptedParentOrigin);

  if (!isEmbedMode) {
    return;
  }

  document.body.classList.add('embed-mode');
  // Pushed, not polled: the host needs it to decide when to save (a tab
  // switch, an idle timer) without exporting documents nobody touched.
  let autosaveTimer: ReturnType<typeof setTimeout> | undefined;
  let attemptedEdit = -1;
  const autosave = () => {
    autosaveTimer = undefined;
    if (!isDeckWiseMode() || !hasUnsavedChanges() || getReadonlyMode()) return;
    if (pendingHostSave || hasPendingEmbeddedSaveRequest() || Date.now() - getLastEditAt() < 1500) {
      autosaveTimer = setTimeout(autosave, 500);
      return;
    }
    if (!saveOutcomeUncertain && attemptedEdit !== getLastEditAt()) {
      attemptedEdit = getLastEditAt();
      postToParent('document:save-requested');
    }
  };
  if (isDeckWiseMode()) {
    let observedPoint = captureSavePoint();
    mountPanelContext({
      isActive: () => isDeckWiseMode(),
      refresh: () => {
        if (observedPoint && isSavePointCurrent(observedPoint)) return;
        observedPoint = captureSavePoint();
        // The SDK modified callback reports only clean/dirty transitions; typing can keep extending one undo point.
        if (observedPoint && hasUnsavedChanges()) markDocumentDirty();
      },
    });
  }
  onDocumentEdit(() => {
    clearTimeout(autosaveTimer);
    if (isDeckWiseMode()) autosaveTimer = setTimeout(autosave, 1500);
  });
  onUnsavedChangesChange((dirty) => {
    postToParent('document:dirty-changed', { dirty });
    clearTimeout(autosaveTimer);
    if (dirty && isDeckWiseMode()) autosaveTimer = setTimeout(autosave, 1500);
  });
  if (isDeckWiseMode()) {
    document.body.classList.add('deckwise-mode');
    window.addEventListener(DECKWISE_SAVE_REQUEST_EVENT, () => {
      postToParent('document:save-requested');
    });
  }
  window.addEventListener('message', (event) => {
    void handleMessage(event);
  });

  window.addEventListener('load', () => {
    postToParent('document:ready');
  });
}
