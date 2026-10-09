import { afterEach, describe, expect, it, vi } from 'vitest';

import { injectLocalChromeCss } from '../../lib/onlyoffice/guards/chrome';
import { installServerlessSaveSemantics } from '../../lib/onlyoffice/guards/serverless-save';
import { markDocumentContentReady } from '../../lib/onlyoffice/open-state';
import { requestSaveDocument } from '../../lib/onlyoffice/save-stream';

function editorWindow() {
  const originalSave = vi.fn((_isAutoSave?: boolean, _isUndoRequest?: boolean) => true);
  const originalDownloadAs = vi.fn((_options: unknown) => undefined);
  const editor = {
    asc_Save: originalSave,
    asc_DownloadAs: originalDownloadAs,
    documentFormatSave: 1,
  };
  class DownloadOptions {
    constructor(readonly fileType: number) {}
  }
  const win = {
    Asc: { editor, asc_CDownloadOptions: DownloadOptions },
  } as unknown as Window;
  return { win, editor, originalSave, originalDownloadAs };
}

describe('DeckWise restricted embed mode', () => {
  afterEach(() => {
    window.history.pushState({}, '', '/');
  });

  it('hides download, Save As, export and print without hiding Save', () => {
    window.history.pushState({}, '', '/?embed=1&deckwise=1');
    const doc = document.implementation.createHTMLDocument();

    injectLocalChromeCss(doc);

    const css = doc.getElementById('oo-local-chrome-css')?.textContent ?? '';
    expect(css).toContain('#slot-hbtn-download');
    expect(css).toContain('#fm-btn-save-copy');
    expect(css).toContain('#panel-saveas');
    expect(css).toContain('#fm-btn-print');
    expect(css).not.toContain('toolbar-save');
    expect(css).not.toContain('header-save');
  });

  it('routes Save and Ctrl+S to the DeckWise host and blocks direct exports', () => {
    window.history.pushState({}, '', '/?embed=1&deckwise=1');
    const { win, editor, originalSave, originalDownloadAs } = editorWindow();
    const requested = vi.fn();
    window.addEventListener('deckwise:save-requested', requested, { once: true });

    expect(installServerlessSaveSemantics(win)).toBe(true);
    expect(editor.asc_Save(false, false)).toBe(true);
    editor.asc_DownloadAs({});

    expect(requested).toHaveBeenCalledOnce();
    expect(originalSave).not.toHaveBeenCalled();
    expect(originalDownloadAs).not.toHaveBeenCalled();
  });

  it('preserves the standalone save and export behavior', () => {
    const { win, editor, originalDownloadAs } = editorWindow();

    expect(installServerlessSaveSemantics(win)).toBe(true);
    expect(editor.asc_Save(false, false)).toBe(true);

    expect(originalDownloadAs).toHaveBeenCalledOnce();
  });

  it('allows the host-requested PPTX save through the export guard', async () => {
    window.history.pushState({}, '', '/?embed=1&deckwise=1');
    const iframe = document.createElement('iframe');
    document.body.appendChild(iframe);
    const originalDownloadAs = vi.fn();
    class DownloadOptions {
      constructor(readonly fileType: number) {}
    }
    const child = iframe.contentWindow as Window & {
      Asc?: {
        editor: {
          asc_Save: (isAutoSave?: boolean, isUndoRequest?: boolean) => boolean;
          asc_DownloadAs: (options: unknown) => void;
          documentFormatSave: number;
          isLoadFullApi: boolean;
          isDocumentLoadComplete: boolean;
        };
        asc_CDownloadOptions: typeof DownloadOptions;
      };
    };
    child.Asc = {
      editor: {
        asc_Save: () => true,
        asc_DownloadAs: originalDownloadAs,
        documentFormatSave: 1,
        isLoadFullApi: true,
        isDocumentLoadComplete: true,
      },
      asc_CDownloadOptions: DownloadOptions,
    };
    installServerlessSaveSemantics(child);
    (window as unknown as { editor?: { downloadAs: (format: string) => void } }).editor = {
      downloadAs: vi.fn(),
    };
    markDocumentContentReady();

    const saving = requestSaveDocument('PPTX');
    await vi.waitFor(() => expect(originalDownloadAs).toHaveBeenCalledOnce());
    window.dispatchEvent(
      new MessageEvent('message', {
        data: { type: 'onlyoffice-file-stream', fileType: 'pptx', buffer: new ArrayBuffer(8) },
      }),
    );

    await expect(saving).resolves.toMatchObject({ name: 'document.pptx' });
    iframe.remove();
    delete (window as unknown as { editor?: unknown }).editor;
  });
});
