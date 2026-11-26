import { LanguageCode, loadLocale, setLanguage } from '@ranuts/shared/i18n';
import { getEditorApi } from './agent-plugin/editor-bridge';
import { deckwiseEditorBridge } from './deckwise-editor-bridge';
import { readNativeTeaching } from './deckwise-native-teaching';
import { routeNativeTeaching } from './deckwise-native-teaching-router';
import { mountDeckwiseTeachingSidebar } from './deckwise-teaching-sidebar';
import { connectNativeTeachingSave } from './docs-native-teaching-save';
import { onlineTeachingTts } from './docs-native-teaching-tts';
import { installSlideSourceIdentity } from './onlyoffice/guards/slide-source-identity';
import { setReadonlyState } from './onlyoffice/readonly';

const writes = new Set([
  'write',
  'lock',
  'confirm',
  'assessment-write',
  'assessment-insert',
  'repair-background-region',
  'media:adopt',
  'adopt-audio',
]);

// This entry runs in the online presentation SDK before binary document import.
const api = getEditorApi();
if (api?.asc_registerCallback && api.asc_unregisterCallback && installSlideSourceIdentity(window)) {
  const readyEvent = 'asc_onDocumentContentReady';
  let sidebar: ReturnType<typeof mountDeckwiseTeachingSidebar> | undefined;
  let closed = false;
  async function mount() {
    api?.asc_unregisterCallback?.(readyEvent, ready);
    const teaching = readNativeTeaching();
    if (!teaching) return;
    const lang = document.documentElement.lang.toLowerCase() || navigator.language.toLowerCase();
    const language = lang.startsWith('zh') ? LanguageCode.ZH : LanguageCode.EN;
    const [save] = await Promise.all([
      connectNativeTeachingSave(api, Reflect.get(window, 'AscCommon')),
      loadLocale(language),
    ]);
    if (closed || api !== getEditorApi()) return;
    // Online authoring requires server acknowledgement of an exclusive teaching commit.
    setReadonlyState(!save);
    setLanguage(language);
    const binding = {
      documentId: teaching.documentId,
      projectId: teaching.documentId,
      sessionId: crypto.randomUUID(),
    };
    const tts = save && onlineTeachingTts(api, save);
    deckwiseEditorBridge.reset(binding);
    sidebar = mountDeckwiseTeachingSidebar(
      binding,
      (type, payload, id) => {
        if (type !== 'document:teaching-request' || !id) return;
        const { operation, args, context } = payload;
        void (async () => {
          if (tts && operation.startsWith('tts:')) return tts(operation, args, context.page ?? 1);
          if (['tts:models', 'tts:list', 'media:list'].includes(operation)) return [];
          throw new Error('在线课件尚未接通此教学操作');
        })().then(
          (result) => sidebar?.receive(id, { result }),
          (error) => sidebar?.receive(id, { error: error instanceof Error ? error.message : String(error) }),
        );
      },
      false,
      (operation, args, page) => {
        if (operation === 'undo' || operation === 'redo')
          return Promise.reject(new Error('在线编辑请直接修改教学内容。'));
        if (!writes.has(operation)) return routeNativeTeaching(operation, args, page);
        if (!save) return Promise.reject(new Error('教学协作服务不可用，当前为只读。'));
        return save(() => routeNativeTeaching(operation, args, page, true));
      },
    );
    sidebar.setOpen(false);
  }
  function ready() {
    void mount().catch((error: unknown) => console.error('WiseOffice could not open online teaching', error));
  }
  api.asc_registerCallback(readyEvent, ready);
  window.addEventListener(
    'pagehide',
    () => {
      closed = true;
      api.asc_unregisterCallback?.(readyEvent, ready);
      sidebar?.dispose();
    },
    { once: true },
  );
}
