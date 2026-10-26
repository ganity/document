import { installDocsNativeBarrier } from './docs-native-barrier-extension';

/** Adapter for exactly the official 9.3.0-138 SDK bytes accepted by the overlay builder. */
export function installDocsNativeBarrier930138(root: any): void {
  const prototype = root.Asc?.asc_docs_api?.prototype;
  if (!prototype || root.isOffline === true) return;
  const noOpSaves = new WeakSet<object>();
  installDocsNativeBarrier<any>(prototype, {
    supported: (editor) =>
      root.isOffline !== true &&
      root.AscCommon?.Qc?.AH === true &&
      editor.Uo === false &&
      typeof editor.asc_Save === 'function' &&
      editor.Pm?.WQ === true &&
      typeof editor.Pm?.wn?.LSh === 'function' &&
      root.g_asc_plugins?.Rd === editor &&
      typeof root.g_asc_plugins.GRf === 'function' &&
      typeof root.g_asc_plugins.HRf === 'function',
    connected: (editor) => [2, 10, 11].includes(editor.Pm?.kTb()),
    restrictions: (editor) => editor.Nra,
    addView: (editor) => editor.asc_addRestriction(128),
    removeView: (editor) => editor.asc_removeRestriction(128),
    idle: (editor) => editor.Ela === true && editor.b9d() && editor.wee() && editor.Pm.kTb() === 2,
    clean: (editor) =>
      editor.Ela === true &&
      !editor.isDocumentModified() &&
      !editor.rRh() &&
      editor.Pm.wn.yEc.length === 0 &&
      !noOpSaves.has(editor),
    save: (editor) => editor.asc_Save(true) === true,
    roundTrip(editor, failed) {
      // Native save lock protects the empty round trip against the automatic save loop.
      noOpSaves.add(editor);
      editor.Ela = false;
      editor.Pm.Vac((event: any) => {
        if (event?.saveLock === false || event?.saveLock === 0) editor.Pm.EId();
        else {
          noOpSaves.delete(editor);
          editor.Ela = true;
          failed();
        }
      });
    },
    observeAck(editor, received) {
      const transport = editor.Pm.wn,
        original = transport.LSh;
      function wrapped(this: any, event: any) {
        const result = original.call(this, event);
        if (noOpSaves.delete(editor)) editor.Ela = true;
        received(transport.JHc);
        return result;
      }
      transport.LSh = wrapped;
      return () => {
        if (transport.LSh === wrapped) transport.LSh = original;
      };
    },
    blockPluginWrites(editor, blocked) {
      const manager = root.g_asc_plugins,
        command = manager.GRf,
        method = manager.HRf;
      const allowed = new Set(['DeckwiseBarrierStatus', 'DeckwiseQuiesceAndDrain', 'DeckwiseResume']);
      function guardCommand(this: any, ...args: any[]) {
        if (blocked()) {
          this.snc(undefined);
          return;
        }
        return command.apply(this, args);
      }
      function guardMethod(this: any, id: string, name: string, args: unknown[]) {
        if (blocked() && !allowed.has(name)) {
          this.snc(undefined);
          return;
        }
        return method.call(this, id, name, args);
      }
      manager.GRf = guardCommand;
      manager.HRf = guardMethod;
      return () => {
        if (manager.GRf === guardCommand) manager.GRf = command;
        if (manager.HRf === guardMethod) manager.HRf = method;
      };
    },
  });
}
