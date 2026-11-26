import { synchronizeNativeTeachingStructure } from './deckwise-native-teaching';

type SaveMessage = {
  type: string;
  deckwiseTeachingSaveId?: string;
  saveLock?: boolean;
  syncChangesIndex?: number;
};

type Transport = {
  syncChangesIndex: number;
  websocketMaxPayloadSize: number;
  _send(message: SaveMessage, encryption?: boolean): void;
  _onUnSaveLock(message: SaveMessage): void;
};

type DocsEditor = {
  canSave: boolean;
  canEdit(): boolean;
  canSendChanges(): boolean;
  isLongAction(): boolean;
  IsUserSave: boolean;
  _onSaveCallback(message: SaveMessage): void;
  CoAuthoringApi: {
    _CoAuthoringApi: Transport;
    get_state(): number;
    askSaveChanges(callback: (result: SaveMessage) => void): void;
    unSaveLock(): void;
    callPRC(
      data: { type: string },
      timeout: number,
      callback: (timeout: boolean, data?: { available: boolean }) => void,
    ): void;
  };
};

type Collaboration = {
  Get_GlobalLock(): boolean;
  Have_OtherChanges(): boolean;
  Apply_Changes(): void;
  Send_Changes(userSave: boolean, info: unknown, updateInterface?: boolean, afterSaveLock?: boolean): void;
};

export async function connectNativeTeachingSave(api: unknown, common: unknown) {
  const editor = api as Partial<DocsEditor> | undefined;
  const sdk = common as
    | {
        History?: { Have_Changes(): boolean };
        CollaborativeEditing?: Collaboration;
      }
    | undefined;
  const history = sdk?.History;
  const collaboration = sdk?.CollaborativeEditing;
  const office = editor?.CoAuthoringApi;
  const transport = office?._CoAuthoringApi;
  if (
    !editor ||
    !history ||
    !collaboration ||
    !office ||
    !transport ||
    typeof collaboration.Send_Changes !== 'function' ||
    typeof editor._onSaveCallback !== 'function' ||
    typeof office.callPRC !== 'function' ||
    typeof transport._send !== 'function' ||
    typeof transport._onUnSaveLock !== 'function'
  )
    return;
  if (
    !(await new Promise<boolean>((resolve) => {
      office.callPRC({ type: 'deckwiseTeachingSaveCapability' }, 3000, (timeout, result) =>
        resolve(!timeout && result?.available === true),
      );
    }))
  )
    return;
  const sendChanges = collaboration.Send_Changes;
  collaboration.Send_Changes = function (...args) {
    // The SDK has applied remote changes and holds the save lock at this point.
    // Append the teaching page order to the pending native history before it is sent.
    if (args[3] && history.Have_Changes()) {
      try {
        synchronizeNativeTeachingStructure();
      } catch (error) {
        office.unSaveLock();
        throw error;
      }
    }
    return sendChanges.apply(this, args);
  };
  let busy = false;
  let uncertain = false;
  return async function save<T>(write: () => Promise<T>): Promise<T> {
    if (uncertain) throw new Error('上次教学保存结果待确认，请重新打开在线课件。');
    if (
      busy ||
      !editor.canSave ||
      !editor.canEdit?.() ||
      !editor.canSendChanges?.() ||
      editor.isLongAction?.() ||
      collaboration.Get_GlobalLock() ||
      office.get_state() !== 2 ||
      history.Have_Changes()
    )
      throw new Error('编辑器正在处理修改，请稍后重试。');
    busy = true;
    editor.canSave = false;
    const id = crypto.randomUUID();
    const index = transport.syncChangesIndex;
    const send = transport._send;
    const acknowledge = transport._onUnSaveLock;
    let writing = false;
    let sent = false;
    let granted = false;
    let resolveSaved: (() => void) | undefined;
    const saved = new Promise<void>((resolve) => {
      resolveSaved = resolve;
    });

    transport._send = function (message, encryption) {
      if (message.type === 'isSaveLock' || (writing && message.type === 'saveChanges')) {
        if (message.type === 'saveChanges') sent = true;
        send.call(this, { ...message, deckwiseTeachingSaveId: id }, encryption);
      } else send.call(this, message, encryption);
    };
    transport._onUnSaveLock = function (message) {
      acknowledge.call(this, message);
      if (sent && message.deckwiseTeachingSaveId === id && Number(message.syncChangesIndex) > index) resolveSaved?.();
    };

    let timer: ReturnType<typeof setTimeout> | undefined;
    const payloadSize = transport.websocketMaxPayloadSize;
    try {
      const response = await new Promise<SaveMessage>((resolve) => office.askSaveChanges(resolve));
      granted = response.saveLock === false;
      if (!granted || response.deckwiseTeachingSaveId !== id || response.syncChangesIndex !== index)
        throw new Error('课件正在由其他成员保存或同步，请稍后重试。');
      if (collaboration.Have_OtherChanges()) collaboration.Apply_Changes();
      if (
        collaboration.Have_OtherChanges() ||
        collaboration.Get_GlobalLock() ||
        transport.syncChangesIndex !== index ||
        history.Have_Changes() ||
        !editor.canEdit?.() ||
        editor.isLongAction?.()
      )
        throw new Error('课件已更新，请核对当前讲解后再保存。');

      const result = await write();
      if (!history.Have_Changes()) return result;
      writing = true;
      editor.IsUserSave = false;
      // A teaching edit and its media travel in the same native change batch.
      transport.websocketMaxPayloadSize = Math.max(payloadSize, 90_000_000);
      try {
        editor._onSaveCallback!(response);
      } finally {
        transport.websocketMaxPayloadSize = payloadSize;
      }
      await Promise.race([
        saved,
        new Promise<never>((_, reject) => {
          timer = setTimeout(() => reject(new Error('教学保存结果待确认，请重新打开在线课件。')), 15_000);
        }),
      ]);
      return result;
    } catch (error) {
      if (writing || (granted && history.Have_Changes())) uncertain = true;
      throw error;
    } finally {
      clearTimeout(timer);
      busy = false;
      if (!uncertain) {
        transport._send = send;
        transport._onUnSaveLock = acknowledge;
        if (!sent && granted) office.unSaveLock();
        if (!sent) editor.canSave = true;
      }
    }
  };
}
