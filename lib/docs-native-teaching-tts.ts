import { adoptNativeTeachingAudio, readNativeTeaching } from './deckwise-native-teaching';
import { routeNativeTeaching } from './deckwise-native-teaching-router';

type RelayReply = { code: number; message?: string; data?: unknown };
type Rpc = {
  callPRC(
    request: { type: string; command: Record<string, unknown> },
    timeout: number,
    callback: (timedOut: boolean, reply?: RelayReply) => void,
  ): void;
};
type Quote = { priceVersion: string; quoteToken: string; maxPoints: number; maxChars: number; expiresAt: string };
type Voice = { mime: string; bytesBase64: string };
type Task = {
  id: string;
  pageId: string;
  actionId: string;
  text: string;
  model: string;
  voice: string | null;
  quote?: Quote & { quoteId: string };
  state: 'prepared' | 'processing' | 'ready' | 'failed' | 'submission_unknown';
  errorCode?: string;
  audio?: { bytes: ArrayBuffer; sha256: string; durationMs: number };
};

export function onlineTeachingTts(api: unknown, save: <T>(write: () => Promise<T>) => Promise<T>) {
  const rpc = (api as { CoAuthoringApi: Rpc }).CoAuthoringApi;
  const tasks = new Map<string, Task>();
  async function relay<T>(command: Record<string, unknown>): Promise<T> {
    const result = await new Promise<RelayReply | undefined>((resolve) => {
      rpc.callPRC({ type: 'deckwiseTeachingTts', command }, 190_000, (timeout, reply) =>
        resolve(timeout ? undefined : reply),
      );
    });
    if (!result) throw new Error('UPSTREAM_OUTCOME_UNKNOWN');
    if (result.code !== 200) throw new Error(result.message || 'ONLINE_TTS_UNAVAILABLE');
    return result.data as T;
  }
  const describe = (task: Task, current: boolean) => ({
    intentId: task.id,
    actionId: task.actionId,
    model: task.model,
    voice: task.voice,
    current,
    phase:
      task.state === 'submission_unknown' ? 'submission-unknown' : task.state === 'prepared' ? 'prepared' : 'submitted',
    task: task.state === 'prepared' ? null : { state: task.state, errorCode: task.errorCode },
  });
  async function generate(task: Task) {
    let received = false;
    try {
      const { priceVersion, quoteToken, maxPoints } = task.quote!;
      const result = await relay<Voice>({
        operation: 'submit',
        intentId: task.id,
        model: task.model,
        text: task.text,
        voice: task.voice,
        confirmation: { priceVersion, quoteToken, maxPoints },
      });
      received = true;
      if (
        result.mime !== 'audio/mpeg' ||
        typeof result.bytesBase64 !== 'string' ||
        result.bytesBase64.length > 22_369_624
      )
        throw new Error('INVALID_TTS_AUDIO');
      const bytes = Uint8Array.from(atob(result.bytesBase64), (c) => c.charCodeAt(0)).buffer;
      const audio = await new OfflineAudioContext(1, 1, 44_100).decodeAudioData(bytes.slice(0));
      const hash = new Uint8Array(await crypto.subtle.digest('SHA-256', bytes));
      const sha256 = Array.from(hash, (byte) => byte.toString(16).padStart(2, '0')).join('');
      task.audio = { bytes, sha256, durationMs: Math.max(1, Math.round(audio.duration * 1000)) };
      task.state = 'ready';
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      task.errorCode = message;
      task.state =
        received || ['UPSTREAM_OUTCOME_UNKNOWN', 'REQUEST_ALREADY_FORWARDED'].includes(message)
          ? 'submission_unknown'
          : 'failed';
    }
  }
  return async function request(operation: string, args: Record<string, unknown>, pageNumber: number) {
    const teaching = readNativeTeaching();
    const page = teaching?.pages[pageNumber - 1];
    if (
      !page ||
      !teaching ||
      args.viewId !== teaching.documentId + '.' + page.pageId ||
      args.expectedTeachingRevision !== teaching.revision ||
      page.locked
    )
      throw new Error('课件已更新，请核对当前讲解。');
    const current = (task: Task) =>
      task.pageId === page.pageId &&
      page.actions.some(
        (action) => action.type === 'speech' && action.id === task.actionId && action.text === task.text,
      );
    if (operation === 'tts:models') return relay({ operation: 'models' });
    if (operation === 'tts:list')
      return [...tasks.values()]
        .filter((task) => task.pageId === page.pageId)
        .map((task) => describe(task, current(task)));
    if (operation === 'tts:prepare') {
      const action = page.actions.find((item) => item.type === 'speech' && item.id === args.actionId);
      if (
        !action ||
        typeof action.text !== 'string' ||
        !action.text.trim() ||
        typeof args.model !== 'string' ||
        (args.voice !== null && typeof args.voice !== 'string')
      )
        throw new Error('请选择配音模型和音色。');
      const recent = [...tasks.values()].find(
        (task) =>
          task.actionId === action.id &&
          task.model === args.model &&
          task.voice === args.voice &&
          current(task) &&
          task.state !== 'failed',
      );
      if (recent) return { intentId: recent.id };
      const id = crypto.randomUUID();
      tasks.set(id, {
        id,
        pageId: page.pageId,
        actionId: action.id,
        text: action.text,
        model: args.model,
        voice: args.voice,
        state: 'prepared',
      });
      return { intentId: id };
    }
    const task = tasks.get(String(args.intentId));
    if (!task || !current(task)) throw new Error('讲解已修改，配音任务与当前文字不匹配。');
    if (operation === 'tts:status') return describe(task, true);
    if (operation === 'tts:quote') {
      if (task.state !== 'prepared') throw new Error('这段配音已有生成任务，请查看任务状态。');
      const quote = await relay<Quote>({ operation: 'quote', model: task.model });
      if (task.text.length > quote.maxChars) throw new Error('讲解字数超过当前配音模型的上限。');
      task.quote = { ...quote, quoteId: crypto.randomUUID() };
      return { quoteId: task.quote.quoteId, maxPoints: quote.maxPoints };
    }
    if (operation === 'tts:submit') {
      if (task.state !== 'prepared') return describe(task, true);
      if (!task.quote || args.quoteId !== task.quote.quoteId || Date.parse(task.quote.expiresAt) <= Date.now())
        throw new Error('配音报价已失效，请重新报价。');
      task.state = 'processing';
      void generate(task);
      return describe(task, true);
    }
    if (operation === 'tts:archive' && task.audio)
      return { intentId: task.id, archived: true, durationMs: task.audio.durationMs };
    if (operation === 'tts:adopt' && task.audio) {
      const audio = task.audio;
      const result = await save(async () => {
        await adoptNativeTeachingAudio({
          ...audio,
          documentId: teaching.documentId,
          expectedRevision: teaching.revision,
          pageId: page.pageId,
          actionId: task.actionId,
          assetId: 'tts-' + audio.sha256,
          mimeType: 'audio/mpeg',
        });
        return routeNativeTeaching('open', {}, pageNumber);
      });
      tasks.delete(task.id);
      return result?.value;
    }
    throw new Error('配音任务还没有可用结果。');
  };
}
