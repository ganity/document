import type { CommunityPluginApi } from './docs-community-plugin';
import type { NativeBarrierReceipt } from './docs-native-barrier-extension';

export type DocsNativeBarrierReceipt = NativeBarrierReceipt;
export interface DocsNativeBarrierCapability {
  available: boolean;
}
export const isBarrierId = (value: unknown): value is string =>
  typeof value === 'string' && /^[a-f0-9]{8}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{12}$/.test(value);
const object = (value: unknown): value is Record<string, unknown> =>
  !!value && typeof value === 'object' && !Array.isArray(value);
const nativeErrors = new Set([
  'NATIVE_BARRIER_INVALID',
  'NATIVE_BARRIER_BUSY',
  'NATIVE_BARRIER_UNAVAILABLE',
  'NATIVE_BARRIER_NOT_FOUND',
  'NATIVE_BARRIER_ACK_UNKNOWN',
  'NATIVE_BARRIER_DISCONNECTED',
  'NATIVE_BARRIER_SAVE_UNAVAILABLE',
  'NATIVE_BARRIER_SAVE_UNKNOWN',
  'NATIVE_BARRIER_RESTRICTION_LOST',
  'NATIVE_BARRIER_STATE_CHANGED',
  'NATIVE_BARRIER_SETUP_FAILED',
]);
export const isNativeBarrierError = (value: unknown): value is string =>
  typeof value === 'string' && nativeErrors.has(value);
export function isBarrierCapability(value: unknown): value is DocsNativeBarrierCapability {
  return object(value) && Object.keys(value).length === 1 && typeof value.available === 'boolean';
}
export function isBarrierReceipt(value: unknown, barrierId: string): value is DocsNativeBarrierReceipt {
  if (
    !object(value) ||
    Object.keys(value).some((key) => !['barrierId', 'state', 'serverIndex', 'error'].includes(key)) ||
    value.barrierId !== barrierId
  )
    return false;
  if (value.serverIndex !== undefined && (!Number.isSafeInteger(value.serverIndex) || Number(value.serverIndex) < 0))
    return false;
  if (value.state === 'draining') return value.serverIndex === undefined && value.error === undefined;
  if (value.state === 'held' || value.state === 'resumed')
    return Number.isSafeInteger(value.serverIndex) && Number(value.serverIndex) >= 0 && value.error === undefined;
  return value.state === 'uncertain-held' && isNativeBarrierError(value.error);
}
export type DocsBarrierTool =
  'native_barrier_capability' | 'native_barrier_quiesce' | 'native_barrier_status' | 'native_barrier_resume';
export function isBarrierTool(value: unknown): value is DocsBarrierTool {
  return (
    typeof value === 'string' &&
    ['native_barrier_capability', 'native_barrier_quiesce', 'native_barrier_status', 'native_barrier_resume'].includes(
      value,
    )
  );
}

/** Only fixed native methods cross the owned plugin port. No command, source or editor object is accepted. */
export function invokeDocsBarrier(
  plugin: CommunityPluginApi,
  tool: DocsBarrierTool,
  barrierId: string | undefined,
  reply: (result?: unknown, error?: string) => void,
): void {
  const capability = tool === 'native_barrier_capability';
  const id = capability ? crypto.randomUUID() : barrierId;
  if (!isBarrierId(id)) {
    reply(undefined, 'DOCS_BARRIER_INVALID_REQUEST');
    return;
  }
  const method =
    tool === 'native_barrier_quiesce'
      ? 'DeckwiseQuiesceAndDrain'
      : tool === 'native_barrier_resume'
        ? 'DeckwiseResume'
        : 'DeckwiseBarrierStatus';
  try {
    plugin.executeMethod(method, [id], (value) => {
      if (value === undefined) {
        if (capability) reply({ available: false });
        else reply(undefined, 'DOCS_BARRIER_UNAVAILABLE');
      } else if (object(value) && Object.keys(value).length === 1 && isNativeBarrierError(value.error)) {
        if (capability && ['NATIVE_BARRIER_NOT_FOUND', 'NATIVE_BARRIER_UNAVAILABLE'].includes(value.error))
          reply({ available: value.error === 'NATIVE_BARRIER_NOT_FOUND' });
        else reply(undefined, value.error);
      } else if (isBarrierReceipt(value, id)) reply(capability ? { available: true } : value);
      else reply(undefined, 'DOCS_BARRIER_INVALID_NATIVE_RESULT');
    });
  } catch {
    reply(undefined, 'DOCS_BARRIER_NATIVE_FAILED');
  }
}
