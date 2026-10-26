# Fixed native barrier bridge

The online presentation bridge now transports a finite native barrier capability
and three fixed operations. It reuses the existing pinned plugin installation,
owned descendant/origin handshake, one MessagePort, session/request IDs and
per-mount generation. There is no generic method name, script, arbitrary native
editor access, new assistant, publication button or server authorization here.
The offline editor and signed launch configuration are unchanged.

## Public mount API

`mountDocsOnlineEditor(...).communityReads` retains `readPresentation` and `close`
and additionally exposes:

```ts
barrierCapability(): Promise<{ available: boolean }>;
quiesceAndDrain(barrierId: string): Promise<DocsNativeBarrierReceipt>;
barrierStatus(barrierId: string): Promise<DocsNativeBarrierReceipt>;
resume(barrierId: string): Promise<DocsNativeBarrierReceipt>;
```

Barrier IDs must match the native extension's lowercase UUID format. The owned
plugin port accepts only `native_barrier_capability`, `native_barrier_quiesce`,
`native_barrier_status`, and `native_barrier_resume` in addition to the existing
bounded text reader. Exact keys are required. Callers cannot select another SDK
method or provide code. Requests share one queue; no operation is automatically
replayed.

Capability uses `executeMethod('DeckwiseBarrierStatus', [randomUuid])`, never
`callCommand`. A finite `NATIVE_BARRIER_NOT_FOUND` response proves the method is
installed; an absent-method `undefined` response reports `available:false`.
The probe is usable while the extension blocks ordinary plugin commands. It
only detects the fixed extension, not current server authorization or readiness
to acquire a barrier. Older engines keep their ordinary text-read ability.

The three operation names map to exactly `DeckwiseQuiesceAndDrain`,
`DeckwiseBarrierStatus` and `DeckwiseResume`. Native results are checked in
`lib/docs-community-barrier.ts` before leaving the plugin port and again on the
client. Receipts contain only:

```ts
{
  barrierId: string;
  state: 'draining' | 'held' | 'uncertain-held' | 'resumed';
  serverIndex?: number;
  error?: string;
}
```

The matching barrier ID is required. `draining` has no index/error;
`held`/`resumed` require a nonnegative safe integer index and no error;
`uncertain-held` requires a finite native error and may retain its former index.
Unknown keys, IDs, states, oversized/invalid indices or unknown errors fail
closed. Arbitrary native error text/configuration is never forwarded verbatim; only the
listed error codes cross the port.

Finite native errors have the `NATIVE_BARRIER_` prefix and one of these suffixes:
`INVALID`, `BUSY`, `UNAVAILABLE`, `NOT_FOUND`, `ACK_UNKNOWN`, `DISCONNECTED`,
`SAVE_UNAVAILABLE`, `SAVE_UNKNOWN`, `RESTRICTION_LOST`, `STATE_CHANGED`,
`SETUP_FAILED`.

## Authenticated parent entry protocol

Existing `deckwise-docs-online/1` / `channelId` / `requestId` ownership checks still
apply before request consumption. Every new request requires the current
`generation`. The parent can send these exact payload additions:

| `type`               | Other required field | Reply                                                   |
| -------------------- | -------------------- | ------------------------------------------------------- |
| `barrier-capability` | none                 | `type:'barrier-capability', result:{available:boolean}` |
| `barrier-quiesce`    | `barrierId`          | `type:'barrier', result:receipt`                        |
| `barrier-status`     | `barrierId`          | `type:'barrier', result:receipt`                        |
| `barrier-resume`     | `barrierId`          | `type:'barrier', result:receipt`                        |

Replies include the original request ID and current generation/project ID/
session epoch/document key/Docs user ID. Late results from replaced, aborted or
closed mounts are suppressed. Foreign origins/windows/channels, reused request
IDs, stale generations, unexpected fields and invalid UUIDs cannot dispatch a
native operation.

A legitimate native refusal is returned as `type:'error', code:<finite native
code>`. Missing capability maps to `ONLINE_NATIVE_BARRIER_UNAVAILABLE`, an
outstanding request to `ONLINE_NATIVE_BARRIER_BUSY`, a transport timeout to
`ONLINE_NATIVE_BARRIER_ACK_UNKNOWN`, and other failures to
`ONLINE_NATIVE_BARRIER_FAILED`. Existing malformed/stale/replay envelope errors
remain unchanged. Barrier errors never call `stop`/`destroy` or turn into the
mount's native-error event. Main must correlate the request instead of treating
a barrier error as an overall native editor failure.

## Uncertainty and existing reads

An older SDK's missing method or a legal native refusal rejects only that call;
it does not close the text-read channel. A barrier timeout rejects that call and
keeps the port/native view alive, without canceling, retrying, resuming or
fabricating a native acknowledgement. After an outstanding native callback
arrives, the caller can explicitly query the same barrier ID. A busy response to a read or barrier call
remains retryable by explicit action; it does not drop the original operation.
Late replies cannot resolve a different request.

The existing text-reader timeout/cancellation behavior stays unchanged. Invalid
native output still invalidates the owned port; it does not destroy the native
editor. Abort/close invalidates the port and suppresses late responses. Neither
`onDocumentStateChange(false)` nor a successful method invocation means `held`.
An `uncertain-held` receipt stays uncertain and may never be reinterpreted as
idle or resumed.

The fixed SDK currently blocks all ordinary `callCommand` calls while held,
including generic text reads. Such a read returns the existing recoverable
`DOCS_PLUGIN_UNAVAILABLE`; capability/status/resume do not rely on that reader.
Publishing must reconcile the actual archived PPTX through the trusted backend/
Host path rather than assuming a live read is available while held.

## Delivery boundary

This slice has focused plugin-port/parent-entry tests and Office type checking.
It does not start another engine or claim a real end-to-end barrier through the
new product bridge. The independent native SDK proof and exact 9.3.0-138 overlay
remain separate prerequisites. Production deployment/capability rollout, Desktop
Main invocation, all-client registration/admission, archive persistence, mapping,
Strict, IME, interrupted native saves and background mutations are not closed by
this bridge. `available:true` is not an all-client publication barrier.

Validation on 2026-10-24: six focused unit files, 182/182 tests passed
(`docs-community-plugin`, `docs-community-plugin-host`, `courseware-online-entry`,
`docs-online-plugin`, `docs-online-editor`, `docs-native-barrier-extension`).
Office `tsc --noEmit` passed. Scope includes old-engine fallback with reads
preserved, fixed dispatch, malformed input/output, original origin/window/session
checks, late generation/abort results, finite errors, and timeout/busy/late-result
status recovery without destroying the native editor. No engine, service,
paid call, new broad E2E, deployment or publication completion was performed.
