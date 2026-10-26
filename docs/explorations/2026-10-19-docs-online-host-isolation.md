# Isolated Docs launch module (W13 prerequisite)

`lib/docs-online-editor.ts` is an unregistered online host module. The existing
`onlyoffice-editor.ts`, offline vendor, converter, save patches and local history
are unchanged. No additional assistant or editing UI is installed.

The caller supplies a Host-pinned Docs API URL, the server's original signed JSON
configuration, an expected document key/user binding, and a lifecycle AbortSignal.
The caller also supplies the same-origin `public/docs-online-host.html` URL.
`mountDocsOnlineEditor` waits for this static host, loads the API in its window and returns
`destroy()`. It never replaces the outer window's offline `DocsAPI` or `editor`.
The static document has no offline scripts. It supplies a real URL origin:
`about:blank` makes the native API generate `parentOrigin=null`. Cross-origin,
unexpected-path and missing-marker host documents reject. Load failure, missing API, construction failure, timeout and cancellation remove
the child frame. Cancellation during construction also destroys the late instance.

The initial finite configuration supports editable/view-only PPTX, identified
users, fast coediting, fixed mode and autosave. Unknown fields reject. Server and
browser must agree the exact configuration before signing; the browser neither
rewrites it nor adds local download/save events. Unsupported optional fields need
an explicit, tested extension. The module has no native editing capability yet.

`parseDocsOnlineConfig` checks JSON shape, URL schemes, token syntax and the
expected binding. **It does not verify the JWT, server authorization, source
revision, callback identity, session epoch, or durable persistence.** The expected
binding must come from the authorized Host session, not document content. The
API script is trusted executable code: pin its deployment in Host; the child
window provides lifecycle/global isolation, not a security sandbox.

`onDocumentReady` and `onDocumentStateChange` are forwarded only. They are not save
receipts and cannot authorize a clean snapshot, candidate adoption or publication.
No local converter, `asc_DownloadAs`, blob source or single-user save patch runs
in this module. The existing direct offline SDK bridge must not be pointed at a
cross-origin online editor. A separate Community plugin bridge remains pending.

C3d now delivers server-issued document bindings, trusted callback archival,
correlated force-save requests and final-save key retirement. Its server tests use
local protocol stubs; real Docs integration is still pending. `session_closed`
and `46309 DOCUMENT_KEY_RETIRED` must stop reuse of the old key. Neither state is
a successful source snapshot. This launch module does not submit or poll saves.

Server prerequisites still needed are distinct authenticated Docs user IDs,
scoped source URLs and signed editor configuration bound to the existing C
session. Source/teaching authority and promotion remain separate from callback
archival. No platform endpoint or canonical contract is invented here. A owns
Host installation and online/offline session switching; B's CE command bridge
and freeze/sync coordination remain pending.

Sources checked for the module's native configuration fields:

- [Document key and source URL](https://api.onlyoffice.com/docs/docs-api/usage-api/config/document/)
- [Callback, coediting and user configuration](https://api.onlyoffice.com/docs/docs-api/usage-api/config/editor/)
- [Server JWT responsibilities](https://api.onlyoffice.com/docs/docs-api/get-started/how-it-works/security/)

Validation: directed unit tests include existing offline lifecycle regressions.
The B21 browser harness checks two isolated scripted API instances and cleanup.
A second phase loads the actual local vendor `api.js` and verifies the inner
editor URL carries the real parent origin; its editor endpoint is absent. Neither
phase proves ONLYOFFICE Docs synchronization or save/reopen. Negative controls disable binding equality and late-event checks;
the corresponding tests must fail before restoring the exact source bytes.
Fixed-version Docs licensing/capacity, real Community plugin commands, two-user
conflicts/reconnection and trusted saved PPTX readback remain unverified.

## Registered editing-window identity (2026-10-24)

The owned courseware online entry accepts the existing
`tenant-<positive integer>-user-<positive integer>` identity and the registered
window form `tenant-<positive integer>-user-<positive integer>-client-<lowercase UUID>`.
The complete value remains bounded to 128 characters; malformed, uppercase,
trailing-whitespace and oversized client identities reject before mounting.
The entry forwards the exact value without stripping the client suffix.
`parseDocsOnlineConfig` / `mountDocsOnlineEditor` retain their generic identified
user compatibility and require exact equality to the authorized user pin. A
configuration signed for another window of the same tenant/user still rejects;
matching the base user alone is insufficient. Server registration and JWT
verification remain server responsibilities.
