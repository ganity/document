# Host-inspected finite fade transport

The default editor context withholds `set_slide_fade`. The trusted Host first
checks the latest clean saved PPTX with the B81 authoring inspector and retains
its result privately, bound to authoritative canonical mapping/locks and current
project/document/session/edit generation/page/source hash. Inspection facts are
not permission. The bound parent relays Host-built approval; this transport does
not authenticate canonical ownership. Main must never accept a model/Renderer RPC
proof as authority. Existing high-risk Host permission must separately allow the edit.

`document:review-fade` accepts `{binding, authorization}`. Authorization is exactly
`{editGeneration, page, sourceHash, slideHash}`, with prefixed SHA-256 hashes from
that retained result. It is copied and bound to the current native snapshot.
`authorization: null` revokes it. A valid grant adds a no-argument capability to
the existing context without changing the native edit generation. Repeated context
reads stay stable; native context changes withhold it. No hash enters that context.

`document:execute-fade` accepts `{command, authorization}`. The command is the
existing binding/request/generation envelope with `tool: "set_slide_fade"` and
empty `input`. Both the review and this dedicated Host channel are required;
ordinary `document:execute` cannot apply fade. The one-use review is consumed by
execution attempts. No model-selected page, effect, duration, source or script.
The existing B80 native helper applies one current-slide smooth click fade at
750 ms, verifies peers/timing/objects/notes/global loop and never saves or retries.

Both messages retain bound-parent source/origin, DeckWise-mode, request-ID,
pending-save and uncertain-save guards. Dirty source refuses review/execution.
Save/commit/readonly messages revoke the grant; open/session reset and native
page/selection/history changes invalidate it. Any intervening ordinary command
consumes the captured snapshot, requiring a fresh review. Main must revoke on
canonical/source/lock changes even if the SDK did not change; the Office process
cannot establish Host canonical ownership or teaching revisions.

No default document-tool registry or standalone assistant is expanded. SDK
readiness alone never grants authority. Source inspection, authoritative Main
binding and permission integration remain separate acceptance gates.
