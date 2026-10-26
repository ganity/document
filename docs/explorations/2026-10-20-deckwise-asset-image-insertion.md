# Asset-bound native image insertion

### Asset-bound native image insertion

The optional `insert_slide_image_from_asset` command accepts exactly
`{assetId, xMm, yMm, widthMm, heightMm}`. The independent existing-Chat Agent
module exposes `ppt_insert_slide_image_from_asset`, adding current `editGeneration`.
Only a PNG asset ID enumerated by the latest capability is accepted. The Host
continues to own authorization/review of both asset and explicit placement;
capability metadata is not an authorization grant. No model URL or bytes are
accepted in the model command. Native image identity and host-only fields are
removed from model-visible receipts.

The current native slide must be writable with an empty canvas selection.
Coordinates are finite 0–1000 mm; dimensions 1–1000 mm; the whole rectangle must
fit on the page. The explicit rectangle stretches the image and may overlap
content, so it requires review. No automatic crop, aspect fitting or placement
inference is provided. Native single-action insertion/readback uses the existing
Host-only preparation. Original peers, page, exact rectangle and new image are
verified; uncertain writes require inspection/Undo and are never auto-retried.

The existing `document:review-image` metadata and `document:execute-image`
command-scoped fresh-byte path are reused. Reviewed metadata is independent of
bytes; byte leases clear on success/failure/reset. Current binding, edit generation,
asset pins, capability, native target, readonly/cancellation and existing embed
save-exclusion checks remain required. Asset hash/length/dimensions and native
PNG decode/CRCs are verified before mutation. Replacement retains its original
selection and receipt contract. Successful insertion receipts contain
`assetId`, `sha256`, `page`, `xMm`, `yMm`, `widthMm`, `heightMm`, `verified:true`;
saving stays explicit (`saved:false` at the Agent boundary).

This is an additive finite producer contract. Older peers continue with their
existing tools; they must not receive an unknown capability. B's producer now
contains ten finite tools including pending B75 resize. The minimal current-A
proposal adds insertion alone to A's eight-tool baseline and preserves A's privacy
and async fixes. Do not mix producer/full recovery patches with the current-A
proposal. Main/IPC/Host permission dispatch and actual owning-Chat acceptance
remain A-owned; authoritative asset authorization remains C-owned.

Verification uses actual built Agent/production Office via the ordinary reviewed-image embed messages, native Undo/Redo and two acknowledged save/reopen rounds, exact ZIP picture/media/peer/notes comparison and ten corruption negatives. A local metadata/authorization fixture is not real Main Chat, C authority or installed-package acceptance. A reverse capability-guard control fails and restores exact source bytes before final tests. Parent-workspace evidence: `docs/validation/courseware/parallel/B/b78-native-insert-asset-20261020`.
