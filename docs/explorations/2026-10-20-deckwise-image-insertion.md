# Native image insertion preparation

### Host-only native image insertion

`prepareSlideImageInsertion(png, {xMm, yMm, widthMm, heightMm}, signal?)`
prepares one already authorized/reviewed local PNG for the current ordinary slide.
Positions are explicit nonnegative millimeters, dimensions at least 1 mm, and the
whole box must fit on the page. The explicit box stretches the image; no automatic
aspect-ratio fitting or overlap avoidance is implied. The canvas selection must
be empty: the helper does not discard a user's active object/text selection.

PNG signature/chunks/CRCs, the 4 MiB budget, 1–4096 pixel dimensions and native
registration/decoded dimensions are checked before editing. Preloading changes
only the image cache. Page, history, interaction, peer geometry/text/image data,
readonly/loading/busy/slideshow/master/notes and slide-lock checks run again after
preload and immediately before applying. Geometry is captured before awaiting.
A stale, cancelled, failed or timed-out preparation never sends an insertion.

The returned frozen preparation has `page`, `box` and a single-use synchronous
`apply()`. It uses the SDK's existing native action/controller image creation path
with the requested box in one Undo action. Native readback verifies exactly one
appended ordinary rectangular picture, its selection, URL/geometry, original
peer order/data and page dimensions. A successful receipt is
`{page, imageId, box, verified:true, saved:false}`. An uncertain native result
requires inspection/Undo and must not be retried automatically. Saving is explicit.
Native object IDs and the newly registered media filename can change on export.

This is an independent Host module, **not a newly advertised bridge/Chat tool**.
The caller owns current session/edit-generation/save checks and authoritative
asset authorization/review immediately before preparation and apply. No URL,
model code, filesystem path, provider call, source-file write or authorization
grant is accepted here. Existing image replacement shares the unchanged bounded
PNG validation/preload path. Public bridge/Main/IPC/Agent mounting remains a
separate owner integration gate.

Validation uses 160 directed image/asset/bridge tests, a separate built IIFE module with the unmodified B75 production Office/CSP, one native insertion/Undo/Redo, two acknowledged saves/reopens, and independent exact-picture/relationship/media/peer/notes ZIP checks with ten corruption negatives. Evidence: `docs/validation/courseware/parallel/B/b77-native-image-insertion-20261020` in the parent workspace. No new Main Chat, asset authorization, installed-package or full advanced-feature acceptance is claimed.

Reverse verification temporarily disables all three context rechecks: the new insertion suite fails; exact source bytes are restored before the final passing suite. See `negative-no-context-guards.log` and `negative-control.json`.
