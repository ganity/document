# DeckWise native PNG replacement

`prepareSlideImageReplacement(png, signal?)` is an unregistered Host-only module
in `lib/agent-plugin/slide-image-replacement.ts`. It returns a one-use synchronous
`apply()` with `{ page, imageId, verified: true }`. It does not save the document.
The existing public tool whitelist is unchanged.

Accept one selected, writable, ordinary rectangular image that fits on the current
slide. Reject crops, masks, tiling, transparency, rotation, flips, groups,
placeholders, text selection, master mode, slideshow, busy state and missing
native tracking APIs. PNG input is local bytes, 33 bytes to 4 MiB, dimensions
1..4096; the native decoder must confirm those dimensions. No model URL is fetched.

Preparation registers and preloads the PNG through the existing native cache.
`ImgApply({ ImageUrl })` is synchronous only after that preload. Native document,
slide, object, selection, geometry, history and pointer state are captured and
rechecked after loading and immediately before apply. Abort/timeout cannot later
dispatch an image command. Post-dispatch uncertainty requires inspection/Undo;
there is no automatic retry. The Host must separately bind the asset and command
to current project/document/session/request, capability and edit generation,
including checks around asynchronous work. This module is not an authorization
boundary or an online Docs write tool.

Actual SDK 9.3.0.133 testing found ordinary native image replacement clears an
existing crop. Therefore cropped images are rejected, without automatically
resetting the crop. Unsupported requests should use a candidate built from a
fresh clean PPTX copy, subject to the existing adoption checks. They must never
write the open source or restore an old SVG.

The final real Office run verifies crop refusal, stale page refusal, cancellation,
replacement, host-acknowledged save and reopening. A separate native crop reset is
fixture setup only and changes its geometry; preservation comparisons start from
the saved/reopened uncropped baseline. The output contains the exact replacement
PNG, unchanged other PNG, three page texts and three notes, and unchanged image
position/size. Office export renumbers shape IDs, updates the modified timestamp
and revision, and removes the target image's empty stretch element. Package checks
allow only those measured normalizations, with six deliberate corruption checks.
This is neither arbitrary PPTX preservation nor a main Chat/model acceptance.

Directed regression: 88 tests across image replacement, editor bridge and text box
position; TypeScript, scoped lint and formatting pass. Reproducible standalone
build, real run, actual interfaces, three patch routes and hashes live in the
parent workspace at `docs/validation/courseware/parallel/B/b32-20261020`.
