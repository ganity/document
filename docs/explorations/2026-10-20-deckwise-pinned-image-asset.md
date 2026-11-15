# Host-private pinned image replacement

`lib/agent-plugin/slide-image-asset.ts:replaceSlideImageFromAsset` connects an
already selected immutable PNG asset to the proven B32 native image replacement.
It captures native document/history/selection before awaiting asset retrieval,
then checks the same Host project/document/session/editGeneration, writable/save
state and installed image capability around retrieval, SHA256 and native preload.
Only the selected ordinary uncropped image changes; no save or automatic retry.

The Host supplies current-context reads and authorized asset bytes. The metadata
pins (`assetId`, lower-case SHA256, PNG MIME, byte length and dimensions) are copied
before await; returned bytes are privately copied before WebCrypto hashing. A
matching hash is integrity proof, never project membership or user consent. URLs,
paths and credentials are neither consumed nor returned. C4a's existing asset
metadata can be projected onto these fields without a new server contract.

A owns selection/review authorization, canonical asset references, authenticated
transport, request replay protection, public registration and capability mounting.
C owns current-user/team authorization and asset retention. This module must not
be registered as an unauthenticated tool or exposed as arbitrary byte/URL input.
Online Docs writes remain unavailable. Finite rotation and boolean horizontal/vertical
flips preserve through native replacement, provided both the original box and rotated
bounds fit the slide. Crops, groups, masks and invalid transforms use a reviewed
latest-clean-PPTX copy candidate; never reset source crop to enable native apply.

Validation: `pnpm exec vitest run test/unit/slide-image-asset.test.ts
test/unit/slide-image-replacement.test.ts test/unit/deckwise-editor-bridge.test.ts`;
`pnpm exec tsc --noEmit`; scoped oxlint/Prettier. Actual Office evidence and A's
unmounted consumer example are in `docs/validation/courseware/parallel/B/b41-20261020`
in the Train workspace. No new dependency or public protocol is introduced.
