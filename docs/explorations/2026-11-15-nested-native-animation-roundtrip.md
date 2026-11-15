# Nested native animation roundtrip repair

SDK v9 `CSpTgt.assignConnection` checks only the slide top-level shape tree.
Valid group-child timing therefore disappears during import; saving after that
import permanently removes the effect from the package. The saved teaching cue
can remain, making a count-only readiness check misleading.

Install an idempotent target-resolution guard before BinaryPPTYLoader construction
and during iframe preparation. Resolve only actual group members whose ancestor
chain reaches the slide tree; retain SDK rejection for detached/cyclic targets.
Do not fake the child's parent or modify the vendor bundle.

Native motion save readback now verifies exact target/preset/subtype/duration and
claims each mirror only once. Playback considers child and ancestor effects so
an unrelated child effect cannot hide a valid ancestor mirror.

Validation: 41 unit tests, scoped oxlint and TypeScript checks pass. A preserved
real Master lesson passes a native motion replacement/undo/redo/save/reopen E2E.
The original fixture loses child pulse timing with the old SDK path; the guarded
path retains it. Real repaired Master pages 1/2 and scene pages 1/3 play naturally
without captured SDK or page errors, with identical teaching and audio bindings.
Evidence: `/home/jhz/deckwise-validation/all-fixes-20261115`.

The paired Desktop build checks source mirrors before playback-only group
projection consumes them. It preserves parser warnings and does not infer full
OOXML fidelity from successful playback. Existing paid media are retained.
