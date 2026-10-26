# DeckWise active SDK export exclusion

`pendingHostSave` covers host-managed saves only. A legacy DeckWise export
without `requireHostCommit` left context, native commands and document opening
available while the SDK was serializing. A second host-managed request could
also enter, fail in the underlying save channel and mark the export uncertain.

The embed router now reads `hasPendingEmbeddedSaveRequest()` from the existing
save-channel owner before tools, document replacement or another save. No save
state is duplicated. Host confirmation and uncertain-export guards retain their
existing behavior. Ordinary non-DeckWise embeds and manual SDK edits keep their
current behavior; edits after the captured save point remain dirty.

Four directed regressions hold a legacy export pending, verify rejection before
bridge/reset/save dispatch and then verify a successful legacy completion and
subsequent host-managed export/acknowledgement. They all fail on the old router.
Real built-Office validation delays the actual converter promise, exercises these
same parent messages and reads the exported PPTX back after a complete reload.
See the B56 handoff for commands, bytes, timing fixture and remaining integration
gates. Model, asset authorization and Main Chat are outside this regression.
