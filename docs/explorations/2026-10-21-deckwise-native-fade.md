# Current-slide native fade preparation

`prepareSlideFade(signal?)` prepares one native smooth, click-advanced 750 ms
fade on exactly the selected current slide. Preparation does not edit. The
single-use `apply()` calls the actual SDK `ApplySlideTransition` and returns
`{page, transition:'fade', durationMs:750, verified:true, saved:false}` after
readback. Save remains explicit; uncertain native results require inspection or
Undo, never automatic retries. One native action supports Undo/Redo.

This is a Host-only module. It adds no bridge command, capability, Agent tool,
registration or UI. Before mounting, the Host must bind the current document,
session, edit generation and save state, inspect/review the latest clean PPTX
target transition using the existing bounded fade-candidate preflight, and
withhold the native command if source extensions, sound or other unsupported
semantics cannot be ruled out. The SDK's seven transition fields alone cannot
prove those source-file properties. `canSetSlideFade()` checks native readiness,
not source compatibility or authorization.

Native readiness rejects multi-slide selection, a different selected page,
readonly/loading/busy/slideshow/master/notes mode, missing APIs/history,
collaboration/delete/transition locks, automatic advance, non-click advance,
non-fade transitions, black fades and unknown transition fields. The native
constructor must leave all unrelated fields undefined. In particular, supplying
`ShowLoop` would change the whole presentation, so it is deliberately omitted.

Preparation pins the native document/page/controller, selection, interaction and
history. Immediately before applying it rechecks those pins, all peer transition
values, shape identity/text/geometry/image data, notes and timing references.
After applying it verifies the same scope, unchanged global loop and exact fade
readback. It does not redraw objects or edit object animations. Native IDs can
change on ordinary export; package checks resolve animation targets against the
corresponding remapped shapes rather than ignoring target references.

Evidence lives in the parent workspace under
`docs/validation/courseware/parallel/B/b80-native-slide-fade-20261021`: directed
unit/guard regression, exact-context negative control, separately built module
with the unchanged B78 production Office, native Undo/Redo and two acknowledged
save/reopens, and independent package/animation/notes/media checks with corruption
negatives. A first prefix-rewritten animation fixture lost effects on its initial
converter save and is retained as a failed negative; the accepted fixture preserves
the existing `p:` XML prefix and retains effects on both target and peer pages.
This does not certify arbitrary animation or advanced PPTX preservation, live
model quality, public Main Chat integration or installed-package acceptance.
