# DeckWise selected text box position

A capability-gated native tool moves the entire rectangular text box containing
the current nonempty plain-text selection to explicit top-left millimeter
coordinates. It reuses the existing selection guard and calls native
`ImgApply({ Position: { X: xMm, Y: yMm } })`. Readback checks object, page,
selection, text, dimensions and exact position. Native Undo/Redo remains available;
there is no implicit save or retry after an uncertain mutation.

Scope is deliberately narrow: one current-slide, unrotated, unflipped, ungrouped,
non-placeholder rectangular text box. Peer slides may contain ordinary, ungrouped
non-placeholder images; their positions, crops, rotations, flips and bytes remain
unchanged. Connectors, groups, unknown object types, placeholders and other peers
withhold the capability. Both coordinates are finite numbers in 0..1000 mm; the
complete box must fit on the page before and after. Missing native geometry/type
APIs, readonly, load, busy, slideshow or master mode withhold the capability. The
tool does not infer positions, resize objects, change stacking, or edit
masters/animations.

The reverse test disables only the native setter: the success regression then
fails with an unverified-position error. Exact source bytes are restored and the
34-test position suite passes. Targeted Office regression covers 419 tests.
Independent production Office verification uses actual forward and reverse mouse
text selections, native position changes to (25,35) and (0,0), Undo/Redo, stale
command/read-only/off-page refusal, and two real saves/reopens. All text, other
geometry/styles and three notes are checked; native object IDs regenerate on
export. The initial fixture had no media. B49 additionally moves a text box on an
actual mixed-content page with four shared PNG images, including a peer image
rotated and flipped; image geometry, references and exact payloads survive native
save/reopen and seven package corruption negatives reject. Main Chat adoption of
this tool is a separate A-owned integration gate.

Evidence and runnable commands live in the parent workspace under
`docs/validation/courseware/parallel/B/b16-20261019`. An initial plain static
server could not serve `/editor` and the probe timed out before editing; the final
run uses the existing production server pattern with editor routing and isolation
headers. No frozen build, source file or vendor SDK was changed.
