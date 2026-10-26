# DeckWise selected-text bold

`set_slide_text_bold` accepts exactly `{enabled: boolean}` for an existing nonempty
plain-text selection in one current-slide text object. It uses synchronous native
`put_TextPrBold`, preserving native history, then verifies current slide/object,
selection, complete shape text and native `get_Bold`. It never selects text, saves,
loads fonts or executes caller code. Uncertain post-dispatch state fails without
retry; the user must inspect and may Undo.

Discovery and execution reuse the existing selection guards. Readonly/view/master,
loading, busy, slideshow, foreign/ambiguous selection and missing native functions
withhold capability. The finite bridge additionally checks binding and generation.
Word's existing bold tool is unchanged.

B86 exercises the compiled finite Agent against a newly built production Office:
mouse selection across differently colored/italic runs, true then explicit false
with reversed selection, Undo/Redo and two acknowledged save/reopen rounds. ZIP
readback compares every character property outside the selected bold flags, other
slides, geometry, notes, media and relationships. This is native/module evidence;
actual Main Chat discovery, permission integration and packaging remain A-owned.
