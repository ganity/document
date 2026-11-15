# Selected-text italic

The finite `set_slide_text_italic` command accepts exactly `{ enabled: boolean }`.
It uses the native `put_TextPrItalic` setter and `get_TextProps().get_TextPr().get_Italic()`
readback on one nonempty current-slide plain-text selection. Native readiness,
readonly, master, slideshow, busy, current shape and selection checks reuse the
existing text-formatting path. Unknown fields and stale editor generations reject.

Both setting and clearing are explicit. The command verifies unchanged page,
shape, selection and complete text after dispatch. An uncertain result is never
retried or reported as saved. The host owns permission, explicit Save and source
association review. This does not enable general rich-text editing or a new UI.

Unit coverage includes invalid inputs, missing capability, reverse selections,
readonly, cancellation and uncertain readback. Native browser validation uses a
mixed-style presentation, mouse selections, Undo/Redo and two saved reopen rounds;
package checks preserve non-selected formatting, peer pages, notes and media.
