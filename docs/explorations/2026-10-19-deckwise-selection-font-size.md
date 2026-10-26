# DeckWise selected-text font size

The native presentation toolbar uses put_TextPrFontSize and get_TextProps().get_TextPr().get_FontSize(). A real SDK probe changed 32 to 44 pt and retained it after export/reopen before the capability was implemented.

The additive set_slide_font_size tool permits 1–96 pt in half-point increments, only on a supported nonempty plain-text selection in one current-slide text object. It validates native readiness, write permission, selection scope, and setter/readback availability; verifies unchanged text/selection and exact size after the synchronous native operation. The existing DeckWise bridge binds it to current document/session/generation and save fences. No source XML, generic script, font-family, color, position or image operation is exposed.

The real Agent module and independent production Office test change size to 44.5 pt, verify Undo to 32 and Redo to 44.5, save/reopen twice, and reject stale generation, readonly and page changes. All page text and notes remain intact. This is direct module verification; A still owns existing-main-Chat registration/permissions and full desktop packaging.

Reverse verification disables only the native setter: the selected-range bridge test fails. Restoring it passes. The failure log and final evidence are retained under the external Plan009 B12 validation directory. Uncertain native edits report failure without automatic retry; users can inspect and use native Undo.
