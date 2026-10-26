# DeckWise selected text color

Scope: opaque solid sRGB on one current-slide top-level plain-text selection.
Reuse the font-size selection guard without coupling either setter to the other.
No assistant, provider, source-file write, public Main/Host registration or new dependency.

The bundled native `put_TextColor(Asc.asc_CColor)` writes text Unifill. A real
production probe saved `336699` correctly while `get_TextProps().get_TextPr().get_Color()`
reported black from the legacy Color field, even after a selection update. Readback
therefore checks `GetCalculatedTextPr().Unifill`: solid type 3, sRGB type 1, no
transparency or modifiers, alpha 255 and exact integer RGB channels. These SDK
shapes/constants are checked; unsupported versions withhold the capability or fail verification.

Regression tests cover native dispatch, exact fill, changed selection/text, readonly,
master, slideshow, load/busy and missing APIs, stale bound generation and invalid input.
Real partial/reverse range save/reopen and reverse verification are recorded under
`docs/validation/courseware/parallel/B/b14-20261019` in the parent workspace.
Advanced charts, masters and animation remain unsupported.
