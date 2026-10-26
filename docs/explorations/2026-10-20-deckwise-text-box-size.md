# Native selected text box dimensions

### Selected text box dimensions

`set_slide_text_box_size` accepts exactly `{widthMm, heightMm}`: finite numbers
in 1–1000 millimeters. It resizes the entire ordinary rectangular text box
containing the current nonempty plain-text selection and preserves its top-left
position. The complete box must fit on the current page. The same narrow native
selection/geometry guard as text-box positioning applies; ordinary image peers
are allowed, while groups, placeholders, connectors, rotations and flips of the
target are excluded. Capability is withheld in readonly, loading, busy,
slideshow/master or unsupported contexts.

Use a fresh advertised capability and edit generation. Native
`ImgApply({Width: widthMm, Height: heightMm})` is followed by exact dimensions,
unchanged position, text, selection, object and page readback. A verified result
is `{widthMm, heightMm, page, verified: true}`; uncertain results require visual
inspection/Undo and must not be retried automatically. Text may reflow or
overflow. No implicit save, auto-fit, group resize or arbitrary XML/JS is exposed.

Validation uses the normal production Office build, browser mouse selection, native Undo/Redo and two real host-acknowledged saves/reopens on a three-page PPTX with notes and shared image payloads. Independent ZIP comparison permits only target dimensions and measured native save metadata/ID normalization. Evidence is in the parent workspace at `docs/validation/courseware/parallel/B/b75-native-text-box-size-20261020`. This is module verification; actual Main Chat and installed-package acceptance remain A-owned gates.
