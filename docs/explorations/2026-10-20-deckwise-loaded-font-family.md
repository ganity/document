# DeckWise selected-text font family

The native `put_TextPrFontName` defers an unloaded-font edit through
`asyncFontEndLoaded`, which uses the selection at callback time. The finite tool
therefore checks all four `fontInfos` face indexes with the native `CheckLoaded`
reader and rejects an active `currentInfoLoaded` operation. Discovery never calls
`LoadFont` or `CheckFontLoadStyles`. It reuses the current plain-text selection
owner, changes the family synchronously, and verifies native family/text/selection.

A real production SDK probe found nine eligible names (some are font aliases).
A rebuilt production Office embed and built finite Agent tool verified capability
negatives, exact family, Undo/Redo, acknowledged PPTX export and full reload/reopen.
See root Plan009 B57 evidence for byte-level readback and the consumer fixture.
This does not establish Main Chat/live model/package integration or font download.
