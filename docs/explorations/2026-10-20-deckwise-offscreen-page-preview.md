# DeckWise offscreen source-page preview

A36 added an optional one-based page selector in its independent Office consumer. B66 ports that finite selector into the canonical B producer so complete B replay patches preserve the capability already used by Main.

The existing readonly, clean, empty-history, binding, source and generation fences remain. The selected slide is drawn directly with native CGraphics; no navigation, selection, SDK prototype or history mutation is used. Both selected page token and displayed page remain checked after drawing. Native text omission validates the requested slide/layout/master, not the displayed page. Raster text still requires explicit clean-plate review.

The embed route bounds page to 1–5000 and rejects pages outside the opened source. Omitted page retains the original current-page behavior. Unit tests cover non-current drawing, malformed/missing page, rendered-token changes, display changes and native history changes. Real production Office tests use saved sources and native save/reopen; Main/Chat/adoption registration remains A-owned.
