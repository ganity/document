# Native source-page rendering for reviewed candidates

The native `getBase64Img()` helper renders a cached thumbnail at roughly one third of the ordinary pixel density. A measured source-page fixture returned only 410 by 256 pixels, which loses text and image detail when used as a visual input for a single-page candidate.

The existing preview module now uses the same native `CGraphics` and slide `draw()` path with a 1600-pixel longest edge and the original page aspect ratio. It draws directly from the opened PPTX rather than enlarging the cached thumbnail. The existing fresh read-only copy, empty history, latest generation, session/source/page identity, save and PNG budget checks remain. Extreme aspect ratios with a short edge below 64 pixels are refused before allocation.

Rendering temporarily enables the SDK shape-to-image flag. A `finally` block restores its prior value, resets the shared font-manager caches as the SDK thumbnail path does, and clears the temporary canvas even when native drawing fails. No source bytes, history points, tool registration or candidate adoption are changed. Actual quality still depends on the native renderer and supported PPTX contents; a clean plate remains a separate authorized and reviewed artifact.
