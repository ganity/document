# Current-page preview from an open PPTX

Single-page scene generation can accept a representative rendered from the exact reviewed PPTX/page, but its caller previously had to supply that image through an unspecified Host renderer. The OnlyOffice slide object already exposes `getBase64Img()`, which draws the native slide to PNG without changing document history.

The existing DeckWise embed now offers a read-only `document:render-page-preview` request. It requires the current editor generation, current slide target, matching binding, and the source/page tokens captured when that PPTX was opened. It rechecks editor, slide identity, selection/session freshness and page token after rendering. The response carries PNG bytes plus source hash/page identity so the generation caller can bind its reviewed candidate to the bytes actually rendered.

The caller still owns project authorization, visual/clean-plate assets, candidate review/adoption and stale-write compare-and-swap. This endpoint does not render an arbitrary page, accept paths or remote URLs, erase image backgrounds, or write the PPTX. It is unavailable if native source/page lineage was not captured. PNG output is bounded at 8 MiB.

The preview endpoint requires a fresh read-only PPTX copy with empty native undo history and no unsaved changes. Even a fully undone edit is refused because the opened source hash must describe the rendered page. After editing or saving, reopen the latest clean PPTX in a separate read-only session before requesting a source-bound preview.
