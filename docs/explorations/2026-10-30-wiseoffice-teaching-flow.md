# WiseOffice teaching flow

The native teaching sidebar now presents a compact ordered list with the selected-step editor expanded inline. It reuses the Host draft/save, local media, source-bound attention and native animation paths. Native PPT content is unchanged by teaching edits.

## Interaction

- Select a step to edit its narration or inspect its media/native animation binding. The list and editor share one scroll container.
- Teaching opens at up to 360 CSS pixels of content width through the native RightMenu and Viewport layout APIs. The canvas resizes with it; closing, switching property panels or disposal restores the prior width. Narrow viewports reduce the requested width.
- Native-style icon controls keep undo, redo, refresh, save and ordering accessible by their existing names. Whole-lesson playback remains the primary labeled action. Only step headers are draggable, so authored text stays selectable.
- Insert narration immediately after the selected step. Use drag and drop or the existing up/down controls to reorder it. Native animation order remains enforced by the shared edit path.
- Spotlight has its own primary entry, with highlight as the secondary action. Both reuse the same mapped object checkboxes and validated native multiple selection. Cue start/end values are displayed in seconds and persisted as integer milliseconds.
- Narration tails expose None, Pause and Timed wait. A new Pause uses 400 ms. Timed waits accept 3–30 seconds and continue automatically. Editing only these settings preserves generated speech and attention cues.
- Whole-lesson, current-page native slideshow and canvas previews retain their existing ownership and cancellation paths.
- The native animation edit action selects the referenced target and opens the editor's existing Animation ribbon. Changes to animation content still use native save and source binding validation.
- Saving the same page retains teaching undo/redo; page/source refreshes clear it. Draft history remains bounded and independent of native PPT history.

## Compatibility

The cloud-web editor encodes timed waits in `pauseAfterMs` using a multiplier of ten. The native Host and fixed player instead validate and consume raw milliseconds with a 30,000 ms maximum. This change matches cloud-web's visible choices and playback behavior while retaining the native protocol. It must not copy cloud-web's encoded numbers into this Host.

Existing raw wait values are retained when opening and saving a page. Values below 3,000 ms appear as Pause and longer values as Timed wait. Only new or explicitly changed settings get the new defaults and validation.

Previously saved independent delays and interaction checkpoints remain readable and playable. The new UI does not create them, migrate them or discard them. Native animation checkpoints remain an internal completion mechanism, not authored pauses.

## Verification

Targeted checks cover narration-only tail settings and range validation, preservation of audio/cues, insertion at selection, selected inspector identity, reorder/undo, same-page save history, source/page lifecycle, attention mapping, native slideshow ownership and seven locale tables. The browser scenario exercises the real editor with local fixture speech/music/video and a Host bridge fixture; no paid generation is invoked.

The reverse check temporarily disabled tail editing, native animation navigation, reordering and automatic native selection, and then restored every file in a `finally` block. All five affected sidebar regression tests failed as expected. Disabling raw wait classification and narration summaries caused both helper tests to fail. The restored implementation passes the targeted suites.

Final validation: 102 tests passed across the five targeted unit suites; the complete real-editor Playwright scenario passed (3.3 minutes), including the native Animation ribbon, native target selection, full lesson playback, spotlight persistence, media transitions and save/reopen behavior. Root TypeScript, targeted oxlint, Prettier and `git diff --check` passed. The browser test clears its explicit animation target selection before exercising a later selection-change refresh, so that regression still changes the selection. Runtime media is bundled from the current desktop source and served locally. Paid speech generation and an installed desktop package are outside this browser fixture.

## Visual revision

The first implementation squeezed a detached editor and a step list into the native 220-pixel content column. The revision uses the existing native panel sizing API, a labeled primary lesson-play action, secondary icon controls, quieter borders and inline step expansion. It retains existing action names for accessibility and automation. Native width and canvas boundaries are asserted in the browser; close/dispose and narrow-window resizing are covered by the layout unit test.

Reverse verification: disabling native panel sizing made the layout test fail; moving the inspector outside the selected step made the narration regression test fail. Both changes were restored before the final checks. No new dependency, teaching action or playback timing rule was introduced.

The final visual-revision browser run passed after the header-only drag change. It asserts at least 350 pixels of teaching content on the desktop fixture, non-overlapping native canvas bounds, visible overflow for the inline inspector, and native property width restoration when opening the Animation ribbon. Screenshots are captured from the actual editor iframe, without the embedding demo controls.
