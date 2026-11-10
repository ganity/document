# Embedded editor C handoff changed-file inventory

Date: 2026-11-10. New delta after 632c5201 is separate from the earlier upstream/integration history.

## New handoff delta

```text
M	bin/font-license-sweep.mjs
A	bin/local-system-fonts.d.mts
A	bin/local-system-fonts.mjs
A	docs/deckwise-c-import-handoff-2026-11-10.md
M	docs/fonts.zh.md
M	lib/converter.ts
M	lib/deckwise-editor-bridge.ts
A	lib/deckwise-object-selection.ts
A	lib/onlyoffice/guards/east-asian-punctuation.ts
A	lib/onlyoffice/guards/east-asian-vertical.ts
A	lib/onlyoffice/guards/image-color-key.ts
A	lib/onlyoffice/guards/image-soft-edge.ts
A	lib/onlyoffice/guards/mixed-script-line-metrics.ts
M	lib/onlyoffice/guards/x2t-worker.ts
M	lib/onlyoffice/iframe-guards.ts
M	packages/converter/src/docx-zip.ts
A	packages/converter/src/pptx-theme-fonts.ts
M	public/sdkjs/common/AllFonts.js
M	test/e2e/font-substitution.spec.ts
A	test/unit/deckwise-object-selection.test.ts
M	test/unit/docx-zip.test.ts
M	test/unit/embed-api.test.ts
M	test/unit/font-catalog-licensing.test.ts
A	test/unit/local-system-fonts.test.ts
A	test/unit/office-east-asian-punctuation.test.ts
A	test/unit/office-east-asian-vertical.test.ts
A	test/unit/office-image-color-key.test.ts
A	test/unit/office-image-soft-edge.test.ts
A	test/unit/office-mixed-script-line-metrics.test.ts
M	vite.config.ts
```

## Complete earlier branch history beyond inspected remote main

Do not treat these 201 inherited commits as 201 new C-task changes.

```text
632c5201 feat(wiseoffice): integrate courseware editor workstreams
283ca850 docs: record document lifecycle regression verification
aec0b49b perf(offline): reuse unchanged prepared runtime assets
498afad9 fix(agent): bound Anthropic requests through response consumption
91481733 fix(editor): preserve document identity and open ownership
a9256a32 fix(history): retain fresh snapshots past record expiry
6b7fbdfd fix(offline): wait for page control before warming editor assets
6e344d27 test(editor): verify assistant and offline follow-ups
9af4846c refactor(agent): separate panel owners and bound cache discovery
46676c56 perf(offline): prepare assistant assets by active capability
848ad372 feat(agent): stream bounded answers from cloud and local services
855e2b1b fix(agent): bound provider response bytes before decoding
0c39e773 fix(agent): omit unused tool fields from plain chat requests
e2d992e1 perf(history): reject oversized exports before reading bytes
dfc6da9f docs: record recovery and provider safety verification
72d0065e fix(agent): bound HTTP requests and classify Gemini failures
0d988801 fix(history): preserve recovery when quota retries fail
ac2dcade fix(editor): prevent duplicate errors during open recovery
6c9bdefd fix(ui): fit phone navigation beside visible home links
1b731576 fix(ui): apply compact homepage navigation after base styles
5c862fb4 fix(ui): balance compact headers and empty-state geometry
bd1a9ad2 fix(ui): preserve negative token arithmetic in rendered styles
62197c9c fix(ui): unify site design tokens and accessible command states
f44ec649 docs(assistant): record workspace consistency validation
64703a41 test(assistant): align history assertions with shared view header
de6b9bf5 fix(assistant): preserve destructive model command colors
2217a716 fix(assistant): unify workspace controls and navigation with design tokens
bbf14ed9 docs: plan assistant workspace token consistency
2427b273 docs(assistant): clarify follow-up delivery after merge
eaf9f0bb docs(assistant): record live preview density acceptance
cfa445d9 fix(assistant): retain readable language in narrow toolbars
4aa314bd fix(assistant): align compact composer controls
259ec4fb fix(assistant): preserve source disclosures and trusted read evidence
cc63f3b8 docs(assistant): refresh final regression evidence
120a3bb7 fix(assistant): settle Undo receipt status and focus
acac6afc docs(assistant): record integrated workspace acceptance
70cf71be test(assistant): assert capability-specific verified receipts
75412184 fix(assistant): keep completed receipt actions compact
e1430495 test(assistant): retain native Stop cleanup contract
905b08b4 feat(assistant): add precise location and owned change receipts
9804b76a test(artifacts): retain sparse payload coverage without array ambiguity
121ead15 docs(assistant): record structured artifact acceptance evidence
4d5c5388 fix(assistant): keep structured artifact save actions visible
76cd7361 feat(assistant): add editable structured artifacts and offline exports
682f0f11 docs(assistant): record artifact revision acceptance evidence
2abc2e35 feat(assistant): review artifact revisions without losing drafts
c09d272f docs(assistant): record standalone reference acceptance evidence
8e2dc834 feat(assistant): read bounded Office and PDF references locally
a746ac1d fix(connectors): validate search controls without control regex
138dc5a5 test(assistant): verify visited history restoration across sessions
13649429 fix(assistant): narrow citation preview state before access
21e0f3f9 docs(assistant): record captured source acceptance evidence
6731aa3f feat(assistant): ground source discovery and captured citations
df110290 docs(assistant): record result action acceptance evidence
615427c9 feat(assistant): stage contextual follow-ups and report save outcomes
e3e0abdd docs(assistant): record navigation regression evidence
3cec35ea fix(chat-ui): preserve active composition focus during navigation
8219bc0d fix(assistant): preserve draft and reading position across navigation
5f08f18c docs(connectors): record OAuth deployment acceptance steps
50babfb2 docs(assistant): define workspace craft acceptance contracts
ede436c6 fix(connectors): finish disconnect before authorizing again
731a3f20 fix(connectors): isolate canceled source operation ownership
71d41ab7 fix(assistant): make source and artifact recovery actionable
e82078ba fix(assistant): keep dictation feedback clear of drafts
cb16ae04 feat(assistant): adapt workspace controls to offline availability
3f856473 fix(assistant): clarify recovery after reference changes
cc4c3c38 feat(connectors): add optional Google Drive authorization backend
346c04a1 feat(assistant): add scoped sources and versioned artifact workspace
3fddd3c5 style(assistant): flatten surfaces and reduce nested controls
8c256763 fix(assistant): refine queue, search and writing setup interactions
beb7e2fd fix(editor): handle superseded startup document opens
b139a371 fix(editor): retain retries for the mounted document
3def8478 fix(editor): complete replacements once teardown starts
60e6c5da fix(editor): harden document opens and isolate assistant settings
b3539332 test(editor): preserve captured conversion methods in the seam probe
8124d248 fix(offline): defer runtime assets attached to assistant UI chunks
f26c8251 fix(editor): register the iframe wheel listener without inline code
d1ab14a0 docs(editor): record review fixes and regression evidence
50daadfb refactor(assistant): extract endpoint, controller and settings owners
44675a25 test(history): guard bounded recovery snapshot reads
80c19d53 fix(history): preserve recovery when snapshots exceed the budget
825c19e2 fix(agent-core): fit optional references within the context budget
df99e78f fix(editor): preserve newer edits across asynchronous saves
6f63fbeb chore(tooling): update workspace checks and jsdom fixtures
03e6b7cd fix(assistant): keep dictation drafts with their conversation
233ea9f4 docs(assistant): record final continuity acceptance
8fd0a599 docs(assistant): record actual timing review gates
dafba377 fix(chat-ui): bind presentation observers to stream ownership
b18a481f fix(assistant): distinguish superseded request timings
af171a78 perf(assistant): measure actual request timing
d734c1bb docs(assistant): record responsive interaction acceptance
0d11a010 docs(assistant): record inactive deletion browser acceptance
4e52f4a1 docs(changelog): describe streaming and conversation continuity
bd99bc15 docs(assistant): record history deletion review gates
008e32a0 fix(assistant): preserve active chat during unrelated deletion
9e1cb42b docs(chat-ui): record native scrolling acceptance
305db837 fix(chat-ui): cancel native follow before scrollbar dragging
42e2542b fix(chat-ui): retain native scroll ownership through growth
d54d6374 docs(assistant): clarify observable request timing boundaries
d44d8a03 docs(chat-ui): record streaming presentation acceptance
491b9865 fix(chat-ui): omit caret line from formatted streams
c7872113 docs(chat-ui): record streaming markdown review gates
b14d3b77 feat(chat-ui): render bounded streaming markdown with stable blocks
84d0d119 docs(assistant): define inactive history deletion continuity
350951e7 docs(assistant): review harness interaction patterns and scroll ownership
ede72f91 docs(chat-ui): record bounded history review gates
8ecff2bd perf(chat-ui): restore long conversations in bounded windows
dce073c7 docs(assistant): record search and budget acceptance evidence
acfb2372 docs(assistant): record search and context budget review gates
819b7e51 perf(agent-core): bound exact context budgeting probes
a28057cc fix(assistant): normalize and bound history search queries
39f45f2e docs(chat-ui): plan bounded Streamdown-inspired presentation
ad4312db docs(assistant): record follow-up continuity browser acceptance
108ad3be docs(assistant): record chat and voice continuity review gates
eaa3e564 feat(assistant): allow next-draft dictation while a request is running
4494faf9 fix(assistant): keep ungrounded chat independent of editor changes
0d40ba68 docs(assistant): record search identity and empty welcome acceptance
5a860cbc docs(assistant): clarify request scope and bounded history contracts
932edb74 fix(assistant): reflect pending request preparation in queue actions
39720ee9 fix(assistant): preserve pending request owners during history hydration
67044cf7 docs(assistant): plan follow-up continuity and performance fixes
b203ab69 fix(assistant): preserve drafts typed during startup history restore
4aa14556 docs(assistant): record final continuity review and local acceptance
88b0b492 feat(assistant): return to recent chats from the header
ab491910 fix(assistant): allow fresh send after cancelling a standalone apply
599d621e fix(chat-ui): preserve send focus across asynchronous preparation
86a4f5ab fix(assistant): serialize every document apply through physical settlement
e74bf3c5 docs(assistant): record selection motion and queue browser acceptance
844cb24b test(assistant): match held queue responses to the loopback JSON transport
a1ff532a fix(chat-ui): restore composer focus before Send disables itself
92602c31 fix(assistant): align and animate selected-text context disclosure
e0738f2d docs(assistant): record disclosure and conversation entry acceptance
85025cd3 test(assistant): discard refused PDF request before a fresh conversation turn
37f11bc3 fix(chat-ui): keep pending requests outside the composer focus boundary
e76dea69 fix(assistant): retain pending requests until execution preflight succeeds
4b0e2892 docs(assistant): describe queued follow-ups and responsive requests
359d5412 test(chat-ui): validate geometry in grouped button selectors
d0b695e2 docs(assistant): distinguish failed validation from an explicit Stop
bef32fe1 fix(assistant): require direct capability receipts before queue advancement
a53750b6 test(assistant): align imported reads and PDF greetings with scoped requests
379dada8 feat(assistant): queue continuous drafts with scoped host outcomes
94dba715 feat(assistant): own FIFO request settlement per conversation
ee3d1d3d feat(chat-ui): support optional host queues with separate Stop
22f6a9e7 fix(assistant): clear applying feedback after cancelled settlement
500c3ba4 fix(assistant): settle applying reviews with their original owner
0b9b56e4 docs(assistant): retain explicit proposal refinement during queue holds
40eff4d0 docs(assistant): preserve fresh questions after stopping an empty queue
151d92f4 docs(assistant): record request checkpoint validation
3d6fcfa8 test(assistant): bind reply fixtures to original document scopes
c8db79d0 docs(assistant): record request performance review and editor checks
4c7ada9d perf(assistant): bound all request reference envelopes
b7226938 docs(assistant): record request call and source-read measurements
bf3c2fce perf(assistant): bound request context and streamline trusted reads
e21ba5a7 docs(assistant): record Stop acceptance and queue dependencies
6229ca9d test(assistant): open history before selecting saved replies
c44d5c05 fix(assistant): classify outcomes for every document capability
c9457ed0 test(assistant): align panel lifecycle fixtures with lazy history
69e7f7d1 fix(assistant): acknowledge Stop immediately and guard turn ownership
ae254cd3 docs(assistant): record history performance and context behavior
253bb8da fix(assistant): amortize history search scheduling and announce progress
b62bdcb8 fix(assistant): bound history indexing and suspend cached pages
90315456 feat(assistant): page and search conversation history incrementally
dfcef646 docs(assistant): define continuity and responsiveness contracts
475308c8 fix(assistant): preserve dictation during first-use setup
d2600c8e docs(assistant): record experience review and acceptance
638c292b feat(assistant): show immediate request progress
a7fb4552 fix(assistant): complete localized feedback and keyboard focus
bf9e8b76 feat(assistant): refine change review and completion
6fcce0ca feat(chat-ui): clarify drafts and conversation results
c1c56824 feat(assistant): add contextual tasks and direct request routing
8163e93a docs(assistant): plan contextual workspace experience
717568d9 fix(docker): serve x2t WASM as gzip on plain HTTP
baf3c987 feat(assistant): refine voice controls and timed explanations
1d3abc0b feat(assistant): replace conversation picker with searchable history
7e605b84 fix(assistant): allow questions after stopping local generation
6b4ec74e perf(model-cache): delete GGUF shards in one cache pass
e812761b perf(history): bound expiry scans with the timestamp index
defb4e45 perf(chat-ui): batch conversation restoration and action updates
89b0095c perf(editor): reduce offline startup and storage overhead
1330b3c5 fix(editor): protect saves and recovery, refresh workspace checks
b73929c7 fix(ci): wait for delayed same-repository Pages checks
75028005 docs(seo): record sharing verification and refresh sitemap dates
13c6b409 fix(seo): complete sharing cards across public pages
b18f4a47 feat(branding): add product artwork for social previews
3a3dcae1 docs(readme): refine positioning and localized guide links
11029513 docs(readme): summarize all core product capabilities
aa68dd16 docs(readme): show live URL and remove redundant source CTA
bb0469e2 docs(readme): keep product titles text only
cce28a96 docs(readme): match title icon size to heading typography
546afb9f docs(readme): align title icons with heading text
717a66f8 docs(readme): highlight AI, WebMCP and PWA workflows
a2adda4a docs(readme): place compact icon beside each title
528de812 docs(readme): simplify product overview and consolidate technical reference
5daa1f1e test(assistant): enter chat before measuring sidebar layout
ee21ed1f fix(assistant): acknowledge inference cancellation immediately
5aad3283 fix(assistant): clarify compact speech input states
7a7b6772 fix(assistant): align sidebar content and compact model settings
8537747a docs(ai): replace demos with complete normal-speed walkthroughs
5635214d feat(chat-ui): animate request and streaming activity
270b5109 feat(branding): align app icons with the monochrome theme
29b1be1a fix(assistant): improve dictation and refresh AI demos
```

## Complete branch file delta versus inspected remote main

```text
M	.github/workflows/preview-smoke.yml
M	CHANGELOG.md
M	CLAUDE.md
M	Dockerfile
M	NOTICE
A	bin/build-social-image.mjs
A	bin/design-tokens.d.mts
A	bin/design-tokens.mjs
M	bin/font-license-sweep.mjs
A	bin/local-system-fonts.d.mts
A	bin/local-system-fonts.mjs
M	bin/offline-assets.mjs
M	bin/pages/render-home.mjs
M	bin/pages/render-page.mjs
A	bin/pages/social-image.mjs
A	bin/prepare-docker-wasm.mjs
A	bin/workspace-aliases.d.mts
A	bin/workspace-aliases.mjs
M	content/de/home.json
M	content/en/home.json
M	content/es/home.json
M	content/ja/home.json
M	content/ko/home.json
M	content/pt/home.json
M	content/zh-CN/home.json
A	courseware-online.html
A	deployment/WISEOFFICE.md
A	deployment/prepare-wiseoffice.py
A	docs/architecture/native-change-receipts.md
A	docs/architecture/structured-artifacts.md
A	docs/changelogs/2026-10-16-site-token-consistency.md
A	docs/changelogs/2026-10-17-open-retry-reporting.md
A	docs/deckwise-c-import-handoff-2026-11-10.md
M	docs/design-system.md
A	docs/developer/assistant-workspace.md
A	docs/developer/google-drive-connector.md
M	docs/embed-api.md
M	docs/embed-api.zh.md
A	docs/explorations/2026-10-13-follow-up-performance.md
A	docs/explorations/2026-10-13-save-recovery-test-workflow.md
A	docs/explorations/2026-10-13-social-sharing-cards.md
A	docs/explorations/2026-10-14-assistant-conversation-browser.md
A	docs/explorations/2026-10-14-assistant-stop-resume.md
A	docs/explorations/2026-10-14-assistant-workspace.md
A	docs/explorations/2026-10-14-docker-http-wasm-encoding.md
A	docs/explorations/2026-10-14-speech-toolbar-stability.md
A	docs/explorations/2026-10-15-assistant-continuity-acceptance.md
A	docs/explorations/2026-10-15-assistant-harness-review.md
A	docs/explorations/2026-10-15-editor-review-fixes.md
A	docs/explorations/2026-10-15-speech-first-permission.md
A	docs/explorations/2026-10-16-pr273-open-lifecycle-regressions.md
A	docs/explorations/2026-10-16-site-token-consistency.md
A	docs/explorations/2026-10-17-document-lifecycle-followups.md
A	docs/explorations/2026-10-17-open-retry-duplicate-report.md
A	docs/explorations/2026-10-17-pr279-ci-followup.md
A	docs/explorations/2026-10-17-project-followups.md
A	docs/explorations/2026-10-17-recovery-and-provider-request-safety.md
A	docs/explorations/2026-10-19-deckwise-selection-font-size.md
A	docs/explorations/2026-10-19-deckwise-selection-text-color.md
A	docs/explorations/2026-10-19-deckwise-text-box-position.md
A	docs/explorations/2026-10-19-docs-community-read-bridge.md
A	docs/explorations/2026-10-19-docs-online-host-isolation.md
A	docs/explorations/2026-10-19-docs-pinned-plugin-launch.md
A	docs/explorations/2026-10-20-deckwise-active-export-exclusion.md
A	docs/explorations/2026-10-20-deckwise-asset-image-insertion.md
A	docs/explorations/2026-10-20-deckwise-current-page-preview.md
A	docs/explorations/2026-10-20-deckwise-image-asset-bridge.md
A	docs/explorations/2026-10-20-deckwise-image-insertion.md
A	docs/explorations/2026-10-20-deckwise-loaded-font-family.md
A	docs/explorations/2026-10-20-deckwise-native-image-replacement.md
A	docs/explorations/2026-10-20-deckwise-native-textless-preview.md
A	docs/explorations/2026-10-20-deckwise-offscreen-page-preview.md
A	docs/explorations/2026-10-20-deckwise-pinned-image-asset.md
A	docs/explorations/2026-10-20-deckwise-png-chunk-integrity.md
A	docs/explorations/2026-10-20-deckwise-reviewed-image-transport.md
A	docs/explorations/2026-10-20-deckwise-source-page-render.md
A	docs/explorations/2026-10-20-deckwise-text-box-size.md
A	docs/explorations/2026-10-20-docs-error-read-invalidation.md
A	docs/explorations/2026-10-21-deckwise-inspected-fade-transport.md
A	docs/explorations/2026-10-21-deckwise-native-fade.md
A	docs/explorations/2026-10-21-deckwise-selected-text-bold.md
A	docs/explorations/2026-10-21-deckwise-selected-text-italic.md
A	docs/explorations/2026-10-24-docs-native-barrier-bridge.md
M	docs/fonts.zh.md
M	docs/media/ai-demos/README.md
M	docs/media/ai-demos/excel-ai.jpg
M	docs/media/ai-demos/excel-ai.mp4
M	docs/media/ai-demos/excel-ai.vtt
M	docs/media/ai-demos/pdf-ai.jpg
M	docs/media/ai-demos/pdf-ai.mp4
M	docs/media/ai-demos/pdf-ai.vtt
M	docs/media/ai-demos/pptx-ai.jpg
M	docs/media/ai-demos/pptx-ai.mp4
M	docs/media/ai-demos/pptx-ai.vtt
M	docs/media/ai-demos/word-ai.jpg
M	docs/media/ai-demos/word-ai.mp4
M	docs/media/ai-demos/word-ai.vtt
A	docs/reference.md
A	docs/superpowers/plans/2026-10-13-runtime-storage-optimizations.md
A	docs/superpowers/plans/2026-10-13-social-cards.md
A	docs/superpowers/plans/2026-10-14-assistant-workspace.md
A	docs/superpowers/plans/2026-10-15-assistant-continuity.md
A	docs/superpowers/plans/2026-10-15-assistant-followup.md
A	docs/superpowers/plans/2026-10-15-document-reliability.md
A	docs/superpowers/plans/2026-10-15-editor-review-fixes.md
A	docs/superpowers/plans/2026-10-16-ai-workspace-craft.md
A	docs/superpowers/plans/2026-10-16-ai-workspace-sources-artifacts.md
A	docs/superpowers/plans/2026-10-16-assistant-interaction-polish.md
A	docs/superpowers/plans/2026-10-16-assistant-workspace-consistency.md
A	docs/superpowers/plans/2026-10-16-site-token-consistency.md
A	docs/superpowers/specs/2026-10-13-follow-up-performance-design.md
A	docs/superpowers/specs/2026-10-13-runtime-storage-design.md
A	docs/superpowers/specs/2026-10-14-assistant-workspace-design.md
A	docs/superpowers/specs/2026-10-15-assistant-continuity-design.md
A	docs/superpowers/specs/2026-10-15-assistant-followup-design.md
A	docs/superpowers/specs/2026-10-16-ai-workspace-craft-design.md
A	docs/superpowers/specs/2026-10-16-assistant-workspace-consistency-design.md
A	docs/superpowers/specs/2026-10-16-site-token-consistency-design.md
M	docs/test-matrix.md
M	editor.html
A	functions/api/connectors/google/[[path]].ts
M	history.html
M	index.html
M	index.ts
A	lib/agent-plugin/artifacts/edit-proposal.ts
A	lib/agent-plugin/artifacts/export.ts
A	lib/agent-plugin/artifacts/payload.ts
A	lib/agent-plugin/artifacts/store.ts
A	lib/agent-plugin/completed-paste.ts
A	lib/agent-plugin/connectors/document-reference-readers.ts
A	lib/agent-plugin/connectors/pdf-reference-reader.ts
A	lib/agent-plugin/connectors/reference-lifetime.ts
A	lib/agent-plugin/connectors/reference-search.ts
A	lib/agent-plugin/connectors/reference-sources.ts
M	lib/agent-plugin/document-agent-tools.ts
M	lib/agent-plugin/document-read-answer.ts
M	lib/agent-plugin/document-task.ts
M	lib/agent-plugin/document-tool-action.ts
M	lib/agent-plugin/document-tool-plan.ts
M	lib/agent-plugin/document-tool-sequence.ts
M	lib/agent-plugin/editor-bridge.ts
M	lib/agent-plugin/native-paste.ts
M	lib/agent-plugin/office-tools.ts
M	lib/agent-plugin/reviewed-action.ts
A	lib/agent-plugin/slide-fade.ts
A	lib/agent-plugin/slide-font-family.ts
A	lib/agent-plugin/slide-font-size.ts
A	lib/agent-plugin/slide-image-asset.ts
A	lib/agent-plugin/slide-image-insertion.ts
A	lib/agent-plugin/slide-image-png.ts
A	lib/agent-plugin/slide-image-position.ts
A	lib/agent-plugin/slide-image-replacement.ts
A	lib/agent-plugin/slide-image-size.ts
A	lib/agent-plugin/slide-image-transform.ts
A	lib/agent-plugin/slide-object-animation.ts
A	lib/agent-plugin/slide-text-bold.ts
A	lib/agent-plugin/slide-text-box-position.ts
A	lib/agent-plugin/slide-text-box-size.ts
A	lib/agent-plugin/slide-text-color.ts
A	lib/agent-plugin/slide-text-italic.ts
A	lib/agent-plugin/slide-text-selection.ts
M	lib/agent-plugin/tools.ts
M	lib/agent-plugin/ui/action-preview.ts
A	lib/agent-plugin/ui/artifact-presentation.ts
A	lib/agent-plugin/ui/artifact-request.ts
A	lib/agent-plugin/ui/artifact-workflow.ts
A	lib/agent-plugin/ui/artifacts-view.ts
A	lib/agent-plugin/ui/composer-selection.ts
A	lib/agent-plugin/ui/controller-bindings.ts
M	lib/agent-plugin/ui/controller.ts
M	lib/agent-plugin/ui/conversation-db.ts
A	lib/agent-plugin/ui/direct-document-request.ts
M	lib/agent-plugin/ui/document-draft.ts
A	lib/agent-plugin/ui/endpoint-lifecycle.ts
M	lib/agent-plugin/ui/history-controls.ts
A	lib/agent-plugin/ui/history-idle.ts
A	lib/agent-plugin/ui/history-list.ts
A	lib/agent-plugin/ui/history-search-model.ts
A	lib/agent-plugin/ui/history-search.ts
A	lib/agent-plugin/ui/history-search.worker.ts
A	lib/agent-plugin/ui/host-read-request.ts
A	lib/agent-plugin/ui/model-cache-settings.ts
A	lib/agent-plugin/ui/model-configuration-form.ts
A	lib/agent-plugin/ui/panel-context.ts
A	lib/agent-plugin/ui/panel-workspace.ts
M	lib/agent-plugin/ui/panel.ts
M	lib/agent-plugin/ui/presentation.ts
M	lib/agent-plugin/ui/request-context.ts
A	lib/agent-plugin/ui/request-error.ts
M	lib/agent-plugin/ui/request-intent.ts
A	lib/agent-plugin/ui/request-queue-controls.ts
A	lib/agent-plugin/ui/request-queue-view.ts
A	lib/agent-plugin/ui/request-queue.ts
A	lib/agent-plugin/ui/request-routing.ts
A	lib/agent-plugin/ui/request-scope.ts
A	lib/agent-plugin/ui/request-timing.ts
A	lib/agent-plugin/ui/result-actions.ts
A	lib/agent-plugin/ui/review-settings.ts
A	lib/agent-plugin/ui/search-highlight.ts
M	lib/agent-plugin/ui/sessions.ts
A	lib/agent-plugin/ui/settings-view.ts
M	lib/agent-plugin/ui/sidebar-entry.ts
A	lib/agent-plugin/ui/source-citations.ts
A	lib/agent-plugin/ui/source-evidence.ts
A	lib/agent-plugin/ui/source-preview.ts
A	lib/agent-plugin/ui/sources-view.ts
M	lib/agent-plugin/ui/speech-input.ts
M	lib/agent-plugin/ui/storage.ts
A	lib/agent-plugin/ui/task-entry.ts
A	lib/agent-plugin/ui/task-receipt.ts
A	lib/agent-plugin/ui/view-header.ts
A	lib/agent-plugin/ui/workspace-labels.ts
A	lib/agent-plugin/ui/workspace-navigation.ts
M	lib/agent-plugin/word-paste-history.ts
M	lib/agent-plugin/word-paste.ts
A	lib/assistant-offline.ts
M	lib/confirm-dialog.ts
M	lib/converter.ts
A	lib/courseware-online-bootstrap.ts
A	lib/courseware-online-entry.ts
A	lib/deckwise-editor-bridge.ts
A	lib/deckwise-object-selection.ts
A	lib/deckwise-page-lineage.ts
A	lib/deckwise-slide-preview.ts
A	lib/docs-community-barrier.ts
A	lib/docs-community-plugin-entry.ts
A	lib/docs-community-plugin-host.ts
A	lib/docs-community-plugin.ts
A	lib/docs-native-barrier-9.3.0-138.ts
A	lib/docs-native-barrier-extension.ts
A	lib/docs-online-editor.ts
A	lib/document-open.ts
M	lib/document.ts
M	lib/embed-api.ts
M	lib/events.ts
M	lib/history-page.ts
M	lib/history/autosave.ts
M	lib/history/db.ts
M	lib/history/index.ts
M	lib/history/session.ts
M	lib/history/store.ts
M	lib/onlyoffice-editor.ts
A	lib/onlyoffice/guards/east-asian-punctuation.ts
A	lib/onlyoffice/guards/east-asian-vertical.ts
A	lib/onlyoffice/guards/image-color-key.ts
A	lib/onlyoffice/guards/image-soft-edge.ts
A	lib/onlyoffice/guards/mixed-script-line-metrics.ts
M	lib/onlyoffice/guards/serverless-save.ts
M	lib/onlyoffice/guards/x2t-worker.ts
M	lib/onlyoffice/iframe-guards.ts
M	lib/onlyoffice/open-failure.ts
M	lib/onlyoffice/save-point.ts
M	lib/onlyoffice/save-stream.ts
A	lib/remote-document.ts
A	lib/remote-file-name.ts
M	lib/save-target.ts
M	lib/web-mcp.ts
M	package.json
M	packages/agent-core/CLAUDE.md
M	packages/agent-core/README.md
M	packages/agent-core/src/context-budget.ts
M	packages/agent-core/src/llm/anthropic.ts
M	packages/agent-core/src/llm/gemini.ts
M	packages/agent-core/src/llm/local.ts
M	packages/agent-core/src/llm/loopback.ts
M	packages/agent-core/src/llm/model-cache.ts
M	packages/agent-core/src/llm/ollama.ts
A	packages/agent-core/src/llm/openai-stream.ts
M	packages/agent-core/src/llm/openai.ts
M	packages/agent-core/src/llm/prompt.ts
A	packages/agent-core/src/llm/provider-http.ts
A	packages/agent-core/src/llm/provider-stream.ts
A	packages/agent-core/src/llm/runtime-assets.ts
M	packages/agent-core/src/llm/types.ts
M	packages/agent-core/src/llm/webllm.ts
M	packages/agent-core/src/llm/wllama-compat-assets.ts
M	packages/agent-core/src/llm/wllama-cpu-runtime.ts
M	packages/agent-core/src/llm/wllama.ts
M	packages/agent-core/src/llm/writing-task.ts
M	packages/agent-core/src/runtime.ts
M	packages/chat-ui/CLAUDE.md
M	packages/chat-ui/README.md
M	packages/chat-ui/src/chat-view.ts
M	packages/chat-ui/src/index.ts
M	packages/chat-ui/src/markdown.ts
A	packages/chat-ui/src/streaming-markdown.ts
M	packages/chat-ui/src/styles.ts
M	packages/chat-ui/src/types.ts
M	packages/converter/src/document-converter.ts
M	packages/converter/src/docx-zip.ts
M	packages/converter/src/file-meta.ts
A	packages/converter/src/pptx-theme-fonts.ts
M	packages/shared/CLAUDE.md
M	packages/shared/src/i18n/messages/de.ts
M	packages/shared/src/i18n/messages/en.ts
M	packages/shared/src/i18n/messages/es.ts
M	packages/shared/src/i18n/messages/ja.ts
M	packages/shared/src/i18n/messages/ko.ts
M	packages/shared/src/i18n/messages/pt.ts
M	packages/shared/src/i18n/messages/zh-CN.ts
M	packages/shared/src/i18n/types.ts
M	playwright.config.ts
M	playwright.docker.config.ts
M	pnpm-lock.yaml
M	public/404.html
M	public/_headers
A	public/_routes.json
A	public/app-tokens.css
A	public/docs-community-plugin/config.json
A	public/docs-community-plugin/index.html
A	public/docs-online-host.html
M	public/embed-demo.html
M	public/home.css
M	public/icons/document-180.png
M	public/icons/document-192.png
M	public/icons/document-32.png
M	public/icons/document-512.png
M	public/icons/document-dark-32.png
M	public/icons/document-dark.svg
M	public/icons/document-light-32.png
M	public/icons/document-light.svg
M	public/icons/document-maskable-512.png
M	public/icons/document-maskable.svg
M	public/icons/document.svg
M	public/landing-prefetch.js
M	public/landing.css
M	public/sdkjs/common/AllFonts.js
M	public/sdkjs/common/wasm/x2t/x2t_helper.js
M	public/sitemap.xml
A	public/social/README.md
A	public/social/document-editor-github-v1.png
A	public/social/document-editor-github-v1.svg
A	public/social/document-editor-v1.png
A	public/social/document-editor-v1.svg
M	public/sw.js
M	public/web-apps/apps/api/documents/api.js
M	readme.de.md
M	readme.es.md
M	readme.fa.md
M	readme.ja.md
M	readme.ko.md
M	readme.md
M	readme.pt.md
M	readme.zh.md
M	styles/assistant.css
M	styles/base.css
M	styles/confirm-dialog.css
M	styles/history.css
M	sws.toml
A	test/e2e/agent-artifact-revision.spec.ts
A	test/e2e/agent-cloud-streaming.spec.ts
M	test/e2e/agent-composer-ux.spec.ts
M	test/e2e/agent-content-review.spec.ts
A	test/e2e/agent-context-events.spec.ts
A	test/e2e/agent-dictation.spec.ts
A	test/e2e/agent-document-references.spec.ts
A	test/e2e/agent-error-recovery.spec.ts
A	test/e2e/agent-history-rendering.spec.ts
M	test/e2e/agent-imported-formats.spec.ts
A	test/e2e/agent-process-feedback.spec.ts
A	test/e2e/agent-request-queue.spec.ts
A	test/e2e/agent-result-hierarchy.spec.ts
A	test/e2e/agent-sidebar-layout.spec.ts
A	test/e2e/agent-sources-artifacts.spec.ts
A	test/e2e/agent-stop-resume.spec.ts
M	test/e2e/agent-task-loop.spec.ts
A	test/e2e/agent-workspace-consistency.spec.ts
A	test/e2e/agent-workspace.spec.ts
M	test/e2e/agent-write-tools.spec.ts
A	test/e2e/assistant-offline-assets.spec.ts
A	test/e2e/deckwise-formats.spec.ts
M	test/e2e/docker-cache-headers.spec.ts
A	test/e2e/editor-csp.spec.ts
M	test/e2e/font-substitution.spec.ts
M	test/e2e/history-page.spec.ts
A	test/e2e/history-quota.spec.ts
M	test/e2e/landing-prefetch.spec.ts
M	test/e2e/main-site.spec.ts
M	test/e2e/offline-seam.spec.ts
A	test/e2e/open-replacement.spec.ts
M	test/e2e/open-retry.spec.ts
M	test/e2e/pdf-agent-contract.spec.ts
A	test/e2e/remote-open.spec.ts
M	test/e2e/save-to-file.spec.ts
A	test/e2e/site-design-consistency.spec.ts
A	test/e2e/social-sharing.spec.ts
M	test/e2e/theme-presentation.spec.ts
M	test/e2e/webmcp.spec.ts
A	test/fixtures/helpers/reference-files.ts
M	test/unit/agent-action-preview.test.ts
A	test/unit/agent-artifact-edit-proposal.test.ts
A	test/unit/agent-artifact-payload.test.ts
A	test/unit/agent-artifact-presentation.test.ts
A	test/unit/agent-artifact-request.test.ts
A	test/unit/agent-artifact-workflow.test.ts
A	test/unit/agent-artifacts-view.test.ts
A	test/unit/agent-artifacts.test.ts
A	test/unit/agent-completed-paste.test.ts
M	test/unit/agent-context-budget.test.ts
A	test/unit/agent-design-tokens.test.ts
A	test/unit/agent-direct-request.test.ts
A	test/unit/agent-document-reference-readers.test.ts
M	test/unit/agent-editor-bridge.test.ts
M	test/unit/agent-entry.test.ts
A	test/unit/agent-google-connector.test.ts
M	test/unit/agent-history-controls.test.ts
A	test/unit/agent-history-list.test.ts
A	test/unit/agent-history-search-worker.test.ts
A	test/unit/agent-history-search.test.ts
A	test/unit/agent-host-read-request.test.ts
A	test/unit/agent-llm-anthropic-lifetime.test.ts
M	test/unit/agent-llm-anthropic.test.ts
M	test/unit/agent-llm-gemini.test.ts
M	test/unit/agent-llm-ollama.test.ts
M	test/unit/agent-llm-openai.test.ts
M	test/unit/agent-llm-webllm.test.ts
M	test/unit/agent-loopback.test.ts
M	test/unit/agent-model-cache.test.ts
M	test/unit/agent-native-paste.test.ts
M	test/unit/agent-office-tools.test.ts
M	test/unit/agent-panel-loading.test.ts
A	test/unit/agent-panel-owners.test.ts
M	test/unit/agent-panel.test.ts
A	test/unit/agent-pdf-reference-reader.test.ts
M	test/unit/agent-presentation.test.ts
A	test/unit/agent-provider-bounds.test.ts
A	test/unit/agent-provider-http.test.ts
A	test/unit/agent-provider-streaming.test.ts
A	test/unit/agent-reference-lifetime.test.ts
A	test/unit/agent-reference-sources.test.ts
M	test/unit/agent-request-context.test.ts
M	test/unit/agent-request-intent.test.ts
A	test/unit/agent-request-queue-controls.test.ts
A	test/unit/agent-request-queue-view.test.ts
A	test/unit/agent-request-queue.test.ts
A	test/unit/agent-request-scope.test.ts
A	test/unit/agent-result-actions.test.ts
M	test/unit/agent-runtime.test.ts
M	test/unit/agent-sessions.test.ts
A	test/unit/agent-settings-view.test.ts
M	test/unit/agent-sidebar-entry.test.ts
A	test/unit/agent-source-discovery.test.ts
A	test/unit/agent-source-routes.test.ts
A	test/unit/agent-sources-view.test.ts
M	test/unit/agent-speech-input.test.ts
A	test/unit/agent-task-entry.test.ts
A	test/unit/agent-task-receipt.test.ts
M	test/unit/agent-ui-controller.test.ts
A	test/unit/agent-view-header.test.ts
M	test/unit/agent-word-paste.test.ts
A	test/unit/agent-workspace-navigation.test.ts
M	test/unit/agent-write-reply.test.ts
M	test/unit/agent-writing-task.test.ts
A	test/unit/assistant-offline.test.ts
A	test/unit/assistant-request-timing.test.ts
A	test/unit/chat-source-presentation.test.ts
M	test/unit/chat-view-layout.test.ts
A	test/unit/chat-view-navigation.test.ts
A	test/unit/chat-view-performance.test.ts
A	test/unit/chat-view-presentation-timing.test.ts
A	test/unit/chat-view-progress.test.ts
A	test/unit/chat-view-queue.test.ts
A	test/unit/chat-view-results.test.ts
A	test/unit/chat-view-save-feedback.test.ts
A	test/unit/chat-view-scroll.test.ts
A	test/unit/chat-view-stop.test.ts
A	test/unit/chat-view-streaming-markdown.test.ts
M	test/unit/chat-view.test.ts
A	test/unit/composer-selection.test.ts
M	test/unit/conversation-db.test.ts
A	test/unit/converter-open-currency.test.ts
A	test/unit/courseware-online-entry.test.ts
A	test/unit/deckwise-editor-bridge.test.ts
M	test/unit/deckwise-embed.test.ts
A	test/unit/deckwise-fade-bridge.test.ts
A	test/unit/deckwise-object-selection.test.ts
A	test/unit/deckwise-page-lineage.test.ts
A	test/unit/deckwise-slide-preview.test.ts
M	test/unit/design-contract.test.ts
A	test/unit/docker-wasm.test.ts
A	test/unit/docs-community-plugin-host.test.ts
A	test/unit/docs-community-plugin.test.ts
A	test/unit/docs-native-barrier-extension.test.ts
A	test/unit/docs-online-editor.test.ts
A	test/unit/docs-online-plugin.test.ts
M	test/unit/document-agent-tools.test.ts
A	test/unit/document-entry-contracts.test.ts
A	test/unit/document-open.test.ts
M	test/unit/document-task.test.ts
M	test/unit/document-tool-action.test.ts
M	test/unit/docx-zip.test.ts
A	test/unit/editor-api-startup.test.ts
M	test/unit/embed-api.test.ts
A	test/unit/endpoint-lifecycle.test.ts
M	test/unit/font-catalog-licensing.test.ts
A	test/unit/history-autosave-start.test.ts
M	test/unit/history-autosave.test.ts
A	test/unit/history-bounded-reads.test.ts
A	test/unit/history-expiry-index.test.ts
A	test/unit/history-handle-cleanup.test.ts
A	test/unit/history-session-lifecycle.test.ts
M	test/unit/history-store.test.ts
M	test/unit/hosting-contract.test.ts
M	test/unit/i18n-locales.test.ts
M	test/unit/landing-pages.test.ts
M	test/unit/landing-prefetch.test.ts
A	test/unit/lib/webllm-fixture.ts
M	test/unit/local-generation-settings.test.ts
A	test/unit/local-runtime-assets.test.ts
M	test/unit/local-source-open.test.ts
A	test/unit/local-system-fonts.test.ts
A	test/unit/office-east-asian-punctuation.test.ts
A	test/unit/office-east-asian-vertical.test.ts
A	test/unit/office-image-color-key.test.ts
A	test/unit/office-image-soft-edge.test.ts
A	test/unit/office-mixed-script-line-metrics.test.ts
M	test/unit/offline-assets-build.test.ts
M	test/unit/onlyoffice-editor.test.ts
M	test/unit/onlyoffice-open-failure.test.ts
A	test/unit/open-ownership.test.ts
A	test/unit/panel-context.test.ts
M	test/unit/preview-wait.test.ts
A	test/unit/remote-document.test.ts
M	test/unit/save-point.test.ts
A	test/unit/save-stream-lifecycle.test.ts
M	test/unit/save-stream-origin.test.ts
M	test/unit/save-target.test.ts
A	test/unit/site-token-consistency.test.ts
A	test/unit/slide-fade.test.ts
A	test/unit/slide-font-family.test.ts
A	test/unit/slide-font-size.test.ts
A	test/unit/slide-image-asset.test.ts
A	test/unit/slide-image-insertion.test.ts
A	test/unit/slide-image-replacement.test.ts
A	test/unit/slide-object-animation.test.ts
A	test/unit/slide-text-bold.test.ts
A	test/unit/slide-text-box-position.test.ts
A	test/unit/slide-text-color.test.ts
A	test/unit/slide-text-italic.test.ts
A	test/unit/social-cards.test.ts
M	test/unit/styles-contract.test.ts
A	test/unit/sw-assistant-reuse.test.ts
M	test/unit/sw-live-core.test.ts
M	test/unit/sw-offline-editor.test.ts
M	test/unit/sw-routing.test.ts
M	test/unit/web-mcp.test.ts
A	test/unit/webllm-stop-resume.test.ts
M	test/unit/wllama-bounded-file-read.test.ts
M	test/unit/wllama-compat-assets.test.ts
M	test/unit/wllama-cpu-loader.test.ts
M	test/unit/wllama-loading-disposal.test.ts
M	test/unit/wllama-provider.test.ts
M	test/unit/workflow-contract.test.ts
A	test/unit/workspace-aliases.test.ts
M	test/unit/x2t-worker-input.test.ts
M	vite.config.ts
A	vite.courseware-online.config.ts
M	vitest.config.ts
```
