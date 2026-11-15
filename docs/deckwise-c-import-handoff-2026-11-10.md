# DeckWise C import: embedded editor handoff

Date: 2026-11-10. Internal remote: `http://192.168.15.230/ligang/document.git`.
Branch: `zpj_deck`. Paired client: `ligang/deckwise`, branch `zpj_deck`.
The client handoff is `docs/project/2026-11-10-c-import-integration-handoff.md`.
This is an integration candidate, not full PowerPoint/web fidelity acceptance.

## Baselines and integration

Remote main inspected: `3cc37956d614ce1985cf7cfd0401e2c9e23f58ba`.
Starting local HEAD: `632c5201` (`feat(wiseoffice): integrate courseware editor
workstreams`). That branch already had 201 commits beyond remote main, including
upstream/history/assistant work. Do not classify the full ancestry as new C work.
Compare the destination before merge/cherry-pick. The final new delta is listed
in `deckwise-c-import-file-inventory-2026-11-10.md`.

`632c5201` supplies native edit commands, page lineage, selection, current/offscreen
page rendering, textless source capture, image handling, save barriers and the
DeckWise embed bridge. See `docs/embed-api.md`, `docs/embed-api.zh.md` and the
`docs/plans/` history. Client ownership/project/session/edit-generation checks
must survive integration. The new object-selection adapter is read-only.

## Latest local modifications

| Files                                                                                                                        | Behavior / reason                                                                                                                                                        |
| ---------------------------------------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| `bin/font-license-sweep.mjs`, `public/sdkjs/common/AllFonts.js`, font licensing/e2e tests                                    | Repair Calibri -> Carlito regular/italic/bold/bold-italic aliases, including previously swept catalogs; mismatched style indices previously produced wrong glyphs        |
| `bin/local-system-fonts.mjs`, `.d.mts`, `vite.config.ts`, `test/unit/local-system-fonts.test.ts`                             | Runtime-only installed Windows Calibri, SimSun, SimHei, Microsoft YaHei, KaiTi and FangSong mappings, with aliases and guarded loopback access                           |
| `packages/converter/src/pptx-theme-fonts.ts`, `docx-zip.ts`, `lib/converter.ts`                                              | Resolve empty theme East Asian fonts from Hans in temporary conversion input for eligible simplified-Chinese source; preserve original file and slide XML                |
| `lib/onlyoffice/guards/east-asian-punctuation.ts`                                                                            | Select Chinese smart-quote font slot using main run language before first layout; Latin text retains Latin behavior                                                      |
| `guards/mixed-script-line-metrics.ts`                                                                                        | Prevent Chinese automatic slide spacing from combining incompatible ascender/descender maxima across fonts; keep inline-object height and exclude unsupported text modes |
| `guards/east-asian-vertical.ts`                                                                                              | Draw supported eaVert Han glyphs upright by inverse rotation around the measured font-cell center; preserve source text/geometry/direction                               |
| `guards/image-color-key.ts`                                                                                                  | Adapt MAIC exact same-sRGB clrChange-to-transparent rendering through a temporary cached canvas; source image bytes/effects remain unchanged                             |
| `guards/image-soft-edge.ts`                                                                                                  | Draw one softEdge on an uncropped/untiled ellipse without visible outline; cached inward alpha feather uses the source EMU radius                                        |
| `guards/x2t-worker.ts`, `lib/onlyoffice/iframe-guards.ts`                                                                    | Install drawing/font adapters before initial conversion/layout and on iframe initialization; late-only installation can leave incorrect cached widths                    |
| `lib/deckwise-object-selection.ts`, `lib/deckwise-editor-bridge.ts`                                                          | Read native selected object identity through the bound embed bridge for Host teaching attention mapping                                                                  |
| `test/unit/office-*.test.ts`, selection/converter/embed/font tests, `test/e2e/font-substitution.spec.ts`, `docs/fonts.zh.md` | Regression and deployment details                                                                                                                                        |

The soft-edge algorithm is a finite visual approximation, not PowerPoint's exact
pixel curve. Multi-effect, crop/tile/outline, unreadable or oversized cases fall
back to the existing renderer. Color-key matching is exact, so near-white JPEG
fringes remain. Vertical punctuation/combining clusters are not covered by the
Han fix. Renderer state and caches must not corrupt subsequent images/glyphs.

## Original file versus internal conversion

The editor receives the original PPTX, but ONLYOFFICE uses x2t/internal document
structures. Temporary preprocessing fixes eligible theme import metadata;
drawing guards adapt rendering. These do not intentionally rewrite original
slide elements. An editor save still round-trips Office XML and must be tested
separately; opening the original is not a proof of full native fidelity.

## Build and deployment

Use pinned Node/pnpm and install the lockfile dependencies. Run the normal build
(`pnpm run build`, shell required), `pnpm exec tsc --noEmit` and targeted Vitest.
Deploy the complete matching dist, not only a guard source file; see
`deployment/WISEOFFICE.md` for the existing `/wiseoffice/` deployment.

For Windows local installed fonts, set `DECKWISE_SYSTEM_FONTS=1` **before starting
the Vite editor service**. Example local service: `pnpm exec vite preview --host
127.0.0.1 --port 4371 --strictPort`. Pair the Desktop's `DECKWISE_EDITOR_URL` with
that origin. Default web/static deployments keep the redistributable font
catalog. This runtime plugin is not a static font bundle: restart its service
after changing its catalog. Do not ship Windows font files in public/, dist/,
an offline package or exported PPTX. Preserve origin/loopback guards.

## Verification and gaps

Fresh handoff run: TypeScript passed. Eleven targeted test files initially gave
186 passes and 3 failures (timeout/state-dependent results during concurrent
suites). A serial rerun of the two affected files, embed API and font licensing,
passed all 118 cases. The other nine files passed 71 cases in the first run.
This is not a full editor suite or final packaged desktop acceptance.

Prior read-only real-deck checks rendered all 28 pages against PowerPoint;
191 top-level shape geometries were compared. Original file hash:
`e54e5cdb20795679c0092f34bc6507358000cc2eb40396b7a2153b7cf5737901`.
Vertical changes affected pages 6/7 only; soft-edge changed page 11 only. Other
pages were checked bytewise against the preceding editor render. These checks
do not certify all text layout, grouped children, animations or saved-file fidelity.

Still open: page 13 hanging punctuation (SDK drops some paragraph flags), page 5
synthetic-bold advance/wrapping, vertical punctuation, other image/shape effects
on pages 14-18/25, and complex native animations in fixed playback. Missing
external source videos are asset omissions, not recoverable from the PPTX alone.
The client validation report lists separate runtime/sidebar regressions; no full
restoration/publish/robot acceptance is claimed.

No original classroom deck, paid audio, runtime profile or proprietary font
binary is included in the handoff. Generated files with no substantive diff
were left out of the commit.
