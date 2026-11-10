# Native teaching in the document

Branch: `feat/native-teaching-implementation`, based on local Part A commit
`285f04c8df2336edd6ca6fc067d00a9627ac1b2a`.

The live ONLYOFFICE model now owns native page teaching. The DeckWise-bound
`document:native-teaching` commands and the embedded sidebar share local reads,
revision-checked changes, native history, locks, confirmations, audio adoption
and media preview. New activation reads already loaded notes. Native files never
fall back to an editable Host teaching mirror.

`lib/deckwise-native-teaching.ts` stores metadata and object/page UUIDs in one
presentation-related custom XML part, and each adopted audio/video payload in a
separate immutable media part. Sentence edits therefore do not rewrite audio.
Adoption verifies hashes and changes the action/media link in one native history
action. Native Undo/Redo restores it. Failed compound writes roll back parts.

The SDK import guard captures original package shape IDs before ID regeneration.
Rename, slide reorder and Undo references preserve logical identity. Slide
copies receive fresh IDs and remapped references while reusing embedded media.
Export updates current package IDs without creating another user Undo point.
Missing or changed animation links reject playback until explicitly repaired.
Cross-document teaching clipboard transfer is not implemented.

The sidebar refreshes on teaching revision changes even when native SDK edit
generation stays unchanged. Hidden panels suspend optional listeners/polling.
Native preview uses live objects, including groups, and locally verified adopted
audio. The native checkpoint editor retains text choices and reference answers.
Quiz preview accepts empty optional explanation text and key-point lists, as the
saved-course contract does. Closed-question answer keys still undergo response
validation. Whole-lesson playback skips pages with no teaching actions instead
of constructing an invalid empty timeline.
Native attention and speech-animation targets can also be descendants of groups;
the exact object must still belong to the requested slide. Object lookup stays
within that slide when validating its chooser or preview. Supported finite native
automatic/parallel effects and delays finish in the SDK before narration begins.
For native pages with speech animation cues, the first preview entry automatically
uses native slideshow playback rather than reporting that canvas playback cannot
render those animations.

Teaching metadata lookup uses the SDK namespace index. Presence checks do not
serialize XML, and reading/editing teaching or preparing save metadata never
serializes the independent media parts. The real-fixture E2E checks all six slide
switches, including the empty third page, against a 1,500 ms navigation budget.

The paired Desktop implementation owns private generation receipts, media
service authorization, verified normal saves and immutable fixed builds. It
preserves group aliases for attention and verified finite native animations;
these do not create a second editable teaching model. This feature requires the
matching Desktop/editor build.

Validation uses an explicitly copied real six-page Part A course: 24 teaching
nodes, 22 adopted audio files (808,350 decoded bytes), Quiz and assessment pages.
The staged Desktop-transformed editor test opens it, decodes every embedded audio,
edits/undoes/redoes teaching, hears actual local audio, associates group highlight
and spotlight, saves and reopens with UUIDs and assessment nodes intact.

Cross-repository details and final evidence:
`deckwise-desktop/docs/design/2026-11-09-native-teaching/`. General old-data
migration is excluded by the user's request. Windows/macOS power-loss and named
external PowerPoint/WPS save compatibility are not qualified by these Linux tests.
