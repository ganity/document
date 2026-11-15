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
automatic/parallel effects and delays retain Part A's concurrent native playback;
speech-linked steps still complete from actual SDK state.
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

## Native speech timing correction, 2026-11-10

Native audio adoption and page writes now materialize missing speech animation
presets in the SDK timing tree, preserving unrelated effects. A detached tree is
verified before publishing it with the teaching/media change in the same history
action. Registered SDK objects must remain available to native Undo/Redo. A
repeated unchanged adoption/write does not duplicate effects.

The generated-course browser regression checks actual pulse and manually added
teeter timing, native Undo/Redo, slideshow playback and PPTX reopen. The paired
Desktop fixed build claims verified persisted speech mirrors exactly once.
Scene planning narration remains generation evidence until the structured native
narration step creates final speech. Live notes take precedence as source text.

The initial Desktop parity audit recorded missing vision/whole-course planning
and assessment/media editor surfaces. The continuation below addresses those
four gaps; the earlier successful playback alone did not establish parity.

## Native authoring parity continuation, 2026-11-10

`narration-visual` renders the live native slide without a teaching mutation.
The paired Desktop feeds these PNGs and live UUID evidence into the shared visual
semantic validator and whole-course narration planner. Native writes validate
and retain object roles/aliases/target eligibility in `objectSemantics`.

`assessment-context` projects native Quiz/discussion pages into the existing
assessment editor. `assessment-write` checks revision, page order and locks, then
commits question/answer/reference changes and rich Quiz image metadata to the
carrier in one native history action. Changed question readings invalidate old
audio and cues. Discussion editor prompt/label fields are separate from strict
playback actions; page copies preserve these details and remap identities.
The projection is transient, with no editable Host classroom authority.

Media listing returns embedded assets plus a native-service request so Desktop
can append pending music/video journal entries. Explicit adoption uses the
existing verified native media service. Narration audio is not reusable music.

The real browser parity regression renders a PNG, edits an image question,
answer and reference, undoes/redoes, saves a copied PPTX and reopens it with the
complete question intact. Native-teaching/preview tests and TypeScript checks
pass. The corresponding Desktop fixed export archives carrier Quiz images and
restores them into player content. Live providers and external Office save
compatibility are not qualified by these local tests.

### Final assessment parity corrections

Native read views expose read-only state before provider submissions. First-edit
Undo on copied slides falls back to an immutable copy seed. Practice, PBL and
discussion project as interaction actions on slides; a unique live question shape
provides the existing question editor with a stable text binding.

Quiz visual projection is shared with the Desktop legacy editor. It preloads
validated PNGs, fences the current Office history, then commits the question,
options, images and carrier in one native action. Bound discussion question edits
preserve paragraph and run formatting. The browser regression verifies visible
text and image count through Undo/Redo and save/reopen, as well as practice text
and its original font size. Desktop generation also packages full Quiz data for
fixed export before any manual assessment edit.

## Replacing persisted speech motions, 2026-11-11

Replacing or removing a speech cue now reconciles its previous native timing
mirror in the same atomic teaching write. Matching includes the logical target,
preset and duration; unrelated native effects and explicitly linked imported
animations are retained. A failed timing/carrier commit restores the old tree.
This prevents an obsolete zoom entrance from surviving a correction to fade and
conflicting with the fixed player after save/reopen. Existing audio parts and
assessment pages are untouched. The real-fixture regression
`deckwise-native-speech-motion-replacement.spec.ts` checks the SDK tree, saved
PPTX, reopen and subsequent removal; its input is an explicitly supplied copy of
the real generated lesson, not a remote model or TTS mock.

### Native readiness without sidebar visibility

`document:teaching-enable` acknowledges `document:teaching-ready` after the bound
SDK context and native teaching initialization succeed. The sidebar no longer
owns this acknowledgement: a closed panel or hidden host does not need sidebar
DOM to execute source-bound native tools. Failed initialization emits the existing
command error and must not emit readiness. See `deckwise-teaching-ready.test.ts`.

## Grouped Quiz picture recovery (2026-11-15)

The native assessment projector resolves exported object names recursively and
matches the question anchor without export-inserted line breaks. A complete
assessment group is rebuilt from the canonical Quiz projection at its original
stacking position. A group containing any unowned object is rejected. Native
history, slide identity and nested object arrays are rechecked after image
preloading; transaction rollback restores the original objects. This also repairs
already-saved grouped Quiz files when adopting an existing image task.

Regression coverage: `test/unit/slide-quiz-projection.test.ts` and
`test/e2e/deckwise-grouped-quiz-recovery.spec.ts`. The latter takes a real generated
fixture through projection, native save and reopen and checks the picture plus
unchanged teaching actions. It uses a local PNG and makes no paid request.
