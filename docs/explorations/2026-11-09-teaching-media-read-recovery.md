# Teaching media read recovery

The native teaching panel previously reported every rejected voice-model or
voice-task read as "voice service unavailable". That conflated expired teaching
bindings, authentication failures and remote model failures, without showing the
underlying error or attempting recovery.

The visible failure was reproduced in the live native panel. Refreshing then
successfully read models, voice tasks and the media library. The original rejected
response had not been logged, so its underlying exception is not established.
Do not interpret the recovery change as proof that the remote service never fails.

Read failures now log the operation and reason and retry once through a fresh
teaching view. Recovery does not submit media jobs and is skipped while editing
or playing. A persistent error names the failed catalog and original reason.
Saved narration and audio remain available.

Unsupported native-animation notices are independent of these reads. Generated
group animations, unsupported presets, delays and ambiguous effects can remain
unavailable to the linking form. Its warning now occupies a separate row instead
of shrinking the animation name into a narrow vertical column.

Validation covers transient and persistent reads and edits made during a pending
read. The Desktop Electron fixture probe injects one failed model read and checks
automatic recovery, no paid operations, local playback, save/reopen and unchanged
source bytes. Existing locale-completeness checks still flag older untranslated
animation keys in five locales; the new catalog messages are translated.

Native slideshow admission and legacy speech-step binding now have separate
checks. Delays, automatic starts and native presets outside the legacy linking
subset do not reject an otherwise playable page. Narration starts when the
native canvas/transition is ready, while unbound animation remains under the
native engine. Idle click steps advance without changing slides. Linked motion
continues to use explicit ownership and temporary schedules, restored on exit.
Shared runtime pause/resume also controls unbound native animation. Native media
and automatic page changes still require their existing supported ownership.
