# Speech animation completion and object location

The saved `wo_bian_cheng_le_yi_ke_shu_20261109_145002.pptx` has valid
adopted audio. Its second page has eight generated entrance cues on the first
speech and six emphasis cues on the next speech. All eight entrances use the
speech-start trigger; simultaneous appearance is the saved instruction, not a
per-word reveal. The first target is the white card background, with separate
title and word shapes. Its saved PPTX has no native timing for these cues: the
teaching scheduler intentionally supplies temporary timing during playback.

The shared runtime released its audio element before notifying speech completion.
`getDuration()` then returned zero; speech-tail validation threw
`SPEECH_TIMING_AUDIO_REQUIRED` despite successful playback. Retain the completed
duration until stop or the next clip. Regression tests exercise the real audio
player and the teaching timeline together, including an end-triggered effect and
the next speech. No TTS regeneration or saved-data rewrite is required.

Object location selected a native shape without refreshing the selection overlay.
Refresh it through the native controller; group members additionally require the
native main-group selection context. Do not remap logical IDs or edit PPT shapes.

Validation on the current dirty request worktrees:

- Runtime: 42 targeted tests, typecheck and bundle build pass.
- Real Chromium editor: four schedule variants pass, checking entrance visibility,
  exit visibility, later emphasis, group-member selection, overlay pixels and
  preservation of timing/history.
- Read-only replay of the actual second page uses its original two MP3s: all
  eight entrance targets are hidden before speech; all 14 effects reach native
  completion, both audio clocks advance, no errors occur, and saved timing is
  restored. The original course and SQLite database are never written.

These checks do not certify all ten pages or a complete new generation task.
