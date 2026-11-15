# Grouped Quiz picture recovery

The real PPT-master export retained assessment object names but nested them
inside an `assessment` group. Native image projection indexed top-level objects
only. The template prefix also differed from the logical page ID, and long
questions contained export-inserted line breaks. Repeating the same image
adoption could not repair that binding mismatch.

The projector now resolves nested names and the question anchor before mutation.
It rebuilds a wholly owned assessment group from the canonical projection at its
original stacking position. Extra user objects, duplicate bindings and nested
changes during PNG preparation fail closed. The existing transaction restores
the original native object references on rollback. Independent native objects
retain the existing projection path.

Before the fix, four of the five new unit cases failed, including the real
failure message `Quiz object binding changed; refresh before saving`. With the
fix, all five pass; the existing native teaching suite also passes (41 combined).
Type checking and scoped lint pass.

`deckwise-grouped-quiz-recovery.spec.ts` passed against a real four-page
PPT-master fixture containing the exported group. It adopted a local PNG, saved,
reopened and verified a native picture plus unchanged teaching actions. No media
supplier was called by this test. The Desktop integration was rebuilt/staged,
and a separate client check adopted the previously completed image order into
the original saved file, saved/reopened it and retained its four pages and
teaching actions. Only original-order query/download/adoption was invoked.

This validates image recovery, not a new end-to-end paid generation or complete
slideshow playback. A missing second option image remains incomplete until its
original request is reconciled or a genuinely missing request is authorized.
