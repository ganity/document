# Authorized image assets in the DeckWise editor bridge

`DeckWiseEditorBridge.setImageAssetSource(source?)` installs a trusted, session-local
source for reviewed immutable PNG assets. `source.list(binding)` synchronously
returns up to 100 unique metadata pins (`assetId`, SHA256, MIME, byte length and
dimensions); `source.load(assetId, signal)` retrieves authorized bytes. The Host
must enforce current user/team/project membership and review consent when listing
and loading. A hash proves integrity, not authorization. No model URL, path,
credential or byte payload is accepted by the editor command.

Install the source after binding the document. `reset()` clears it; replacement
or removal aborts an active command and requires fresh context. Repeated context
reads re-evaluate the current authorized list and advance editGeneration when it
changes. The bridge advertises `replace_slide_image_from_asset` only for a current
writable, supported native image selection and a nonempty valid list. Its schema
contains only an `assetId` enum. Unsupported groups/crops/masks/placeholders stay
unavailable; use a reviewed latest-clean-PPTX copy candidate instead.

Execution consumes the request once, checks current binding/generation/selection,
and delegates to `replaceSlideImageFromAsset`. Authorization metadata, source
identity and native state are checked around asynchronous retrieval/preload. The
existing consumer copies bytes and verifies exact PNG size, dimensions and SHA256.
An uncertain edit is never automatically retried. The returned result is a native
verification, not a save acknowledgement. Agent output removes transient imageId.

The DeckWise embed save handler refuses saves while `bridge.isExecuting`; existing
pending/uncertain-save guards refuse new commands. Direct bridge consumers must
provide the same save exclusion. This source API is an internal integration seam,
not a new public postMessage transport. A owns authenticated source mounting,
Main/Rust/Skill registration and actual Chat acceptance; C owns asset access.
Without that mounting, no image capability is advertised.

Validation: 183 directed unit cases across bridge, embed save exclusion and native
PNG helpers; workspace declaration rebuild and app typecheck; scoped lint and
format. B53's standalone current bridge bundle also ran through the B51 Agent
tool against the real local Office SDK, saved and reopened: six operations, no
browser errors, three pages/notes and four pictures retained. Host authorization
was a fixture. The old local Office shell supplied SDK/save infrastructure; this
does not validate a rebuilt full embed application, Main Chat or live C access.
