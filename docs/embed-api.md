# iframe Embed API

This project supports embedding into any web application via iframe. The recommended pattern is: **the parent system handles auth, file fetching, and upload; the iframe handles editing only.** Tokens, cookies, and business APIs stay in the parent — the editor never sees them.

A working demo is available at `/embed-demo.html` (includes sha256 logging for debugging).

---

## Embedding the editor

```html
<iframe
  id="documentEditor"
  src="https://your-deployment/editor?embed=1"
  style="width: 100%; height: 720px; border: 0"
></iframe>
```

### DeckWise restricted mode

Add `deckwise=1` to keep the editor's Save button and Ctrl+S while hiding and
blocking Download, Save As, other-format export, and Print:

```html
<iframe src="https://your-deployment/editor?embed=1&deckwise=1&embedOrigin=https://your-system.example.com"></iframe>
```

In this mode the editor accepts `document:open-buffer` for all supported formats:
Word (`docx`, `doc`, `odt`, `rtf`, `txt`), Excel (`xlsx`, `xls`, `ods`, `csv`),
PowerPoint (`pptx`, `ppt`, `odp`), and PDF (`pdf`). Set `readonly: true` for preview.
Unsupported formats return `document:error`; the host can offer the system's default application.
Recognizing a format does not guarantee conversion of every file; the underlying
converter may still reject some larger PPT and ODP files.
URL payloads remain disabled. Saving uses the document's own format, with legacy
`doc`, `xls`, and `ppt` saved as `docx`, `xlsx`, and `pptx`; exporting to another
format is blocked. Clicking the internal Save button
posts `document:save-requested`; the host should answer by sending the normal
`document:save` command so the edited file returns through `document:saved`.

> The editor lives at `/editor`; the homepage `/` is a static landing page. Older links to `/?embed=1`, `/?src=`, `/?file=` and `/?new=` still work -- `/` redirects them to `/editor` with the same query.

Commands are accepted only from the immediate parent window. Without
`embedOrigin`, the first accepted parent command fixes the origin for the lifetime
of the editor page; later commands from a different origin are ignored.
To restrict messages to a specific origin from startup, add `embedOrigin`:

```html
<iframe
  id="documentEditor"
  src="https://your-deployment/editor?embed=1&embedOrigin=https://your-system.example.com"
></iframe>
```

---

## DeckWise PPTX command bridge (W01.1–W01.2)

This optional bridge reuses native tool execution without running a model or
opening an editor-side Chat. It is available only in `deckwise=1` mode and only
for presentations. Existing embed/save callers need no changes.

Include a host-owned binding on `document:open-buffer`:

```js
const binding = { projectId: 'workspace-scope', documentId: 'document', sessionId: 'unique-open-session' };
sendEditorCommand('document:open-buffer', { fileName: 'lesson.pptx', buffer, editorBinding: binding });
sendEditorCommand('document:get-context', binding);
```

`document:context` echoes the binding plus `editGeneration`, `kind: 'slide'`,
one-based `page`, `selectedText` (at most 8000 characters), `selectionCharacters`,
`readOnly`, and the current `capabilities`. All commands need a unique envelope
`id` of 1–128 letters, digits, underscores or hyphens. Each binding ID is a
nonempty string of at most 128 characters. Binding is reset at the next open;
opening without it disables command access. The native API can still be loading
after `document:opened`; context requests then return `document:error`.

### Optional saved-source page lineage

Set `trackPageLineage: true` on a bound PPTX `document:open-buffer` to opt in.
At that open's native content-ready callback, the bridge hashes the opened
file and assigns random tokens to the actual slide objects. `document:context`
then includes `pageLineage: {sourceSha256, openedPageTokens, pageTokens}`. Arrays
follow native page order. Moved pages keep tokens, new/copied pages receive new
tokens. Tokens expire on reopen; they are not logical or OOXML IDs.
`sourceSha256` always denotes the session's opened bytes, even after saving.

A host-managed `document:saved` includes an optional `pageLineage` receipt:
`{protocol:'deckwise-office-page-lineage/1', projectId, documentId, sessionId,
saveRequestId, sourceSha256, outputSha256, openedPageTokens, pageTokens}`.
The output digest covers the exact returned File. The receipt is withheld when
the document/session/pages/history change during export or hashing, or when the
opening baseline is unavailable. Missing evidence never implies identity success.
Normal save/ack behavior remains unchanged. The optional font-size capability is described below.

Only the trusted Host may consume a receipt directly from the bound Office view.
Compare binding, save request, independently computed file hashes and a complete
mapping bound to the opened source. Persist the receipt with exact saved revision
and mapping revision before using it after restart. Persistence failure preserves
the PPTX and old teaching data, with teaching reassociation fenced. Receipts are
not signatures/authorization; renderer/model supplied JSON is not evidence.
Page-only evidence never grants object identity. Inserted/deleted/copied pages
require explicit identity review; the initial Host helper refuses them. Changed
narration/cues/audio still require the existing review/CAS policy.

Use the returned binding and generation to execute an advertised tool:

```js
const id = crypto.randomUUID();
editorWindow.postMessage(
  {
    type: 'document:execute',
    id,
    payload: {
      ...binding,
      requestId: id,
      editGeneration: context.editGeneration,
      tool: 'get_presentation_text',
      input: { maxChars: 8000 },
    },
  },
  editorOrigin,
);
```

The response is `document:command-result` with matching envelope ID and payload
binding, `requestId`, `editGeneration`, `tool`, executor `status`, and `result`.
`get_presentation_text` returns bounded `text`, `truncated`, `slidesRead`, and
`totalSlides`; `maxChars` is an optional integer in 1–8000. This reader reads the
presentation, while the context identifies the current page and selection.

Two existing native write tools are enabled on writable, tracked targets:
`replace_selection` replaces a supported native text selection, and
`add_slide_text` inserts a text box on the current page. Both accept only
`{ text }`, limited to 8000 UTF-16 code units. Replacement requires nonempty text
and a nonempty selection; insertion requires nonblank text. The native executor
preserves requested whitespace/newlines, checks exact text and unaffected
content, and uses its existing rollback checks on failure. No generic JavaScript
or OOXML write is exposed. Unsupported tools and extra fields are rejected.

Writes require executor `status: 'verified'`, a current continuation target,
and a verified native result: `{ replaced: true, length, verified: true }` or
`{ page, verified: true }`. The reader's `sent` status is not mutation proof.

`set_slide_font_size` is additionally advertised only for a supported nonempty
plain-text selection in one writable current-slide text object. Its sole input
is `{ fontSizePt }`, from 1 to 96 in half-point increments. It checks native size
and unchanged text/selection, returning `{ fontSizePt, page, verified: true }`.
Master mode, slideshow, busy state, missing SDK methods and ambiguous selections
withhold the capability. The host must use fresh generation/ownership checks,
correlate exact size/page, and never automatically retry an uncertain edit.
It does not save. Existing peers with only the original three tools remain valid.

`set_slide_font_family` uses the same one-object plain-text selection guards.
Its input is `{ fontFamily }`; choose an exact name from the latest capability's
`inputSchema.properties.fontFamily.enum` (at most 100 names, each 1–128 chars).
Only families whose regular, italic, bold and bold-italic files are already
loaded are offered, with an idle native font loader. This avoids the SDK's
asynchronous font-download callback applying to a later selection. The native
family, unchanged text/selection and page are checked before returning
`{ fontFamily, page, verified: true }`. Font aliases may use installed substitute
faces. No font download, automatic save or automatic retry is exposed.

Generation binds a native target snapshot, not a persisted document revision.
Changing pages, selection, native history, frame or user interaction invalidates
it. Commands are serialized and consume their snapshot/request ID; after an
execution, fetch fresh context and use a new ID. Duplicate IDs, expired sessions,
stale snapshots and changes during reads return `document:error`. The bridge
keeps at most 4096 executed request IDs per open session. The host must never
automatically replay a timed-out command. Parent-source and origin checks remain
mandatory; these routing identifiers do not implement team permissions.
Resetting the editor binding aborts its active action. A host timeout alone
does not cancel native execution; inspect uncertain writes before retrying.
Ending the host wait does not cancel native execution already in progress.

## Host-managed save confirmation

A host that writes returned files to disk should send `document:save` with
`payload: { requireHostCommit: true }` and a unique ID (1–128 word/hyphen
characters). The editor returns `document:saved` with that ID and the real
exported File, but keeps the SDK save point pending. The host must preserve the
ID through asynchronous Blob conversion and only consume the current export
once. This option disables fallback to the unedited original file.

After writing, send `document:commit-save` with the **same ID** and
`payload: { success: true }`, or `{ success: false }` if persistence failed.
The editor replies `document:save-committed` with `{ success, dirty }`. Only a
matching successful acknowledgement and an unchanged native history clear dirty.
Edits made during conversion/persistence stay dirty. When a successful write
changed disk contents but the editor moved (including undo), its obsolete saved
mark is invalidated so an old SDK clean callback cannot falsely report success.

Only one save can await host confirmation. Early, stale and foreign-source
acknowledgements are ignored. Failed disk writes allow another explicit export;
uncertain SDK export failures require inspection and reopening before retry,
because a late untagged SDK stream could otherwise satisfy a newer request.
The host must bound export/acknowledgement waits and avoid switching or closing
the document during persistence. A timeout is not cancellation.

In DeckWise mode, an active SDK export blocks editor context/commands, opening
another document and a second save, even without `requireHostCommit`. After a
successful legacy export this guard ends; host-managed saves continue waiting
for the matching persistence acknowledgement. Manual edits during export still
keep the document dirty.

Without `requireHostCommit: true`, the existing embed export behaviour remains.
The option is available to all embedding hosts; DeckWise opts in explicitly.
Older editor deployments do not implement confirmation, so hosts must never
report a clean saved state merely from receiving `document:saved`.

## Sending commands

Include an `id` on each command to match it to the response:

```js
const iframe = document.getElementById('documentEditor');
const editorOrigin = 'https://your-deployment';

function sendEditorCommand(type, payload = {}) {
  const id = `${Date.now()}-${Math.random().toString(16).slice(2)}`;
  iframe.contentWindow.postMessage({ id, type, payload }, editorOrigin);
  return id;
}

window.addEventListener('message', (event) => {
  if (event.origin !== editorOrigin) return;
  const { id, type, payload } = event.data || {};
  if (!type?.startsWith('document:')) return;

  switch (type) {
    case 'document:ready':
      console.log('Editor ready');
      break;
    case 'document:opened':
      console.log('Opened', id, payload);
      break;
    case 'document:saved':
      console.log('Saved', payload.fileName, payload.file);
      break;
    case 'document:error':
      console.error('Error', payload.message);
      break;
  }
});
```

---

## Opening a document

### From URL

> **The URL must allow CORS.** The browser -- not a server -- fetches the file,
> so the response needs an `Access-Control-Allow-Origin` header that covers the
> page the editor runs on. Without it the request is blocked before the editor
> sees a single byte, and the same is true of `?src=` and `?file=` on the URL.
> This is the single most common thing to get wrong when integrating: if a file
> opens when you download it by hand but not through the editor, check the
> response headers first. Cross-origin redirects have to keep the header too.
> When you cannot add it (a third-party host, a signed URL, anything behind
> auth), fetch the file in the parent page and pass the bytes with
> `document:open-buffer` instead.

Remote URL downloads are limited to **100 MiB of decoded response data** and
**60 seconds**, including the response body. Download larger files separately
and open them with `document:open-file` or `document:open-buffer`. These limits
apply to URL downloads, not local file opens.

A failed, timed-out or superseded URL open returns `document:error` for its
request ID. Opening another document cancels the older download; late responses
and pending conversions cannot replace the newer editor. Encoded `filename*`
response headers take precedence over `filename`.

```js
sendEditorCommand('document:open-url', {
  url: 'https://example.com/files/demo.xlsx',
  fileName: 'demo.xlsx',
  readonly: false,
});
```

If the URL requires auth headers, pass `fetchOptions`. For protected files it is preferable to fetch in the parent system and pass the binary:

```js
sendEditorCommand('document:open-url', {
  url: 'https://example.com/api/files/1',
  fileName: 'demo.xlsx',
  fetchOptions: { headers: { Authorization: `Bearer ${token}` } },
});
```

### From a file picker

```js
const input = document.createElement('input');
input.type = 'file';
input.accept = '.xlsx,.xls,.csv,.docx,.doc,.pptx,.ppt';
input.onchange = () => {
  sendEditorCommand('document:open-file', { file: input.files[0], readonly: false });
};
input.click();
```

### From an authenticated fetch (recommended for protected files)

```js
const response = await fetch('/api/files/1', {
  headers: { Authorization: `Bearer ${token}` },
});
const buffer = await response.arrayBuffer();
sendEditorCommand('document:open-buffer', { fileName: 'demo.xlsx', buffer, readonly: false });
```

---

## Read-only mode

Set at open time via the `readonly` field, or toggle at any time:

```js
sendEditorCommand('document:set-readonly', { readonly: true });
```

In read-only mode, editing is disabled and `document:save` returns `document:error`.

---

## Saving and uploading

The save command exports the current document and returns a `File` via `document:saved`. Default format is the open document's own format (a `.docx` saves as DOCX, a `.csv` as CSV, ...); pass `targetExt` to change it.

```js
sendEditorCommand('document:save', { targetExt: 'XLSX' }); // XLSX, DOCX, PPTX, CSV
```

By default the command waits for the editor to return the **edited** file. If it times out, `document:error` is returned — this prevents accidentally uploading the original unchanged file. To opt in to returning the original on timeout:

```js
sendEditorCommand('document:save', { targetExt: 'XLSX', returnOriginalOnTimeout: true });
```

Upload from the parent:

```js
window.addEventListener('message', async (event) => {
  if (event.origin !== editorOrigin) return;
  const { type, payload } = event.data || {};
  if (type !== 'document:saved') return;

  await fetch('/api/files/1', {
    method: 'POST',
    headers: { Authorization: `Bearer ${token}` },
    body: payload.file,
  });
});
```

> **Note:** Do not rely on `file.size` alone to detect changes. `.xlsx` is a zip archive — a minor edit can produce the exact same byte count. The built-in `/embed-demo.html` logs a `sha256` hash on every save for easier debugging.

---

## Query current state

```js
sendEditorCommand('document:get-state');
// Response: { type: 'document:state', payload: { readonly: false, hasDocument: true, dirty: false } }
```

---

## Tracking unsaved changes

The editor tells the parent whether the open document has edits the parent does not hold yet, so the parent can save when it matters (switching documents, an idle timer, closing its own page) without exporting documents nobody touched.

```js
window.addEventListener('message', (event) => {
  if (event.origin !== editorOrigin) return;
  const { type, payload } = event.data || {};
  if (type === 'document:dirty-changed') {
    console.log(payload.dirty ? 'Unsaved changes' : 'Everything saved');
  }
});
```

- `document:dirty-changed` is pushed when the flag flips, not on every keystroke: `{ dirty: true }` after the first edit, `{ dirty: false }` once a `document:save` covered every edit (or undo went back to the saved state).
- A successful `document:save` clears the flag: the parent now holds the bytes. Edits made while the export was running are not in those bytes, so they keep the document dirty -- check `payload.dirty` on `document:saved` and save again when it is `true`.
- If the upload on the parent side fails, the editor cannot know: keep your own "needs saving" state until the upload succeeds.
- `document:save` with `returnOriginalOnTimeout: true` never clears the flag, since the file returned on timeout may be the original.
- Opening another document resets the flag to `false`.

A save-on-idle loop for the parent could look like this:

```js
let idleTimer;
window.addEventListener('message', (event) => {
  if (event.origin !== editorOrigin) return;
  const { type, payload } = event.data || {};
  if (type === 'document:dirty-changed' && payload.dirty) {
    clearTimeout(idleTimer);
    idleTimer = setTimeout(() => sendEditorCommand('document:save'), 30_000);
  }
  if (type === 'document:saved' && payload.dirty) {
    // Edits landed during the export: they still need a save.
    clearTimeout(idleTimer);
    idleTimer = setTimeout(() => sendEditorCommand('document:save'), 30_000);
  }
});
```

---

## Message reference

| Direction       | Type                        | Description                                     |
| --------------- | --------------------------- | ----------------------------------------------- |
| parent → iframe | `document:open-url`         | Open document from URL                          |
| parent → iframe | `document:open-file`        | Open document from `File` / `Blob`              |
| parent → iframe | `document:open-buffer`      | Open document from `ArrayBuffer` / `Uint8Array` |
| parent → iframe | `document:set-readonly`     | Set read-only or editable                       |
| parent → iframe | `document:save`             | Save and return `File`                          |
| parent → iframe | `document:get-state`        | Query current state                             |
| iframe → parent | `document:ready`            | Editor initialised                              |
| iframe → parent | `document:opened`           | Document opened                                 |
| iframe → parent | `document:readonly-changed` | Read-only state changed                         |
| iframe → parent | `document:saved`            | Save complete, file returned (with `dirty`)     |
| iframe → parent | `document:dirty-changed`    | Unsaved-changes flag flipped (`{ dirty }`)      |
| iframe → parent | `document:state`            | Current state response                          |
| iframe → parent | `document:error`            | Operation failed                                |

`set_slide_text_color` is an optional native selection capability with input
`{ "colorHex": "336699" }` (six uppercase RGB digits, no prefix). It uses the
same current-slide plain-text and writable-state guards as font size. Native
solid sRGB fill, exact text/selection and page are verified before returning
`{ "colorHex": "336699", "page": 2, "verified": true }`. Legacy `get_Color()`
does not reflect PPT Unifill; verification reads the calculated native text fill.
No automatic save/retry, alpha, theme color, gradient or shape fill is supported.

## Native PNG replacement integrity

The existing reviewed-asset replacement path checks every PNG chunk CRC and requires bounded, complete chunks, one IHDR, image data and a final IEND before native registration. Truncated chunks, trailing bytes and corrupted checksums reject without changing the selected image; decoder success alone is insufficient. Asset IDs, hashes and the explicit-save boundary remain unchanged.

## Reviewed image transport

In bound DeckWise mode, the trusted parent may send `document:review-image` with
`{binding, asset}`. `asset` is exactly the pinned `{assetId, sha256, mime:'image/png',
byteLength, width, height}` metadata for one approved PNG; `null` revokes it.
The response is the current `document:context`. This registers only a session-local
selection, not server authorization. Parent Main must own approval and reauthorize
membership/source/session before each execution.

Send `document:execute-image` with `{command, bytes}` and a message ID equal to the
finite command's requestId. `command` is the existing bound
`replace_slide_image_from_asset` command; `bytes` is a fresh Uint8Array verified
against the reviewed pins. Only this command lease can load them; ordinary
`document:execute` cannot reuse them. Bytes clear on success, failure or reset.
Current generation, one supported selected image, read-only/save/session guards,
PNG integrity and native verification remain required. The correlated
`document:command-result` does not save the document. No credentials, URL or
filesystem path belong in either payload.

## Current page preview

The current slide is drawn directly by the native renderer to a canvas with a
1600-pixel longest edge and its original aspect ratio. The cached low-resolution
slide thumbnail is not used. Unsupported page dimensions or missing native
drawing APIs reject; the renderer flag and shared font cache are restored after
drawing, including failures. The temporary canvas is released before returning.

`document:render-page-preview` accepts `{ binding, editGeneration, omitNativeText?, page? }` (the optional flag must be boolean). Optional `page` is a one-based safe integer, at most 5000 and within the opened source. It renders that page directly without navigating the displayed page; omitted `page` retains the current-page default. The response token belongs to the rendered page.
It renders a page from the opened PPTX and returns
`document:page-preview` with the source SHA-256, opened page token, one-based
page number, `image/png` MIME and a bounded PNG `Uint8Array`. It requires the
fresh current generation and session-bound source/page lineage; stale pages,
active saves, unknown fields, unavailable native renderers and oversized or
non-PNG output are rejected. Rendering does not edit or save the source. A host
may use this as reviewed visual input for a separate page candidate and must
retain the returned source/page identity with that candidate.

The preview endpoint requires a fresh read-only PPTX copy with empty native undo history and no unsaved changes. Even a fully undone edit is refused because the opened source hash must describe the rendered page. After editing or saving, reopen the latest clean PPTX in a separate read-only session before requesting a source-bound preview.

With `omitNativeText:true`, the temporary renderer skips native text and returns
`nativeTextOmitted:true` with the PNG. Only ordinary opaque shapes and pictures
on the page, layout and master are supported, including nested ordinary groups
whose descendants meet the same rules. Traversal rejects nesting deeper than
16 group levels or more than 1,000 total objects across those trees, including
cycles. Unknown objects, charts, tables, warped text, signatures, translucent
fills and outer shadows reject before drawing. Normal previews
retain their previous behavior. This render does not remove text inside raster
images and is **not an automatically certified clean plate**. The Host must
review it before using it as a candidate background; unsupported pages need a
separately reviewed clean plate. Source/hash/page/generation/read-only/save
fences and the 1600px/8MiB budgets apply unchanged.

### Selected text box dimensions

`set_slide_text_box_size` accepts exactly `{widthMm, heightMm}`: finite numbers
in 1–1000 millimeters. It resizes the entire ordinary rectangular text box
containing the current nonempty plain-text selection and preserves its top-left
position. The complete box must fit on the current page. The same narrow native
selection/geometry guard as text-box positioning applies; ordinary image peers
are allowed, while groups, placeholders, connectors, rotations and flips of the
target are excluded. Capability is withheld in readonly, loading, busy,
slideshow/master or unsupported contexts.

Use a fresh advertised capability and edit generation. Native
`ImgApply({Width: widthMm, Height: heightMm})` is followed by exact dimensions,
unchanged position, text, selection, object and page readback. A verified result
is `{widthMm, heightMm, page, verified: true}`; uncertain results require visual
inspection/Undo and must not be retried automatically. Text may reflow or
overflow. No implicit save, auto-fit, group resize or arbitrary XML/JS is exposed.

### Host-only native image insertion

`prepareSlideImageInsertion(png, {xMm, yMm, widthMm, heightMm}, signal?)`
prepares one already authorized/reviewed local PNG for the current ordinary slide.
Positions are explicit nonnegative millimeters, dimensions at least 1 mm, and the
whole box must fit on the page. The explicit box stretches the image; no automatic
aspect-ratio fitting or overlap avoidance is implied. The canvas selection must
be empty: the helper does not discard a user's active object/text selection.

PNG signature/chunks/CRCs, the 4 MiB budget, 1–4096 pixel dimensions and native
registration/decoded dimensions are checked before editing. Preloading changes
only the image cache. Page, history, interaction, peer geometry/text/image data,
readonly/loading/busy/slideshow/master/notes and slide-lock checks run again after
preload and immediately before applying. Geometry is captured before awaiting.
A stale, cancelled, failed or timed-out preparation never sends an insertion.

The returned frozen preparation has `page`, `box` and a single-use synchronous
`apply()`. It uses the SDK's existing native action/controller image creation path
with the requested box in one Undo action. Native readback verifies exactly one
appended ordinary rectangular picture, its selection, URL/geometry, original
peer order/data and page dimensions. A successful receipt is
`{page, imageId, box, verified:true, saved:false}`. An uncertain native result
requires inspection/Undo and must not be retried automatically. Saving is explicit.
Native object IDs and the newly registered media filename can change on export.

This is an independent Host module, **not a newly advertised bridge/Chat tool**.
The caller owns current session/edit-generation/save checks and authoritative
asset authorization/review immediately before preparation and apply. No URL,
model code, filesystem path, provider call, source-file write or authorization
grant is accepted here. Existing image replacement shares the unchanged bounded
PNG validation/preload path. Public bridge/Main/IPC/Agent mounting remains a
separate owner integration gate.

### Asset-bound native image insertion

The optional `insert_slide_image_from_asset` command accepts exactly
`{assetId, xMm, yMm, widthMm, heightMm}`. The independent existing-Chat Agent
module exposes `ppt_insert_slide_image_from_asset`, adding current `editGeneration`.
Only a PNG asset ID enumerated by the latest capability is accepted. The Host
continues to own authorization/review of both asset and explicit placement;
capability metadata is not an authorization grant. No model URL or bytes are
accepted in the model command. Native image identity and host-only fields are
removed from model-visible receipts.

The current native slide must be writable with an empty canvas selection.
Coordinates are finite 0–1000 mm; dimensions 1–1000 mm; the whole rectangle must
fit on the page. The explicit rectangle stretches the image and may overlap
content, so it requires review. No automatic crop, aspect fitting or placement
inference is provided. Native single-action insertion/readback uses the existing
Host-only preparation. Original peers, page, exact rectangle and new image are
verified; uncertain writes require inspection/Undo and are never auto-retried.

The existing `document:review-image` metadata and `document:execute-image`
command-scoped fresh-byte path are reused. Reviewed metadata is independent of
bytes; byte leases clear on success/failure/reset. Current binding, edit generation,
asset pins, capability, native target, readonly/cancellation and existing embed
save-exclusion checks remain required. Asset hash/length/dimensions and native
PNG decode/CRCs are verified before mutation. Replacement retains its original
selection and receipt contract. Successful insertion receipts contain
`assetId`, `sha256`, `page`, `xMm`, `yMm`, `widthMm`, `heightMm`, `verified:true`;
saving stays explicit (`saved:false` at the Agent boundary).

This is an additive finite producer contract. Older peers continue with their
existing tools; they must not receive an unknown capability. B's producer now
contains ten finite tools including pending B75 resize. The minimal current-A
proposal adds insertion alone to A's eight-tool baseline and preserves A's privacy
and async fixes. Do not mix producer/full recovery patches with the current-A
proposal. Main/IPC/Host permission dispatch and actual owning-Chat acceptance
remain A-owned; authoritative asset authorization remains C-owned.
