# Community plugin read bridge prerequisite

The online host and offline direct SDK bridge remain separate. This batch adds
an unregistered Community plugin entry, finite native reader and owned-channel
Host client. There is no Connector, second assistant, caller-supplied code, write,
save command, automatic retry or server authorization in this protocol.

Only `get_presentation_text` is accepted, with an integer `maxChars` from 1 to 8000. Each call uses a fixed `Asc.plugin.callCommand` function with recalculation
disabled. It reads current slide/page count and public shape text from current
native presentation data. Notes, chart/table text, fields and every possible
PPTX object are not promised. Counts above 1000 slides or 4096 shapes on a slide
reject. The result is context text, not durable page identity or saved content.

The protocol is `deckwise-docs-read/1`. It uses one dedicated MessagePort and
Host-owned session ID, strict request IDs and finite input/output validation.
Busy, repeated and unsupported requests never invoke another native command.
One active port is allowed per plugin scope. A native `undefined` return reports
`DOCS_PLUGIN_UNAVAILABLE` without closing the channel: real Docs skips commands
during collaborative action locks. The caller may issue a new read when idle;
no command is automatically replayed. Invalid results, timeout or cancellation
close the Host channel; restart the owned plugin/editor explicitly rather than replay a possibly
lost native queue. A result from a disposed plugin is not delivered.

The plugin asset is hosted on the trusted top-level Host origin. After native
plugin initialization it announces readiness to that origin. The Host listener
must be installed before plugin start and accepts only readiness from a window
under its own editor iframe, with the exact origin and protocol. A sibling
editor cannot borrow the channel. The plugin accepts a transferred port only
from its top-level Host at its own origin. This is a lifecycle/ownership check,
not server identity or JWT validation; same-origin executable code is trusted.
An embedded Host below a different-origin top-level shell is outside this first
prototype. Bind it explicitly before extending this transport.

`public/docs-community-plugin/` supplies only our index/config. Deploy the
separately built entry/Host modules and install the plugin bootstrap from the
matching pinned Docs deployment alongside `index.html` as `plugins.js`. It is
not downloaded at product runtime. No external plugin list is accepted from
document content or URL parameters. Nothing is registered in the offline entry
or in the current signed online config validator; A/C own that installation and
signed-config acceptance. B's existing six offline tools are unchanged. Online
write capabilities must stay empty until selection/generation/concurrency and
save boundaries are verified against real Docs.

The real-browser harness loads the official plugin bootstrap retained by hash
as test evidence, starts our plugin through native `editorConfig.plugins`, and
calls the actual local v9.3 SDK via the dedicated port. It verifies fresh reads
after native page navigation/manual editing and two ordinary local save/reopen
rounds. This is neither a matching DocumentServer pair nor two-user coediting.
The bootstrap expects its `config.json` beside the plugin entry; a differently
named config URL or missing adjacent config breaks initialization. Those failed
probes are retained in B22 evidence.

Negative controls remove Host origin/descendant checks and the endpoint's fixed
tool whitelist. Relevant tests must fail; exact product bytes are then restored.
Current C3f also provides signed source-pinned editor configuration and actual
B21 consumption, in addition to C3e source pinning and C3d callback/key retirement.
The finite plugin configuration extension, page/teaching reconciliation and real
Docs integration remain open. This read protocol does
not replace or duplicate any of those server services.

Primary API references:

- [Plugin entry](https://api.onlyoffice.com/docs/plugins/configuration/entry-point)
- [Asc.plugin](https://api.onlyoffice.com/docs/plugins/interacting-with-editors/overview/asc-plugin)
- [Presentation API](https://api.onlyoffice.com/docs/office-api/usage-api/presentation-api/ApiPresentation/)
- [Shape document content](https://api.onlyoffice.com/docs/office-api/usage-api/presentation-api/ApiShape/Methods/GetDocContent)

## 2026-10-24 amendment

The initial read-only protocol above describes its original delivery boundary.
The same authenticated owned channel now also carries fixed native barrier
capability/quiesce/status/resume operations; it still accepts no arbitrary code
or method name. See [the fixed bridge contract](2026-10-24-docs-native-barrier-bridge.md)
for the exact messages, old-engine behavior and uncertainty rules. Existing
ordinary read behavior is retained; a held native extension may temporarily
refuse `callCommand` reads. This amendment does not claim all-client publication
coordination or production native-overlay deployment.
