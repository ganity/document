# Pinned Community plugin launch

B23 extends the isolated B21 online launcher with the B22 read-only Community
plugin. Existing callers without a plugin pin still reject all plugin config.
No offline entry, default plugin loading, URL parameter, Connector, Agent or UI
registration changes.

`DocsOnlineBinding.communityPluginConfigUrl` is an optional **Host-installed**
asset URL, independent of the returned config/document. When present, config must
contain exactly one `pluginsData` URL equal to that pin and exactly the fixed
Community plugin GUID in `autostart`. Missing, extra or unknown fields reject.
The pin must be a normalized HTTPS or loopback HTTP URL ending in `/config.json`,
without query/fragment/credentials. Mount additionally requires the top-level
Host origin because the plugin's trusted top-window handshake requires it.

The native JWT payload must contain the same pinned plugin fields. This catches
adding plugin config after signing. Decoding is bounded by the existing config
size limit and handles UTF-8. It **does not verify the signature, expiry or
server authorization**; the matched Docs server must do that. Host preserves
all original config fields/token. Never append plugins to C's existing signed
JSON or copy its plugin URL to manufacture the independent Host pin.

Mount installs owned-window discovery before constructing `DocsAPI.DocEditor`,
waits for the plugin handshake, and returns optional `communityReads` alongside
`destroy`. An internal abort lifecycle closes discovery and the connected port
on cancellation, failure or destruction. Constructor failure consumes the
pending promise rejection. Existing no-plugin launch behavior remains intact.
The reader exposes only native shape text, with the B22 bounds and capability
limits; it is not canonical editor context, edit generation or saved-source
success. A timeout requires an explicit editor/plugin restart.

Validation: 209 directed tests (34 new, 175 regression), workspace declaration
rebuild/typecheck, strict scoped lint, format and independent module build.
Negative controls remove the signed-payload check, exact URL pin, and internal
lifecycle abort independently: the corresponding new tests fail, then exact
source bytes restore and all 34 new tests pass.

The actual browser uses the built launcher/Host/plugin modules, official test
bootstrap and real local native Office v9.3 plugin/SDK. A **scripted Docs API
shell** creates that offline editor as a descendant; ephemeral HMAC-signed
fixture config is preserved exactly. Three mounts, five native plugin reads,
manual native text change and two ordinary save/reopen rounds pass. ZIP readback
preserves both other pages, three notes and both image payloads; source is
unchanged. This is not DocumentServer/JWT verification, authenticated two-user
editing, plugin writes, main Chat or matched-server installation acceptance.

C3f already supplies signed source-pinned identities/config without plugins.
C must sign the jointly pinned native extension; A must install the fixed local
resources and consume this mount result. Do not alter C/A authority or claim
W13 complete. Real online generation/concurrency/write guards remain pending.
