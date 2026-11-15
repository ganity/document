# Invalidate Community reads after a native Docs error

The isolated online launcher used to forward `onError` while leaving its read port open. A Host could then read cached editor contents from a failed session as if the channel was usable. Native errors now close the established client and abort pending plugin discovery before the Host callback runs. No native command is replayed.

An already-open Docs frame remains visible so the user can use its recovery UI; the Host must explicitly dispose/remount after checking current authorization, source and session. An error during plugin discovery rejects mounting and runs the existing cleanup path. This does not detect all network loss: browser network state and successful local plugin reads alone cannot prove server synchronization or authorization.

Two focused regressions cover the established-client callback ordering and errors during the handshake. B29 separately runs actual Docs/C expiration, socket and member-revocation probes. Its evidence states exactly what was observed; revoking a platform membership alone does not promise to disconnect a previously authorized Docs socket. No online write capability is added.
