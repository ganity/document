# Native PNG chunk integrity

The previous image preparation path checked the PNG signature/dimensions and native decode, but the actual browser/native loader accepted a frozen C4-style fixture with an invalid IDAT CRC. A correct immutable SHA256 only pins those invalid bytes. This must fail before image registration or document mutation.

The shared native preparation function now verifies chunk framing, CRC32, a single IHDR, image-data presence and terminal IEND under the existing 4 MiB cap. It retains current-selection, generation, read-only, cancellation, preload, one-use native apply and explicit-save checks. No new tool, dependency, protocol, credential or transport is introduced.

The old invalid PNG and pre-fix native save stay in frozen evidence as negatives. A separate generated PNG has a new asset ID/hash and is used for positive verification; invalid archived input is never repaired under its existing identity. Actual browser evidence, regression and package readback are recorded in B60's parent-workspace handoff.
