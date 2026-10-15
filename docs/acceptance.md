# Acceptance

Fixture results below were verified locally on Node.js 24 during extraction.
CI independently checks supported Node versions; a green workflow is separate
from real-provider acceptance. No real email or OAuth credential is a fixture.

| Conditions | Action | Evidence |
|---|---|---|
| Missing host credential / foreign host URL | Refuse startup or host construction | PASS: host.test.mjs |
| Valid local host | Send only the installation credential and scoped payload | PASS: host.test.mjs |
| Invalid JSON / oversized response / inactive socket | Fail boundedly without transport retry or uncaught error | PASS: host.test.mjs |
| Actual entrypoint idle, then SIGTERM | Exit cleanly | PASS: host.test.mjs |
| One account fails, another is healthy | Mark retrying and continue the other account | PASS: runtime.test.mjs |
| Operation ready vs uncertain | Request execution only for ready IDs | PASS: runtime.test.mjs |
| Operation execution fails | Preserve unknown outcome; do not retry in the cycle | PASS: runtime.test.mjs |
| Import pending vs idle / unavailable host | Wait 500 ms vs 30 seconds | PASS: runtime.test.mjs |
| Large or JSON-escaped bodies | Split by serialized bytes; advance cursor only after durable batches | PASS: microsoft.test.mjs |
| Graph cursor expired | Reset folder import within the initial-history bound | PASS: microsoft.test.mjs |
| Hostile or foreign delta link | Reject outside the pinned /me namespace | PASS: microsoft.test.mjs |
| Message moved between folders | Distinguish deletion tombstones from archive/trash moves | PASS: microsoft.test.mjs |
| Later import pages contain older messages | Apply the selected initial-history bound on every page | PASS: microsoft.test.mjs |
| Real provider OAuth, receive, send and reconnect | Owner activates and validates with a real account | UNVERIFIED: activation required |
| Native Windows / physical phone callback | Validate on the actual device | UNVERIFIED: activation required |

Source/digest review, host-generation rejection, immutable operation claims and
restore safety belong to host integration tests; these fixtures do not claim
to verify host authorization or provider delivery.
