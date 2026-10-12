# Plan 022: Verify Add-ins Published recording

Scope: 4be-A, Architecture v4.11 D18 discovered-control refinement, catalog35.
Eligible NAS recording/guide coverage is accepted on candidate03. This is not
full 4be-A acceptance. D18-LOCAL-01 remains proposed, awaiting approval.

The actual `modAdmin.Verify_AddinsPublished` callback records REQUESTED then
VERIFIED/FAILED. VERIFIED means five required filenames exist with nonzero size;
it proves neither hashes, deployment promotion nor Domain application. Effects
remain Unknown because the callback prepares local workbook surfaces. No source
event references, entered values or paths enter the activity record. The owning
file check runs once and retains its original result wording. Workbook and folder
are captured before optional tracking; an empty captured folder cannot resolve
against a later target. Session changes and storage failures remain non-blocking
with truthful missing evidence and one tracking notice.

## Focused D13 evidence, 2026-10-12

Controller: `tests/tooling/Test-Slice4beAdminVerifyAddins.ps1`, `-Phase RED|GREEN`.
Fixtures use generated non-executable files; these are never loaded as add-ins.
All calls use the packaged public callback; only result-message presentation is
intercepted. This gate is not native-dialog or guide-presentation acceptance.

| Gate | Candidate | Result | Controller suffix | Worker suffix |
|---|---|---|---|---|
| Recording RED | admin-reconcile-recording-06 | 72 PASS / 13 expected FAIL | 55f9c91896204a8db78b1b653f35dd6d | 376a8b1e41d34aa5a62e8d5f91e07530 |
| Recording GREEN | admin-verify-01 | 85/85 PASS | 8d61f755d674491faa3f115a4a55776a | fccee5a2d6bc4361bbc4a260308d4d6a |
| Context boundary RED | admin-verify-01 | 126 PASS / 4 expected FAIL | 7fd2d856a3254cbf83cca0cdd2b10e41 | 948715221fc04e089d4a4b0f6893b38f |
| Context boundary GREEN | admin-verify-02 | 130/130 PASS | fa75dd235092465eb6d3f8a6be0da339 | 9038c97c772b4c028f77fe83fc901936 |

Candidates live under `deploy/validation-`; controller evidence under
`reports/runtime/admin-verify-controller/`; workers under
`reports/runtime/slice4be-admin-verify/`. Raw artifacts remain ignored.

Recording RED fails only registration plus four recording assertions for each
successful, missing and zero-byte case. All business-result, single-invocation,
policy suppression, original/custom workbook, file and inventory checks pass.
GREEN preserves the exact 85 check identities. Context RED preserves that prefix
and adds five actual faults: workbook switch, sign-out, target switch, and failures
of the real pending-file write for attempt and outcome. Its four failures are the
sign-out notice plus target-switch result, root binding and notice. No injected
success or fabricated activity is used. Remaining 126 checks pass.

Each qualified gate includes five instrumented compiles, exact console/JSON
agreement, normal Excel closure and settings/package/tooling preservation. Delayed
Excel Application-event audits report zero failures through 10:32:55.4756716,
10:36:58.3002166, 10:43:54.4633973 and 10:48:39.8649449 UTC respectively.
The final GREEN preserves all 130 boundary-test identities. It restores the
original root/result and one notice after sign-out or target changes; both real
write-failure paths retain their prior GREEN evidence.

Earlier recording attempts had a fixture activation problem and are not accepted
RED. Activating the intended workbook inside the VBA caller and checking the
resolved workbook fixed it. A boundary probe inserted into a continued procedure
declaration produced a compile error; that is also not behavioral RED. The probe
now locates a unique statement inside the procedure. Its controller completed
owned cleanup and restored settings before any rerun.

## Candidate qualification

`validation-admin-verify-02` has five cold compile passes plus Operations cold
startup. `reports/runtime/admin-verify-build-02/verification.json` confirms only
the intended five existing components changed and two were added; 320 prior
components are identical. Frozen recording-06 packages and settings are preserved,
Excel closed normally, and no native failures occurred through 10:45:34.2275846 UTC.
Compiled package pins are stored beside that report. Candidate remains unpromoted.

This intermediate candidate is superseded by candidate03 below. Earlier
Reconcile/Settings/chain evidence does not verify changed packages.

Static `admin-verify-static-02`: 334 components, 6405 procedures, 139642 lines;
28 existing oversized-module caps retained without growth, dynamic calls 9/45,
194 duplicate candidates. Three schemas and 486 PowerShell parsers pass. No new
exception. Regenerate after final acceptance-test edits before closing this entry.

## Expected-conclusion choice gap

The first guide extension stopped because its fixture omitted ACTION_PATH_MAINT;
155 checks passed but it is not accepted guide evidence. The corrected fixture
explicitly grants and verifies that permission. Its next run passed 158 checks
before assigning the absent VERIFIED dropdown value caused VBA error380. Ending
that test invocation was followed by an oleaut32.dll/c0000005 Excel failure at
11:12:36.3570813 UTC and an automatic `/restore` instance. The crash-parent link
was verified and recovery closed through the normal window Close action, without
COM reattachment or force termination. The original controller stayed alive and
restored settings; packages/tooling remained unchanged. This run is not clean
acceptance or a claimed native-error repair.

Controller `a781698ed03f43ae9913ac51d166409a`; worker
`8320ebcdf0a84e449957edd723d0b684`, in the evidence roots above. The exact failed
form action was `frmActionPathExpectation/cboExpectedOutcome/Write`. Future guide
tests inspect offered values before assigning a required outcome.

A smaller `-ExpectationChoicesOnly` gate now uses an actual recorded callback,
library selection, Expected conclusion form, control-selection handler and Cancel.
It reads the displayed choices without assigning an absent value. Qualified RED
on verify-02: **22 PASS / 1 FAIL**, solely
`AdminVerify.Expectation.VerifiedOutcomeOffered`. Controller
`34eaead8a0fd4bb08b9c76d75b5c7ffd`, worker `8cfe8f14b1ba4afd9d6f3b4b363c7d70`;
UTC11:15:12.8333293–11:16:51.9582191. Five compiles, exact console/JSON, normal
closure, preservation and zero native failures through11:17:03.9582191 pass.
REQUESTED/FAILED remain available, no unregistered outcomes appear, and Cancel
preserves the journal, files, inventory and custom value. Source audit localizes
the omission to `modExpectationDraft.Choices`' explicit outcome list. Adding the
registered VERIFIED choice implements the existing D18 contract; it changes no
completion meaning, permission or local-target eligibility.

Qualified GREEN on `validation-admin-verify-03`: **23/23**, exact RED identities.
Controller `9587cca0b26d4065a41027ec37a69cc9`, worker
`075ecbb7d7f4472e935f6b666cde4d3c`. Five instrumented compiles, normal closure,
exact console/JSON and preservation pass; zero native failures through
11:20:33.8062877 UTC. This verifies choice availability and cancellation, not the
remaining complete guide journey or a native-crash repair.

Build03 qualification: five cold compiles plus Operations startup; six existing
components changed, two added, 319 prior components identical. Frozen recording-06
packages and settings preserved, normal closure, zero native failures through
11:18:13.1449968 UTC. `admin-verify-build-03/compiled-pins.json` is the current
candidate pin set. Build02 remains preserved as the outcome-choice RED baseline.

Final-source static `admin-verify-static-06` retains334 components/6405 procedures/
139642 lines, all28 existing caps without growth, dynamic9/45 and194 duplicate
candidates. Three schemas and486 PowerShell parsers pass. This replaces earlier
static snapshots for the current files; no new exception is required.

## Complete recording/guide gate, 2026-10-12

`Test-Slice4beAdminVerifyAddins.ps1 -DeployRoot deploy/validation-admin-verify-03
-Phase GREEN -CheckGuidePaths`: **202/202**, exact130 prior prefix. Controller
`feb21c086739433bb0f54a6df7ca5eaa`; worker `ac6cb0c8ed734814b477e75d13579b5e`.
UTC11:20:36.1573757–11:48:24.8107722; zero native failures through
11:48:36.8107722. Five instrumented compiles, exact console/JSON, matched compiled
pins, normal closure and settings/package/tooling preservation pass.

Three independent journeys invoke the actual callback twice and persist exact
records/journals. Source and observed captures stop normally; a real failure of
the second outcome write produces an immutable Incomplete capture while the
file check still runs once and succeeds. Actual publication preserves all records.
Actual guide/expectation controls author two VERIFIED steps. A reader cannot
author; its fresh observed recording concludes CommandCompleted, with exact guide,
journal and publication references and no Domain-application claim. Interrupted
evaluation remains Incomplete with one matched step and one outcome mismatch;
neither evaluation borrows source-recording success.

How-To, Diagnostic and Compare both preserve identical evidence and pass layout,
multiline text, provenance and displayed-conclusion checks. All six owned-window
captures were visually reviewed: four observed captures (three views plus scrolled
conclusion) and two interrupted captures (Compare both plus conclusion). Original
activity/journals, dummy files, inventory and host custom values remain unchanged.
Result-message presentation is still intercepted only for the Admin callback;
these captures establish the real guide views, not a native result-dialog claim.

This closes the scoped guide proof. Current-package regression evidence follows.

## Current-package regressions

Reconcile on verify-03: **360/360**, exact330 prior prefix; controller
`admin-reconcile-controller/30e3fac433da4e91be5a305704188978`, worker
`slice4be-admin-reconcile/4fda40d42dd14c8abe6355b940e89494` under `reports/runtime/`.
UTC11:48:54.9447675 to12:16:41.8405233; zero native failures through
12:16:53.8405233. Five instrumented compiles, exact console/JSON, matched compiled
candidate03 pins, preservation and normal Excel closure pass. Reviewed interrupted,
Awaiting and applied-result captures retain the distinct conclusions and exact
source provenance. Submitted observations and earlier evaluations stay immutable.
Settings on the same candidate: **202/202**, exact prior ordered checks and
console/JSON agreement, controller
`admin-verify-regression/settings-915db266eaa2422fa38ce0150e708d74`, worker
`slice4be-tracking-settings/b1746d731f8a499595ad6cc100682930`.
UTC12:18:51.8300311 to12:23:35.7997097; zero native failures through
12:23:47.7997097. Five instrumented compiles, matching candidate pins, preserved
settings/packages/reports and normal final exit pass. The preference restart exits
unassisted and releases ten completed-host COM references without errors.
Operations Settings works without the Admin package and retains context/binding
and permission checks.

Packaged smoke on verify-03: **86/86**, exact prior order, both unassisted exits;
`admin-verify-regression/smoke-1a7aafcfcdd14717a31e17c85f9213f3`.
UTC12:23:55.6419149 to12:24:18.5386200; zero native failures through
12:24:30.5386200. Compiled pins and settings/package/report preservation pass.
Full Release1 regression on verify-03: **32/32 chain, 48/48 live-role and 15/15
Create Warehouse**, each retaining exact prior order;
`admin-verify-regression/chain-39a292c059474d2c9761e42e341e90a5`.
UTC12:24:41.5678891 to12:30:21.1755859; zero native failures through
12:30:33.1755859. Compiled pins and settings/package/report preservation pass;
Excel closes normally. This includes reusable Production's consecutive batches,
captured Receiving/Shipping actions, Boxing, final refresh and restart/reconciliation.

First Settings attempt stopped27/2 at its stale catalog34/154 fixture assertion:
`admin-verify-regression/settings-beba48834f14481782f87a08220a9f38`, worker
`slice4be-tracking-settings/ce01dec64004440b83c75656b03af13e`.
Normal closure and settings/package/report preservation pass. This is neither
accepted regression evidence nor product RED. Updated only the fixture's expected
catalog/count to the specified35/155; no prior check is removed.

Scoped acceptance: Verify Add-ins Published recording, suppression/context/write
failure boundaries, independent guide conclusions and the VERIFIED picker pass
on the unpromoted candidate03. Static06, compiled packages, layout/captures and
current-package regressions pass. Existing result-message presentation is
intercepted in the focused gate; no new native result-dialog acceptance or native
crash repair is claimed. D18-LOCAL-01 and remaining Admin coverage keep4be-A open.
