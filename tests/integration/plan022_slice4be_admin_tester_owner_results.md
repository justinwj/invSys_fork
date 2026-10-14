# Tester Setup authority and disposable ownership

4be-A prerequisite,2026-10-13 Pacific. **D8-TESTER-01 approved and scoped-qualified
on unpromoted tester-owner-08.** All six tightenings apply. This is not full4be-A
or Release1 acceptance; the legacy headless discrepancy below remains open.
Catalog41/172 and D18-REPLAY-01/B0 boundaries remain unchanged; no Tester
recording ID is added.

## Focused progression

Controller paths below are under ignored `reports/runtime/`.

| Package / controller | Result | Meaning |
| --- | --- | --- |
| create-recording-02 / `admin-tester-owner-controller/077e6f9d746c458ab37b5eeb0c904de6` | RED25/4 | Actual cleanup deletes after four authority changes. Names admit unowned files. |
| create-recording-02 / `admin-tester-owner-controller/df6babd0a112436bb145d9be7da377f9` | RED31/14 | Authorized Setup works and preserves source context, but creates no receipt and writes after authority changes. |
| create-recording-02 / `admin-tester-owner-controller/21c950522bc040ffae898d33b7b1a511` | RED33/30 | Unmarked, published and nested destinations are adopted; feedback includes entered values. |
| tester-owner-03 / `admin-tester-owner-controller/9000d03790be4fc9b4f9af065d775bed` | GREEN67 | Initial authority, owned reopened cleanup, Cancel, conflicts and fixed feedback. Narrower than the approved contract. |
| tester-owner-03 / `admin-tester-owner-controller/29a0854ffb304fe9b99a019424345b12` | RED66/2 | Lock/error diagnostic writers expose entered values; changing the runtime override before seed creates files outside the reserved root. |
| tester-owner-04 / `admin-tester-owner-controller/9d727f0dd5a44771bb4756087cbd09fa` | RED12/4 | Receipt omits generated `invSys.Publish.log`; default Find clips; both Close controls remain disabled after success. |
| tester-owner-06 / `admin-tester-owner-controller/a4cbc4b89b5a42e8b5fceec1ca5914f4` | GREEN157 | Combined authority, ownership, redaction, scope, layout and interrupted cleanup checks. |
| tester-owner-06 / `admin-tester-owner-controller/96e524183aa545b19f7961d21d541e3b` | RED18/1 | Saved designer caption still says Create Warehouse despite the runtime title. |
| tester-owner-07 / `admin-tester-owner-controller/878bb9701dd5426eb4d92627141d1615` | GREEN164 | Adds saved source-rights revocation, dirty source Auth and corrected saved heading. |
| tester-owner-08 / `admin-tester-owner-controller/5da9780e6d0248fb8c504eb68ddef642` | GREEN164 | Final package retains all focused checks after launcher extraction and headless-facade preservation. |

A refined layout diagnostic adds heading checks (RED14/4), identifying the omitted
artifact by relative generated name with entered root and warehouse ID removed.
These defects are fixed in source. Candidate07 has two reviewed default/enlarged
captures; Find, Close and the saved heading fit and remain usable. Its static
ratchet found four added lines in oversized modAdmin; the unchanged launcher body
is now extracted into small Admin-only modTesterPresentation for candidate08.
Completed behavioral runs above compile five instrumented packages, preserve pins
and settings, close normally and pass delayed Excel1000/1001/1002 audits.
Compile/harness failures are not RED: the extraction's optional-constant cycle was
corrected; open-file hash checks now run after closing the borrowed test workbook
without saving.

## Boundary and coverage

Admin captures the original workbook/session through its real Ribbon callback.
Headless Core owns authorization, exact-root reservation, provisioning and cleanup.
Versioned receipts retain immutable creator provenance; fresh source authorization
is still required. Borrowed workbooks and unmarked/foreign targets refuse. Unknown
files survive; partial cleanup retains ownership for a fresh retry. Feedback and
nested diagnostics are fixed. Explicit headless fixture provisioning remains
separate. No Core-to-form progress callback is added.

`Test-Slice4beAdminTesterSetupOwner.ps1` invokes packaged callbacks and handlers.
`-CheckSetup` covers entry/destination refusal; `-CheckOwnership` adds provenance,
receipt integrity, borrowed files, direct Core denial, interrupted verification,
phase/root changes, picker returns and layout. `-CaptureEvidence` captures the
actual form. Adapters supply dialogs and inject authority changes; they do not
replace owner results. Print interception evaluates actual arguments in memory,
retaining only a leak boolean/count. File checks inspect only appended logging.
No credentials or operational rows are written to test receipts.

The packaged headless comparison found an introduced ByRef-return regression:
candidate07 returns4 PASS/2 FAIL versus baseline create-recording-02's5 PASS/1 FAIL.
`Headless.FailedSetupReturnsNormalizedSpec` passes on the baseline and fails on07;
the primitive facade now returns normalized/defaulted fields even on failed Setup.
Controllers: `tester-headless/24a9a42d77fb4cc1bd447e1219f356f7` and baseline
`tester-headless/0100abd94b704b64aecff86eb2e8b128`. Both preserve pins/settings and
close normally. The legacy `SharePointUnavailable.LocalSetupStillSucceeds` case
fails on both packages: it expects success after an invalid publication path,
whereas existing provisioning returns failure. That pre-existing discrepancy is
open; no offline-publication behavior is changed or claimed accepted here.
Candidate08 (`tester-headless/a46b9f6482ae426a9a3e08dfbf3b7718`) restores the same
5 PASS/1 existing FAIL as the baseline, with normal closure and preserved pins/
settings. `tester-headless-comparison-08.json` confirms every baseline result is
retained and there are no new failures; it explicitly does not label the full
legacy headless suite GREEN.

Final source maintenance (`tester-owner-static-final`, including its ratchet receipt):
357 components,6495 procedures,141064 lines; dynamic calls9/45 and duplicate
groups194 unchanged. All prior size caps hold; oversized modules decrease28->27.
modAdmin2435->2425, modTesterSetup1655->86, modAuth1749->1743. All530 PowerShell
files parse. The two unused Admin getter wrappers were removed after reference
review; the operator-path caller now calls Core directly, while the report facade
and all required headless entry points remain. Five-package compile and actual
callback tests protect this extraction; no scanner-driven deletion was performed.

Candidate08 passes five cold packaged compiles and the Operations dependency
cold-start check (`tester-regression-08/Compile.log`). Its focused controller
records preserved package/tooling pins and settings, normal Excel closure and
zero delayed Excel1000/1001/1002 events. Default/enlarged visible captures under
`slice4be-admin-tester-owner/ab10a62d5ac74319a9a090713f19f566` were reviewed: correct
heading, masked PIN fields, fixed completion, unclipped Find and enabled Close.
Actual cursor diagnostic at2026-10-14T03:28:56Z (`navigation-cursor/f6d87a335d074621968402196007aa83`)
reaches its exact target and restores the original position, without error5.

Accepted Create recording/guide regression on08: **381/381 GREEN**, retaining
all110 owner checks and independent guide conclusions; controller
`admin-create-owner-controller/13e833d02e9e4b87a0818675ff9074e1`.
Settings regression is **202/202 GREEN** at
`create-recording-regression/settings-d871076cfade4b58bea4998380720dd2`.
Both gates preserve pins/settings and pass normal closure plus delayed native audits.
Recording/published-read regression is **80/80 GREEN** at
`create-recording-regression/recording-80b3aa4aaa3440e88a0164b01a6640dd`, with the
same preservation and native-closure checks.
Packaged smoke is **86/86 GREEN** at
`create-recording-regression/smoke-82581b8656e04dfe9570264023fa0a0c`; its tracked
report is restored, packages/settings are preserved, and the delayed native audit is clean.

Full Release1 chain is **32/32 GREEN**, including **48/48 live-role** and **15/15
source Create** checks: `create-recording-regression/chain-b8972c4dc1134be4bb22106ad5e9553c`.
`tester-regression-08/gates.json` records all seven successful gates with zero
delayed Excel events. Its preservation receipt confirms normal closure and unchanged
packages/settings; tracked smoke and chain reports were restored. The final diff
checks pass, and unrelated user changes retain their original hashes.

The full legacy headless suite is not GREEN; its single pre-existing offline-publication
discrepancy remains explicitly open. Candidate08 retains all five baseline passing
checks and introduces no headless regression. No current/NAS package promotion was
performed. Overall4be-A is incomplete; this prerequisite establishes neither Tester
recording nor replay. Reachable yields tested are cleanup confirmation and both
folder pickers, including cancellation; synchronous Core phases disable Excel events
and add no DoEvents, container-combination dialog or separate rename action.
