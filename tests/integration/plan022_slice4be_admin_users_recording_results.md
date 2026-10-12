# Slice4be-A Users & Roles write recording

Last verified2026-10-12 UTC. **Focused GREEN353/353 and regression qualification pass.**
Users recording/guide and full4be-A acceptance remain open. Unpromoted candidate:
`deploy/validation-admin-users-recording-01`; no station deployment.

Architecture v4.11 D18's discovered-control refinement governs catalog38/160:
ADMIN_USER_SAVE, ADMIN_USER_DEACTIVATE, ADMIN_USER_DELETE. All157 earlier definitions
remain unchanged. Approved D8-USERS-01 continues to govern authority and persistence.
Core supplies SAVED only after confirmed Auth persistence, REJECTED before writing,
DENIED for authority refusal and FAILED for execution/persistence failure. Declined
Delete confirmation supplies CANCELLED. Only SAVED supports CommandCompleted;
no source-event references or Domain application are asserted. Required actor IDs
remain; account/credential payloads do not enter records.

The real button handlers capture inputs/context before optional observation and
delegate through `RunUserWriteForm` / `modAdminUsersAction.Execute` to the unchanged
authority checks in Core. Core's existing primitive APIs gain an optional outcome
string. The Admin-local input type is never serialized or passed across XLAMs.
The existing confirmation and result wording are retained. Optional tracking has
one fixed notice; suppression is quiet. No normal-write schema repair is introduced.

## D13 evidence

All paths below are relative to ignored `reports/runtime/`; no runtime payloads
or screenshots are committed.

- Qualified RED: controller `auth-read-controller/0816c54c0f5a425b9f458c84620c7b98`,
  worker `slice4be-auth-read/3fe92f1e8cda434bb70528d1ae805086`, unchanged
  `validation-admin-users-authority-01`: **33PASS/37 expected missing-record/catalog
  FAIL**. All nine actual-handler business fixtures pass: saved, dirty-authority
  refusal and real WorkbookBeforeSave cancellation for each command.
- GREEN: controller `auth-read-controller/cc8e0251e2404c11ae92254db2ebc9a9`, worker
  `slice4be-auth-read/5a563ca70d0e4724a8a1e0ac35bd4274`: **353PASS/0FAIL**. All70 RED
  check names and all259 prior authority checks retain their order. Added cases
  prove signed-in Reader denials, declined Delete and global/action/user suppression.
  Correlated pairs, exact actor/effect, empty source references, privacy and honest
  command conclusions pass. Suppressed actions still persist without warnings.
- Each controller's `recording-verification.json` checks exact log/JSON agreement,
  five instrumented compiles, settings/package/tool preservation, fixture restoration
  and normal closure. Delayed Excel audits are clean through19:30:30.4792496UTC for
  RED and19:42:49.2099799UTC for GREEN. No intervention in either qualified run.
- Four reviewed GREEN images show healthy/invalid Auth, the full context refusal
  and retained Settings Save. Credential controls are blank. Modeless display and
  Delete responses use declared test seams; ordinary modal launch/confirmation and
  the tracking-notice presentation are not accepted by these captures. Existing
  long-path/account-copy-button clipping remains open. See `visible-review.json`.

The first attempt, controller `e8896b128871469db21699f31c67306a`, confirmed six missing
record pairs but its cancelled-save fixture forgot to enable Excel events; exclude
its aggregate. The next attempt, `f17e7c7c2e114a14b05e8c0b22d628b7`, exposed unreliable
private MsgBox interception after probe insertion. Seven exact generated-fixture
Delete prompts were declined to let the controller restore normally; exclude this
run. The qualified probe explicitly replaces only the known confirmation call with
its fixed response seam, retaining the actual handler and Core command. These
harness failures are not product RED, native acceptance or a D16 registry decision.

## Build and maintenance

`admin-users-recording-build-01`: five cold compiles and Operations cold start;
329 compiled components are unchanged. The expected five changed components are
the Users form and Core catalog/action codes/Auth command/evaluation modules; the
sole new component is `modAdminUsersAction`. Frozen predecessor packages/settings
are preserved. Clean native audit through19:31:41.1431443UTC. Pins were captured
after cold compilation and before focused GREEN.

Final `admin-users-recording-static-02`:342 components,6428 procedures,140068 lines;
all28 oversized-module caps hold, dynamic calls remain9/45, all three JSON schemas
and504 PowerShell parses pass. The Users form shrinks; the observer is75 lines.

**Bounded duplicate exception:**196 groups versus195 adds only the three one-line
native button delegates to one shared implementation
(`duplicate:d4f38cb98346406f:mBtnDeactivate_Click+mBtnDelete_Click+mBtnSave_Click`).
Their distinct operation literals are erased by the scanner. The previous reviewed
Core Deactivate/Delete adapter exception remains, with hash2273cf884b026faf after
adding the outcome argument. This review does not waive any other duplication,
size or dynamic-call limit. No unrelated code was removed to satisfy a metric.

## Retained regressions

All prior ordered checks pass on this same pinned candidate, with settings,
packages and tracked reports restored and clean delayed Excel audits:

- Smoke86: `admin-users-recording-regression/smoke-3c54391f944e4942ab3b0d658ca21a51`;
  both host exits are unassisted.
- Settings202: `admin-users-recording-regression/settings-b53a1f07633044cb8435f4279f36cb59`;
  worker `slice4be-tracking-settings/559c18aaecc54c34bc01912acf0f046e`.
- Seed114: `admin-seed-outcome-controller/d5b4247e57f74e028742627d05b28bf7`;
  worker `slice4be-admin-seed-outcome/12d2d490b9f14a0c8d10d4a1ba2aa7c0`.
- Console118: `admin-console-controller/c26a679a3d8742ea852456f4aba854c8`;
  worker `slice4be-admin-console/04c98f708203455b81d8d05d2b31e472`.
- Chain32/live-role48/Create Warehouse15:
  `admin-users-recording-regression/chain-2d05ca9193754fdebd79e05fd63e45ea`.
  All cleanup waits complete without intervention. The final audit through
  20:00:09.5420332UTC has zero Excel failure events.

`admin-users-recording-qualification-01/shared-verification.json` and the smoke/
chain controller `verification.json` files contain exact comparisons. This run
requires catalog38/160 in the existing catalog-sensitive checks; their earlier
check names and assertions remain. Static evidence also regenerated in the chain.

## Remaining acceptance

Context/reentrancy/storage failure observations,
continuous recording, exact Viewer publication, independent saved-guide conclusions
and native control acceptance remain open. The shared expected-conclusion editor
currently omits SAVED; protect the actual editor with RED before adding that choice.
Other Users & Roles controls and existing layout defects remain pending.
D18-LOCAL-01 and D16-VBA-LOAD-01 remain unapproved. No local recording exclusion,
new replay adapter or full4be-A acceptance is claimed.

Commands:

```powershell
# Qualified RED on the unchanged authority package; focused mode only.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File tests/tooling/Test-Slice4beAuthReadOnly.ps1 -DeployRoot deploy/validation-admin-users-authority-01 -Phase RED -AdminRefresh -AdminWriteDiagnostic -RequireWriteAuthority -CoreAuthority -UserRecording -UserRecordingOnly
# Expanded GREEN retains the full authority and visible gate.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File tests/tooling/Test-Slice4beAuthReadOnly.ps1 -DeployRoot deploy/validation-admin-users-recording-01 -Phase GREEN -AdminRefresh -AdminWriteDiagnostic -RequireWriteAuthority -CoreAuthority -UserRecording -CaptureEvidence
```
