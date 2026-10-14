# Tester Setup cleanup authority diagnostic

4be-A prerequisite,2026-10-13 Pacific. **RED; D8-TESTER-01 approved, not implemented.**
The user approved all six tightened requirements on2026-10-13 Pacific. Expanded
RED and implementation remain required; do not request approval again.
Unchanged qualified package: `deploy/validation-create-recording-02`, catalog41/172.

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File tests/tooling/Test-Slice4beAdminTesterSetupOwner.ps1 -DeployRoot deploy/validation-create-recording-02 -Phase RED
```

Controller `admin-tester-owner-controller/077e6f9d746c458ab37b5eeb0c904de6`:
**25 PASS/4 expected FAIL**, exit1. All five instrumented packages compile.
The actual Admin Ribbon opens `frmSetupTesterStation`; its actual Delete Generated
handler removes the disposable file after sign-out, Reader sign-in, switching to
another warehouse's Admin session, and sign-out inside the confirmation prompt.
The test verifies those states independently. Authorized deletion, Cancel,
confirmation counts, unrelated canaries and both source Config hashes pass.

The canaries are deliberately ordinary files with recognized names, created only
inside the controller's verified temporary root. No operational workbook is a
cleanup target. This also demonstrates that generated ownership is not checked;
it is not proof of successful provisioning or valid workbook contents.
Adapters expose handlers, make presentation modeless and supply confirmation input.
They do not stub deletion results. A test-only diagnostic sink discards raw report
payloads; this run establishes no production logging/redaction acceptance.

Known cause: the form calls `modTesterSetup.DeleteTesterStationGenerated` directly.
That service accepts the TEST/TESTER name heuristic and builds candidate filenames,
without a live source-session check or generated-ownership receipt. Source audit
also finds Setup directly provisioning/seeding an entered existing runtime; Setup
has not been exercised by this focused diagnostic. D8-A preserves explicit station
provisioning, while D8-CREATE-01 applies to Create Warehouse only. **D8-TESTER-01**
was therefore proposed in v4.11 and is now approved before implementation.

Normal closure restores settings and preserves all package/tooling pins. Delayed
native audit through2026-10-14T00:09:04.292UTC finds zero Excel1000/1001/1002 events
and Excel closed. The cursor diagnostic reached/restored its target; no error5
or input lock was observed. Ignored receipts contain only scoped runtime evidence.

Product source and XLAMs are unchanged. Reuse the preceding Create qualification;
no new compile/layout/full-chain/static GREEN or release acceptance is claimed.
Next, expand packaged RED for Setup, verified disposable ownership,
partial/reopened cleanup, open/dirty/foreign artifacts and diagnostic redaction;
then implement and run the required GREEN/regression gates before recording coverage.
