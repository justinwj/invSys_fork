# Add Warehouse Root persistence feedback

4be-A / Admin prerequisite before D18 recording,2026-10-13 Pacific.
Architecture v4.11 **D18-ROOT-01 approved2026-10-13**. Owner correction is
scoped-qualified on unpromoted `deploy/validation-root-feedback-01`; catalog39/161.
Root action recording remains unimplemented; this is not4be-A acceptance.

## Finding and scope

Actual packaged `btnAdminWarehouseRoot` -> `modAdmin.Add_WarehouseDirectoryRoot`
calls `modAdminConsole.RememberWarehouseScanRoot`. That Sub suppresses errors at
SaveSetting; the callback then refreshes the directory and may report
"Warehouse root remembered and directory refreshed." A test-only VBA wrapper
refuses the actual WarehouseScanRoots write with VBA error70, without changing
Windows permissions. Exact before/after preference comparison proves no write,
while the real callback still reports remembered. This is not desktop error5 or
physical registry-outage evidence. No preference value, path or credential is
emitted by the probe; reports contain booleans/counts/check IDs only.

Approved D18-ROOT-01 makes successful write plus exact readback the persistence
fact and gives unconfirmed writes a combined warning/directory result. It retains
the existing refresh attempt and existing success/cancellation behavior.
Recording outcomes remain unspecified.

## Qualified RED

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File tests/tooling/Test-Slice4beAdminDirectoryColumns.ps1 -DeployRoot deploy/validation-directory-recording-01 -Phase RED -RootWriteDiagnostic
```

Controller `admin-directory-columns-controller/c145dd397cdf4ec9a114a7f3112447d0`,
worker `slice4be-admin-directory-columns/8fe9aeb7d6ae47afa8f03b5ea64014a1`:
**53 PASS / 1 expected FAIL**, solely
`WarehouseRoot.WriteFailure.NoFalseRememberedMessage`.
UTC2026-10-13 05:01:36.770--05:02:32.640; clean delayed native audit through
05:02:52.640. Five instrumented compiles, exact log/JSON, frozen compiled pins,
tooling/settings/package preservation and normal Excel closure pass.

All30 preceding directory checks remain in order. Success and blank cancellation
pass; write refusal is reached exactly once. The existing single directory action,
original workbook, custom columns, other workbook and saved generated warehouse
authority are preserved. This diagnostic characterizes existing feedback as well
as the defect; after approval, expand its expectations for the new warning,
readback failures/mismatches and directory failure before changing implementation.
It is not a completed behavior gate or native InputBox/Ribbon-input acceptance.

Initial excluded controller `31d2086b9317456e88c96ec63c970dfe`, worker
`347194a77e1041c4990c73204c40159c`, stops30/1 because the new test function was
dot-sourced into installation scope and unavailable during execution. Correcting
test scope produces the qualified RED above; no product change was made. Its pins,
restoration and normal closure pass; native audit through05:00:54.725 is clean.

## Approved-contract RED and candidate

Expanded RED on the frozen directory candidate: controller
`admin-directory-columns-controller/fdfebf1c45bb4156834b6db2d9754936`, worker
`slice4be-admin-directory-columns/1a077d530dbf4db9a6efe1712ad41b6e`: **96/15**.
Failures are exactly absent readback and incorrect remembered/warning feedback
across write refusal, read refusal/mismatch and directory failure combinations.
All30 inherited checks pass. Five compiles, exact log, pins, restoration, normal
closure and native audit through15:27:12.434UTC pass. Faults are confined to
unsaved test packages; directory failure is a fixed owner-boundary injection.

Runtime changes are six added lines in `modAdmin` and `modAdminConsole`: return
confirmed persistence only after a successful write and binary-equal readback,
then combine unconfirmed persistence with the single original directory result.
Build `root-feedback-build-01` passes five cold compiles and Operations cold start;
only these two compiled components differ,334 remain unchanged. Settings and frozen
directory candidate are preserved; native audit through15:28:32.343UTC is clean.

Excluded first candidate run: controller `9cdc1ec91651404598e5869d0423e7ed`,
worker `4b506498a1be46a992214577c14bbceb`,111/2. All nine functional cases pass;
the worksheet capture found an owned menu and a new normalization test packed its
arguments incorrectly. Test corrections request foreground before caption fallback
and parenthesize the input expression. No runtime change. Pins, restoration and
normal closure pass; native audit through15:31:33.354UTC is clean.

Second excluded run: controller `ba518a295de94e8a903d941c0bdeb5f7`, worker
`ee169d1f3aef4f0b956de7fad8ac342e`,114/1. Functional/list/cancellation and worksheet
capture pass. The native observer assumed standard OK control ID1; this VBA dialog
uses another ID. Read-only inspection proved the exact owned warning and sole OK
caption. A guarded acknowledgement allowed cleanup; no force termination or product
retry. Native audit through15:35:59.679UTC is clean. The helper now finds the unique
owned OK button while retaining exact process/class/title/body checks.

## Qualified focused GREEN

Same command with root-feedback-01 and `-Phase GREEN`: controller
`admin-directory-columns-controller/b95303aad9c248a4ad3cc1dbb403f96f`, worker
`slice4be-admin-directory-columns/5650556ff4df4fec80e7c75b546e4168`, **120/0**.
All111 expanded RED check IDs remain in order and pass. Additional checks cover
normalization, case-insensitive deduplication, newest-first/eight-root ordering,
whitespace cancellation and native warnings after successful/failed refresh.
The three captures were reviewed: complete two-line warnings, warning icon/OK,
and preserved local columns. Prompt input is a test seam; native Ribbon/InputBox
entry is not claimed. Five compiles, exact log/JSON, tooling/package pins,
settings/authority restoration and normal closure pass; native audit clean through
15:38:20.154UTC. Readback failure cases prove actual persistence may have happened;
the warning does not claim rollback or unchanged settings.

Fresh packaged smoke retains86/0 and both unassisted exits on this candidate:
`root-feedback-regression/smoke-605f03d137f74a92a92e2d9638c5c0e8`, clean native
audit through15:39:07.457UTC, settings/packages/tracked reports preserved.

Full Release1 chain `root-feedback-regression/chain-d178ef1f2f2c40f68902f456b789e5c3`
retains exact prior32 chain/48 live-role/15 Create Warehouse checks, all passing.
Normal closure, settings/package/report restoration and native audit through
15:45:52.647UTC pass. Fresh neighboring directory recording retains exact149/0:
controller `admin-directory-columns-controller/84f6f06051cd408c98847324cb4b324b`,
worker `slice4be-admin-directory-columns/307a24e1bce24e72902e2751e7e20814`.
Its five compiles, reviewed local-column capture, pins, preservation and normal
closure pass; native audit is clean through15:49:59.402UTC.

Static `root-feedback-static-01`:343 components/6432 procedures/140177 lines,
six added lines, all28 prior caps preserved, dynamic9/45 and duplicates196
unchanged; three schemas and516 PowerShell parses pass. Build comparison proves
334 other compiled components unchanged; `unchanged-procedures.json` additionally
proves every procedure outside the two approved owner changes is identical.
Earlier Settings202, Console188/six captures and independent directory guide219
remain scoped reuse from the directory receipt, not fresh runs on this candidate.
The new dialogs have native feedback/layout evidence; InputBox/Ribbon entry and
root recording/guide proof remain outside this owner correction.

## Prior preservation evidence

At the original diagnostic checkpoint there was no VBA, form, Ribbon, schema,
package or deployment change. Existing directory219,
Console188, Settings202, smoke86 and chain32/live48/Create15 remain the preceding
qualified product baseline. Test-only additions are gated off in ordinary runs.
Product static baseline remains343 components/6432 procedures/140171 lines,
all28 caps, dynamic9/45 and duplicates196; no new runtime scan is needed for this
diagnostic-only increment. All515 PowerShell scripts parse and diff checks pass.
Raw evidence stays ignored under `reports/runtime`.

Next: define and protect Add Warehouse Root recording under D18 semantic
inheritance using confirmed owner results, preserving this120-check owner gate.
