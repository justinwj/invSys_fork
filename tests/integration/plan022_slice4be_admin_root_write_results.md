# Add Warehouse Root persistence feedback diagnostic

4be-A / Admin prerequisite before D18 recording,2026-10-12 Pacific.
Unchanged, unpromoted `deploy/validation-directory-recording-01`; catalog39/161.
Architecture v4.11 **D18-ROOT-01 is proposed, not approved or implemented**.

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

Proposed D18-ROOT-01 makes successful write plus exact readback the persistence
fact and gives unconfirmed writes a combined warning/directory result. It retains
the existing refresh attempt and existing success/cancellation behavior. Approval
is required before runtime changes; recording outcomes remain unspecified.

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

## Preservation and continuation

No VBA, form, Ribbon, schema, package or deployment change. Existing directory219,
Console188, Settings202, smoke86 and chain32/live48/Create15 remain the preceding
qualified product baseline. Test-only additions are gated off in ordinary runs.
Product static baseline remains343 components/6432 procedures/140171 lines,
all28 caps, dynamic9/45 and duplicates196; no new runtime scan is needed for this
diagnostic-only increment. All515 PowerShell scripts parse and diff checks pass.
Raw evidence stays ignored under `reports/runtime`.

Next: obtain D18-ROOT-01 approval, expand focused RED for the approved failure
feedback/readback contract, then implement and qualify it before adding recording.
