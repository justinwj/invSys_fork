# Aggregation prerequisite diagnostic

4be-A remaining Admin coverage,2026-10-13 Pacific. **Diagnostic RED, not acceptance.**
Unchanged unpromoted `deploy/validation-tester-recording-02`, catalog42/179;
code baseline `546b029b`, docs `0a21a86`. Proposed D8-AGG-01 in v4.11 requires
approval before product implementation. D8-TESTER-01 remains approved/qualified.

## Reproduction and boundary

Run `tests/tooling/Test-Slice4beAdminAggregation.ps1 -Phase RED` with Excel closed.
The isolated worker generates two warehouses through existing Admin owners, seeds
and publishes snapshots, then opens `btnAdminHqAggregator` through actual Ribbon
dispatch and invokes `frmAggregationSources.mBtnAggregate_Click` through a test
entry. Discovery roots are restricted to generated fixtures; modal presentation
is made modeless in disposable package copies to permit authority/workbook changes.
Source discovery, permission checks and aggregate results remain real. Instrumentation
records only fixed stages, counters, booleans and numeric errors; no result stubs.

Source Config/Auth/inventory/published-snapshot bytes are pinned. Output comparisons
use case-sensitive WarehouseId/System_Key identities and quantities in memory;
no rows, credentials, entered paths or raw owner reports enter committed evidence.
The supplemental direct Core control supplies a normalized output path. It is not
a replacement for the failing packaged action and is not interactive authority proof.

## Observed results

All paths below are under ignored `reports/runtime/`.

| Controller under `admin-aggregation-controller/` | Worker under `slice4be-admin-aggregation/` | Result |
| --- | --- | --- |
| `119d2433364f4befa5d7510cd10ca9b4` | `0382dbbf3bd046b2a925c87c96022078` | 58 PASS /4 FAIL; authorized output and workbook binding fail |
| `647771951593473ab85e4a5ee3466a29` | `e6b5b323e0b442fabc711b1a65d955b8` | 58 PASS /4 FAIL; fixed phase/error observation added |
| `9013978937f7481599945845ac71589b` | `b94a0a92885147a7a2d9ddc6fc2e58df` | 65 PASS /7 FAIL; Core dispatch/stage and normalized-path control added |

Final failures: Admin.ExpectedWriteBoundary; SwitchWorkbook.OriginalWorkbookBound,
ExpectedWriteBoundary and OtherWorkbookPreserved; SignedOutHint, OtherWarehouseAdmin
and NewSession.AuthorizedCoreDispatchOnly. All five instrumented projects compile.

- An active-workbook switch makes the actual owner receive the unrelated workbook;
  its worksheet structure changes. The original binding is not preserved.
- SignedOutHint has no live sign-in/capability; OtherWarehouseAdmin has no capability
  for the original warehouse. Both reach the actual Core writer. A newly signed-in
  session also reaches it from the old form. Plain SignedOut and Reader stop at the
  provisioned-permission check. Every precondition is checked before dispatch.
- All five cases reaching Core stop at SaveAs with Excel error1004. None produces
  the expected output, including the authorized baseline. **No persisted unauthorized
  write is proved.** Lack of output in stale cases is not authorization GREEN.
- The independent normalized-path Core control succeeds, with exact selected
  entities/quantities and unchanged sources. Admin builds the path using doubled
  interior backslashes; this is a suspected cause of the callback SaveAs failure.
  A path-only packaged comparison is still needed to establish causality.
- Both warehouses' Config/Auth/inventory/published snapshots remain byte-identical
  throughout. This does not erase the unrelated operator-workbook mutation.

Each controller verifies restored settings, preserved five-package/tooling pins
and normal Excel closure. Delayed Application1000/1001/1002 audits find zero Excel
events through06:05:49.332Z,06:09:42.947Z and06:13:57.192Z on2026-10-14 respectively.
Excel1004 is an application error, not desktop Win32 error5. Actual cursor movement
and restoration passed before the run at05:58:44Z; passive monitoring found no error5.
No visible layout or broad Release1 acceptance is claimed by this diagnostic.

## Contract and next gate

D17 governs selected read-only sources and advisory output; D8-AGG-01 explicitly
proposes the unresolved interactive authority/ownership contract. It is pending,
not inferred from these tests. After approval, extend RED for direct Core denial,
saved revocation, yields, closed original workbook, changed output Config,
borrowed/dirty/read-only output, redacted outcomes and source rejection before
implementation. Preserve D17 identity rules, existing headless authorization and
all prior GREEN evidence. No recording control or replay adapter is introduced.

This increment changes only diagnostic tooling and documentation; no product VBA,
form, Ribbon, schema or package is changed. Reuse the qualified Tester product
static/regression evidence, without claiming this diagnostic passes those broader
gates anew. Parse all tooling, inspect diffs/links/status and preserve unrelated work.

Closing verification: all536 repository PowerShell files parse, including475 in
tests/tooling;96 local links in the changed coverage records resolve. Diff whitespace
checks pass. No product source diff exists. Unrelated user-file hashes match the
opening checkpoint, and generated runtime files remain ignored.
