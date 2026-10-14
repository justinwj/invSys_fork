# Aggregation prerequisite diagnostic

4be-A remaining Admin coverage,2026-10-13 Pacific. **Diagnostic RED, not acceptance.**
Unchanged unpromoted `deploy/validation-tester-recording-02`, catalog42/179;
code baseline `546b029b`, docs `0a21a86`. Proposed D8-AGG-01 in v4.11 requires
approval before product implementation at that diagnostic checkpoint.
**Approval update,2026-10-14 Pacific:** the user approved tightened D8-AGG-01,
including joint source/authority/feed rechecks, durable honest-retry state and the
path-only packaged comparison. The approval blocker is cleared. Expanded RED and
implementation remained pending at approval; local expanded RED is recorded below.
D8-TESTER-01 remains approved/qualified.

## Local approved-contract RED,2026-10-14 Pacific

**Stopped for the user's desktop-lock condition,2026-10-14 at21:32:12Z
(14:32:12 Pacific).** Owned monitor `resume-d8-agg-01/write-boundaries-entry-12`
records150 samples, one lock condition, zero error5/mouse-lock/other failures,
and `Stopped=True`; the macro dispatcher refuses further work. Goal is paused
under the user's explicit conditional stop instruction; D8-AGG-01 remains approved.
Product VBA is still unchanged, and4be-A is incomplete.

After the complete104 PASS/46 FAIL gate below, additional actual-handler probes
protect feedback redaction, same-form retry, post-save failure, pre-save
source/feed/sign-out changes, malformed/duplicate sources, unknown output columns
and destination overlaps. A focused `-WriteBoundariesOnly` selector runs these
write cases and the retained NewSession case without repeating the path comparison.
The duplicate fixture now copies a generated row directly rather than assigning
mixed COM numeric/string values; its behavioral case has not yet been reached.

| Controller / worker under the existing ignored runtime roots | Interrupted result |
| --- | --- |
| `284090309d21490c9035d371ffcb1089` / `5b3ca6761e7f4e1a9df54e9cedbde1db` | 100 PASS/56 FAIL, including a stale-monitor harness failure before NewSession completes; fresh decoy restores the original workbook-switch mutation RED |
| `9e24d501299449848c058b5acdcd0fb4` / `bf14b6b9fd8449658cec401ce9c5858b` | 24 PASS/19 FAIL, including a duplicate-fixture COM harness failure; completed cases expose pre-save changes and malformed-source acceptance |
| `9e1b652822644f26b9b499dad0d2a61b` / `3fff33580ed6494db1a2294c3d2f7a5e` | 19 PASS/11 FAIL, including the required desktop-lock stop; reached failure-before/after-save and source-change-before-save cases |

These are partial evidence, not completed RED gates. Reached failure cases retain
no durable destination/source/phase-bound retry record and allow another write
from the same form. Reached feedback checks expose entered source/feed paths.
No exact persisted-output claim is added for the injected post-save failure.
Unknown-column/overlap, corrected duplicate, interrupted/reopened retry and the
remaining phase checks still require verification before product implementation.
The direct Core baseline adapter remains supplemental, not proof of a new API.

All three controllers preserve package/tooling pins and restore settings, close
Excel normally, and record zero delayed Application1000/1001/1002 Excel events
through21:22:02.3276397Z,21:28:17.1729832Z and21:32:41.5673421Z respectively.
The last two owned monitors exit normally. Sharing-error retries and a numeric
publication-failure record were added only to ignored monitor orchestration;
one healthy short shutdown does not establish long-run monitor reliability.

Product VBA remains unchanged from `ff5db426`. An isolated source-built five-package
baseline, `deploy/validation-aggregation-baseline-01`, replaces the absent prior
candidate for these local diagnostics; it is unpromoted. The actual Ribbon/form
handler remains the entry. Path normalization and fault injection occur only in
disposable test copies, never in product source or operational workbooks.

| Controller under `admin-aggregation-controller/` | Worker under `slice4be-admin-aggregation/` | Result |
| --- | --- | --- |
| `77daae3b04e94f79ad52f9c8f6bcf3a1` | `0a3e518cf9f74a029a16493b43335191` | 74 PASS /10 FAIL; complete retained65/7 plus12 drive/UNC comparisons |
| `a20328abf5b04c74aaa02408fee6512e` | `c36af2f15e584ac4aa848f6f4ca6fa91` | 88 PASS /35 FAIL; adds post-inspection, borrowed-output and failure/retry boundaries; no harness exception |
| `7f1b2815847d466e97e2a3730e755e12` | `6440610f466e485dab3a17496f5e579c` | 104 PASS /46 FAIL; adds saved revocation, dirty/unreadable authority, read-only output and closed original workbook; no harness exception |

The path-only packaged comparison holds the source set, workbook, authority, feed
and package fixed within each variant. Original drive joining fails at SaveAs1004;
normalization alone yields complete persisted output with exact identities/quantities.
Both UNC variants persist the exact data through the existing local administrative
share, but the original dispatched path is noncanonical; normalization preserves
the UNC prefix and satisfies the canonical destination assertion. Source bytes are
unchanged in all four comparisons. This establishes the drive failure's cause in
this fixture, not a universal claim that doubled separators fail on UNC/NAS.

The expanded boundary gate proves the fault actually ran before trusting its result.
After inspection, changing/replacing or deleting a selected snapshot, changing the
saved trusted feed, or signing out still reaches the old writer. Changed/missing
sources can report success; a changed feed creates redirected output. Normalization
is enabled only in these disposable boundary copies to expose effects independently
of the known drive joining defect. Clean and dirty borrowed output workbooks are
closed by the old writer; output bytes/saved-state preservation fails. Injected
failure immediately before SaveAs deletes the existing output and leaves no durable
retry record. These are protecting REDs, not acceptance.

The separately labelled direct-Core baseline adapter forwards to the exact primitive
writer reached by the current interactive handler, which has no context parameter.
Its positive control succeeds, and it also persists after sign-out. This supplements
the packaged authority results; it is not the new interactive command's API proof,
does not replace packaged actions, and changes no scheduled/headless authorization.

Both completed controllers preserve all five package pins, tooling and settings,
close Excel normally, and pass delayed Application1000/1001/1002 audits with zero
Excel events. Expanded audit ends2026-10-14T20:38:49.9944419Z. Its owned monitor
records397 samples with zero error5/lock/mouse-lock observations and exits normally.
Earlier controller-entry, argument-count and monitor-report concurrency failures
are harness failures, not behavioral RED. The controller now retains a real child
process handle; heartbeat reads allow atomic replacement with bounded read-only
retries and reject stale/missing reports before any macro dispatch.

The broader150-check gate adds actual saved source-rights revocation after inspection,
dirty Config/Auth, unreadable Config, read-only output and a closed original workbook.
Rights loss and dirty authority still permit writing. A read-only output reaches the
writer even though persistence fails; unreadable Config already refuses. Closing the
original workbook redirects work and mutates the decoy. Its later original
SwitchWorkbook.OtherWorkbookPreserved result is masked by that earlier decoy
preparation; a fresh decoy is now created before the retained original matrix and
must be verified by the next run. No prior mutation assertion is discarded.
This controller preserves pins/settings, closes normally and has zero delayed Excel
events through2026-10-14T20:55:46.1078420Z. Its owned monitor records614 healthy
samples and exits normally. These scoped REDs do not establish product acceptance.

Remaining pre-implementation coverage includes feedback redaction, applicable later
yield/phase and interrupted/reopened retry. No new product compile/layout/static/regression
GREEN, package promotion, Aggregation recording or4be-A completion is claimed.

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
defined the unresolved interactive authority/ownership proposal at the diagnostic
checkpoint; its tightened contract is now explicitly approved, not inferred from
these tests. Extend RED for direct Core denial,
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
