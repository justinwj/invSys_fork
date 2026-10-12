# View Warehouses: D14 local-column preservation

4be-A prerequisite, 2026-10-12. Governing rule: Architecture v4.11 D14
header-extension rules and Warehouse Directory enforcement clarification;
Plan 022; Controls 1.546. This adds no recording ID, warehouse identity,
annotation association, or business authority. Full 4be-A remains incomplete.

## Contract and implementation

The actual `btnAdminWarehouses` / `modAdmin.Open_WarehouseDirectory` callback
prepares the original local workbook, refreshes its directory and presents the
existing result. Previously, `RefreshWarehouseDirectory` deleted every ListRow,
erasing unknown local text. It now clears only managed fields and reuses rows
through `modWarehouseDirectoryRows`. Unknown cells stay at their positions,
including rows beyond the new result count. Discovery order and deduplication
remain unchanged. Empty discovery clears stale managed fields and retains the
existing no-results message, even when local cells remain.

The focused test uses actual packaged Ribbon dispatch and owner handlers.
Only root discovery is confined to Admin-generated fixtures; messages are
observed without dialogs. It does not prove the complete scan-root resolver,
native Ribbon input, or result-dialog layout. No saved warehouse values are
included in committed evidence. File hashes protect both generated authorities.

## RED evidence

Frozen baseline: `deploy/validation-admin-console-03`, compiled pins in
`reports/runtime/admin-console-build-03/compiled-pins.json`.

- Controller `admin-directory-columns-controller/2e41f9f044424de08a9415e1b9bc37de`;
  worker `slice4be-admin-directory-columns/079a471193c94838bc9b065cc152541a`:
  **19 PASS / 2 FAIL**. Only `Directory.After.Text` and `.Tail` fail. The
  formula survives Excel's calculated-column behavior; it is not a failing claim.
  Five instrumented compiles pass. Normal preserved closure; delayed native audit
  through `2026-10-12T14:23:23.3438005+00:00` finds no Excel failures.
- Expanded controller `7a2c6897668f4c73bb4398c6da592c30`; worker
  `1b9124c30c554ec7a7cd5a2dc5171478`: **26 PASS / 4 FAIL**. The same text/tail
  failures also occur after empty discovery; managed clearing, existing message,
  restored discovery, headers, formula, notes and saved authority pass. Five
  compiles, pins, settings restoration and normal closure pass; delayed audit
  through `2026-10-12T14:26:33.5899690+00:00` finds no Excel failures.
- Excluded expanded attempt `082ad14de3554e8da463afe0c9ce4e35` / worker
  `e9e5fd829af54fce850cf8153ea44872` also had an off-by-one test comparison
  against the no-results message. That assertion failure is not product RED;
  the corrected run above supersedes it. Closure preserved packages/settings.

## Candidate and remaining qualification

Unpromoted `deploy/validation-admin-directory-01`: five cold compiles and
Operations cold start pass. `reports/runtime/admin-directory-build-01` records
one changed component (`modAdminConsole`), one added helper, no removals and
328 identical compiled components against console-03. The predecessor remains
unchanged; build settings restored and Excel closed normally. Delayed audit
through `2026-10-12T14:27:34.4797032+00:00` finds no Excel failures.

Focused GREEN controller `e6d9de617ea940629af58d6e084b866b`, worker
`2e1f1b7d953c493aa79e51a5eb0cddd7`: **31/31**, retaining the exact 30-check RED
prefix. Five instrumented compiles, package/tooling pins, settings restoration,
saved authority and normal closure pass. Delayed native audit through
`2026-10-12T14:28:53.3763209+00:00` is clean. Reviewed
`directory-local-columns.png` shows the retained note, formula result and tail
value; managed paths are hidden in this generated-fixture capture. This is local
cell visibility proof, not a new form layout or native Ribbon-input claim.

Fresh Admin Console controller `admin-console-controller/eeeb7e85cc2c438db4ee8602df086bec`
passes **118/118**, preserving the exact console-03 check order. Five instrumented
compiles, pins, restoration and normal closure pass; delayed audit through
`2026-10-12T14:32:58.8973205+00:00` is clean.

Static `reports/runtime/admin-directory-static-01`: 337 components, 6412
procedures, 139745 lines; all 28 existing module caps hold without growth.
Dynamic calls remain 9 literal / 45 unresolved; duplicate bodies remain 194.
Three JSON schemas and all 492 PowerShell parsers pass. No new exception.

Packaged smoke controller `admin-directory-regression/smoke-1500b59f4d984081a29aa709a1549a97`
passes **86/86** with exact prior check order, both unassisted Excel exits,
preserved package pins, restored settings/reports and a clean delayed native audit
through `2026-10-12T14:33:32.6075568+00:00`.

The first chain is **not accepted**: controller
`admin-directory-regression/chain-ad621256fa0b45459756f8595b2e2767` records chain
5 PASS / 1 FAIL, live-role 32 PASS / 1 FAIL and Create Warehouse 15/15. The live
gate reports RPC `0x800706BE` at inventory projection rebuilding through
`modProcessor.RunBatchReportForAutomation`. Windows records an Excel crash at
`2026-10-12T14:35:25.5044874Z`, module `ntdll.dll`, exception `c0000028`;
that module is failure evidence, not a proven root cause. The observed recovery
process had `/restore`. Two normal WM_CLOSE requests closed its windows, without
COM attachment or force termination. This is assisted recovery closure, not
unassisted acceptance. The original controller remained alive through restoration;
settings, package pins and tracked reports were preserved and Excel closed.
Desktop monitoring continued with zero error 5. No native-crash repair is claimed.

The second unchanged-candidate chain, controller
`admin-directory-regression/chain-833ce42b13d24d67be9d6d397e7e58f3`, reports
**32/32, live48/48, Create15/15**, but is also **not accepted**. Its native audit
records Excel `combase.dll/c0000005` at `2026-10-12T14:42:36.7639138Z`.
All assertions passing does not satisfy the clean native-lifecycle gate. This
run closed without our intervention and restored settings, packages and reports;
the exact failing owner/phase is not established by these untraced reports.

Next diagnostic uses the existing metadata-only chain owner/phase trace, with
fresh pure tooling **74/74** at `shutdown-trace-contract/6e4d772f3c8c496d854fcc2e7e0a2739`.
It does not change product runtime or contract. Trace generation is in
`reports/runtime/admin-directory-native-trace-01`; controller
`admin-directory-native-regression/chain-157975d59bcc4da8acd9a03cc87014f6` again
ends at chain5/1, live32/1, Create15/15. Owner/start-time matching attributes the
`2026-10-12T14:48:26.6774796Z` ntdll/c0000028 fault to `LiveRole.excel`, after
`Projection.30` and inside `Projection.31` / `modProcessor.RunBatchReportForAutomation`.
No return boundary is recorded. This localizes the failing process/call after
deletion, not the exact VBA cause. Two normal WM_CLOSE requests close the owned
recovery process; no COM reattachment or force termination. Settings, reports and
package pins are restored/preserved. The delayed audit records two native events;
this diagnostic is not acceptance.

The narrower existing `Test-Slice4beProjectionLiveControl -Cut AfterProjection
-Setup Chain -TraceBoundaries` stops during probe installation, before business
checks: `modInventorySchema.EnsureTableWithHeaders.7` has a missing/ambiguous
anchor. Root `projection-live-control/b4fd39bcecd84d099ee747f8e33bc862` records
0 business passes / 1 harness failure, clean closure and zero native events
through `2026-10-12T14:51:40.9293918+00:00`, with restoration and pins preserved.
This is not product RED. The source anchor `lo.Name = tableName` still exists;
the matcher uses ordinal case-sensitive comparison. Compiled VBA casing is a
hypothesis to verify before correcting the trace tool. Do not repeat broad chains
until the focused diagnostic can reach its intended boundary.

Evidence reuse is scoped by `admin-directory-build-01/reuse-scope.json`: all
106 other procedures in `modAdminConsole` are unchanged against `504f0de2`;
only `RefreshWarehouseDirectory` and `AddWarehouseDirectoryRow` change. Together
with 328 identical compiled components, this retains the preceding Settings202,
Console188, Verify202 and Reconcile360 guide baselines at their original scopes.
The fresh Console118 protects the modified module's Console behavior. These are
not fresh full guide/Settings runs on directory-01; exact original receipts remain
in `plan022_slice4be_admin_console_results.md` and its referenced receipts.
