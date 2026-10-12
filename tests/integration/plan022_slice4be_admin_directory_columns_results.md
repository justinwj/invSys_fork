# View Warehouses: D14 local-column preservation

4be-A prerequisite, 2026-10-12. Governing rule: Architecture v4.11 D14
header-extension rules and Warehouse Directory enforcement clarification;
Plan 022; Controls 1.547. This adds no recording ID, warehouse identity,
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

The initial narrower `Test-Slice4beProjectionLiveControl -Cut AfterProjection
-Setup Chain -TraceBoundaries` stops during probe installation, before business
checks: `modInventorySchema.EnsureTableWithHeaders.7` has a missing/ambiguous
anchor. Root `projection-live-control/b4fd39bcecd84d099ee747f8e33bc862` records
0 business passes / 1 harness failure, clean closure and zero native events
through `2026-10-12T14:51:40.9293918+00:00`, with restoration and pins preserved.
This is not product RED. The case-sensitive matcher rejected VBA's recased
anchor. The tooling correction and subsequent native comparisons follow below.

## Focused native comparison: tooling unblocked, product failure unresolved

`Test-ProjectionTraceAnchors.ps1` calls the actual trace installer against fake
CodeModules and stops before writing a probe. Exact/case-only spelling must pass
preflight; missing, duplicate and differently cased duplicate anchors must refuse
before mutation. Tooling RED **7/4** at `projection-trace-anchors/30044594bab74800921de652cdb5601d`
becomes **11/11** at `53e36e27d609481a9a99cccc92485966` after changing the matcher
to `OrdinalIgnoreCase`. This is tooling RED/GREEN, not product RED. Existing
shutdown-trace protection remains **74/74**, root
`shutdown-trace-contract/9b06cf88bacb4f39b08b560b0534b6b8`.

All following controls use the same pinned directory-01 packages and actual
ordered validator, `AfterProjection` cut with `Chain` setup. A passing cut retains
the exact first35 checks; the remaining13 live checks are not executed.

| Control / root under `reports/runtime/projection-live-control` | Observed result |
| --- | --- |
| VBA boundaries: `bb0af0b955b44bc5af6b3bbaf01612c4` | 35/35; 37 anchors installed, four instrumented compiles,77 allowlisted markers. Normal closure, clean audit through14:58:25.2002304UTC. |
| Uninstrumented: `fe7ddb511a7c4ba0954de9acbcb89506` | 32/1; processor RPC0x800706BE after deletion, no return marker. Matching Excel owner fails at15:00:57.9814448UTC with ntdll/c0000028. |
| Read-only package/reference inspection: `60b71a9fa69047e0bcc8e87595864432` | 35/35; four executing packages and managed references match the candidate, none broken. No explicit compile, source edit or save. Normal closure, clean audit through15:09:11.8643105UTC. |
| Temporary source-loading option, otherwise uninstrumented: `c85c5985222e4557aedef669d8c93829` | 35/35; normal closure, clean audit through15:15:22.1392592UTC. Original registry-value absence restored after closure. |
| Ordinary control after restoration: `473f12d2c83141dc9dd831ee6ad4b9ac` | 14/1; fails earlier at Receiving Confirm Writes, HRESULT0x80020009. No projection markers and no Application1000/1001/1002 events through15:18:38.6353286UTC. |

All five restore settings/source reports and preserve five package pins and the
tracked live report. The uninstrumented failure's owned `/restore` process required
two normal WM_CLOSE requests; no force-kill or COM reattachment. Its delayed audit
through15:02:36.3303121UTC records two native events. All three passing controls
close unassisted. Inspection can affect timing, COM-reference lifetimes or internal VBA
state: it is not proven inert. These observations do not establish a native cause
or repair, and none is a full-chain acceptance run.

The source-loading hypothesis comes from Microsoft's historical
[Excel2016 MSI compiled-stream workaround](https://support.microsoft.com/en-us/servicing/office/update/excel/august-8-2017-update-for-excel-2016-kb3085435).
That article does not establish the cause in the current Excel build; its old
update was not installed. With Office closed, the diagnostic temporarily set
`HKCU\Software\Microsoft\Office\16.0\Excel\Options\ForceVBALoadFromSource` to
DWORD1, then restored its previously absent state at15:15:22.7592438UTC. Only that
value changed; prior state/restoration are recorded under
`reports/runtime/projection-source-load/c9372bfe14dd41d7aa2cacfa80f8de80`.
No permanent station setting, package change or deployment requirement is adopted.

The post-restoration failure is not the same proven native signature: its original
Excel process exits and a new Excel child appears, with the captured original
owner as its parent but without a `/restore` switch. Exact child creation identity
and parent ownership were checked before one normal WM_CLOSE; no COM attachment
or force termination. `child-provenance.json` and `recovery-close-request.json`
retain the distinction. This control never reaches projection recovery, so it
cannot establish that source loading prevents that particular fault. General
execution failure remains unresolved; do not turn the setting into a deployment
requirement or credit instrumented success as ordinary-chain acceptance.

Static02 retains337 components/6412 procedures/139745 lines, all28 caps, dynamic9/45
and194 duplicate candidates. Three schemas and493 PowerShell files parse. Exact
verification is `reports/runtime/admin-directory-static-02/ratchet-verification.json`.
The older trace01 source copies were sanitized after execution to remove the legacy
fixture credential assignment; `retained-source-redaction.json` records before/after
hashes. Their lifecycle evidence is unchanged, but retained copies are not the
original executed bytes. The focused controller transforms credentials before writing.

Evidence reuse is scoped by `admin-directory-build-01/reuse-scope.json`: all
106 other procedures in `modAdminConsole` are unchanged against `504f0de2`;
only `RefreshWarehouseDirectory` and `AddWarehouseDirectoryRow` change. Together
with 328 identical compiled components, this retains the preceding Settings202,
Console188, Verify202 and Reconcile360 guide baselines at their original scopes.
The fresh Console118 protects the modified module's Console behavior. These are
not fresh full guide/Settings runs on directory-01; exact original receipts remain
in `plan022_slice4be_admin_console_results.md` and its referenced receipts.
