# Slice 4be-A native-failure localization

Last verified: 2026-10-11 UTC. Diagnostic tooling only, under Architecture v4.11
D13 and Plan022. Runtime remains `6f583238`, catalog31, with the five frozen
`deploy/validation-general-settings-02` packages unchanged. No contract amendment,
runtime repair, clean full-chain gate or 4be-A completion is claimed.

## Findings

The original General02 chain's `combase.dll/c0000005` fault remains unresolved.
Its process started after the live-role report; exact owner/call was not captured.
New instrumentation identifies test-owned Excel instances by process ID plus
creation time and records fixed phase/macro/check identifiers without arguments
or workbook values. It preserves the original chain and its business calls.

| Diagnostic | Result and limit |
| --- | --- |
| Isolated original restart | 12/12 retained checks. Quit returned with zero workbooks, but the owned Excel process remained for the 210-second passive observation and exited after the worker ended. No native fault. The diagnostic's lifecycle criterion failed. |
| Restart with existing collection control | 12/12; Excel exited while the original worker remained alive; no native fault. This establishes a lifetime difference, not a crash cause or approved repair. |
| Instrumented full chain | Earlier native `ntdll.dll/c0000028` fault during live-role projection recovery. Chain 5 PASS/1 harness failure; live-role 32 PASS/1 harness failure; Create Warehouse 15/15. Later restart/evidence stages were not reached. |
| Late stages, first clone | 15/15 retained tail checks, four traced Excel owners, no native fault; **fixture isolation failed** because three source inbox files changed. |
| Late stages, corrected clone | 15/15, exact prior tail order; all 13 source files and their file set preserved, settings/packages preserved, unassisted exit and zero delayed Excel failures. The earlier 17 full-chain checks were not executed. |
| Final collected restart, corrected clone | 12/12, exact prior order; source file set/hashes, settings and packages preserved; two cloned path cells initialized, Excel exited while the worker was alive and delayed audit was clean. |

The full-chain fault occurred at **08:48:52.5688535Z**, in the live child created
at **08:47:49.1973544Z**. Last successful check:
`InventoryDomain.ProjectionRecovery.Delete`; failing step: **Delete and rebuild
canonical inventory projections**. This bounds that run's failure; it does not
identify a failing native API or prove a relationship to the original combase
fault. Recovery closure was **assisted** through the native dialog, retaining
recovered files for later viewing. No second Excel COM attachment was used for
recovery. Tracked reports and settings were restored; frozen packages were unchanged.

## Fixture correction and protecting evidence

The first clone retained source `tblWarehouseConfig.PathDataRoot` and
`PathBackupRoot`. `tblStationConfig` had no `PathInboxRoot`; the processor's
`BuildConfiguredInboxPathProcessor` therefore used configured DataRoot. A Core
override alone did not redirect inbox paths. Production, Receiving and Shipping
inboxes in the **generated disposable source fixture** changed. Earlier isolated
controls did not verify source hashes and must not be credited with that proof.

The diagnostic now validates all configured paths, maps only source-root paths
to the clone, resolves columns by header name, and saves only the cloned config.
Unknown columns are untouched by setup; outside-root paths are refused. This is
a test-isolation correction, observed failing before the correction and passing
afterward, **not D13 product behavioral RED/GREEN**. No operational workbook was
used as the fixture. The earlier changed test files have not been silently restored.

`Test-Slice4beShutdownTrace.ps1` passes **56/56** without Excel: original business
calls and return/exception behavior retained, metadata redaction, anchor drift
refusal, path boundaries, copied unknown fixture file, and partial-report scope.
The final report generator removes inherited full-chain/D13 acceptance wording
from partial reports. Historical partial reports retain their original text;
their `check-verification.json` explicitly qualifies that text as non-acceptance.
An absolute-output-path construction error found during tool validation was
corrected; it was a harness error, not product RED.

Final static evidence remains319 components /6340 procedures /138712 lines,
9 literal /45 unresolved dynamic calls,192 duplicate groups and28 unchanged
oversized caps. Three schemas validate and455 PowerShell scripts parse. No XLAM
was rebuilt; prior same-candidate compile/layout evidence keeps its original scope.
The desktop monitor completed375 samples from08:38:42Z to09:10:02Z with zero cursor,
desktop, capture or error5 failures. Excel and the monitor are closed. This does
not establish an RDP/tscon repair.

## Exact receipts

All paths below are relative to ignored `reports/runtime/`; raw runtime reports,
process identities and recovery images are not committed.

- Original fault: `general-recorded-regression/chain-d09373c0a47a4468b269a2e9f887eff0/qualified-verification.json` and `native-timing.json`.
- Passive restart: `slice4be-shutdown-control/84d9e3435606425d8f97961ed1b6a0d4/result.json`, `check-verification.json`, `delayed-audit.json` and `phases.jsonl`.
- First collected control: `slice4be-shutdown-control/b722004343fe44ffbbf2e070d8f0c161/result.json`, `check-verification.json` and `delayed-audit.json`.
- Full trace: `general-native-traced-regression/chain-1b7a92fe8b9746838d79783484d37d8e/native-verification.json`; `general-chain-native-trace-01/live-failure-summary.json` and recovery evidence.
- Failed isolation: `general-native-late-regression/chain-c9219365af5649e08feba642a3cb0ea9/native-verification.json`; `general-chain-native-late-01/source-fixture-change.json` and `check-verification.json`.
- Corrected late stages: `general-native-late02-regression/chain-eaddcb36ffa640a9bbae1590ec4c8e41/native-verification.json`; `general-chain-native-late-02/check-verification.json`, `fixture-paths.json`, `owners.jsonl` and `phases.jsonl`.
- Final collected control: `slice4be-shutdown-control/b790571fde1f4cfea2541c80273f5daa/verification.json`; `native-final-control/fixture-verification.json`.
- Pure tooling checks: `shutdown-trace-contract/9fd0fcd1a6054ba0819d0927ea6308f0/checks.json`.
- Static: `general-native-localization-static-03/ratchet-verification.json`.
- Desktop: `rdp-desktop-probe-326f94864c23433fae8df9a428aa2dca/summary.json`.

## Remaining diagnostic boundary

No native failing API or repaired cause is established. Keep the original fault
and the earlier live-role fault separate. Before another full-chain attempt,
add owned-process and before/after markers to the live child's projection
deletion/rebuild calls, retaining the late-stage trace. A full-chain retry must
preserve all original checks and qualify native failures independently of counts.
Do not convert the collection control into a runtime fix without a protecting
reproduction. Retain General871, smoke86 and UOM228 at their documented scope;
broader Settings, Operations/Admin and final visible acceptance remain open.
