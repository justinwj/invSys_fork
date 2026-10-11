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

The continuation below supersedes the earlier untraced-owner limitation. No native
failing API or repaired cause is established. Retain General871, smoke86 and UOM228
at their documented scope; broader Settings, Operations/Admin and final visible
acceptance remain open.

## Live process and delayed teardown localization (2026-10-11 UTC)

The test generator now records the live child's exact process identity and
before/after markers for all51 original projection-recovery statements. Tooling
checks pass **63/63**, retaining the previous56 and verifying unchanged live calls,
unchanged statements, paired boundaries, metadata-only mapping and single
instrumentation. The duplicate-instrumentation guard was first observed failing
(62 PASS/1 FAIL), then corrected; this is tooling evidence, not product RED.

The unchanged General02 full chain runs **17:53:49.1682882--17:56:47.0073936Z**.
It repeats **chain5 PASS/1 FAIL, live32 PASS/1 FAIL, warehouse15/15**. Reached original
checks retain their order;27 chain and16 live checks are unreached. At
**17:55:27.5805077Z**, Application1000 identifies the exact `LiveRole.excel` owner
with **ntdll.dll/c0000028**. Markers1--30 returned; marker31 entered
`modProcessor.RunBatchReportForAutomation` and never returned. Projection deletion,
inbox-trigger save and context restoration had completed. This identifies the
packaged call, not the failing statement inside VBA. It agrees with the earlier
projection-boundary evidence at older candidates; it is not a new failure class.
Later restart/evidence stages are not reached. Recovery closure is assisted,
with recovered files retained. Settings, reports and all frozen packages are preserved.

The existing `RestartProjectionReplay` control then copies this failed generated
fixture, rebinds cloned paths and runs the original six checks in fresh Excel.
It records **3 PASS/3 FAIL**: missing-projection setup, rebuilt tables and duplicate-free
replay pass; one-event application, exact identity and processed status fail.
Those failed checks are unresolved and invalidate a successful replay claim.
The processor returns at17:57:45.4441165Z. Quit returns at17:57:46.3996515Z with zero
open workbooks, but Excel remains throughout210 seconds of observation despite
the existing collection control. The worker exits at18:01:19.6987236Z; the same
owned Excel process faults **167.9 ms later**, at18:01:19.8666213Z, with
**combase.dll/c0000005**. This reproduces the original signature after worker exit;
it does not retrospectively identify the original chain's untraced owner or prove
a native root cause. Recovery closure is assisted and retains files. All13 source
fixture files/hashes, local settings and frozen packages are preserved.

Exact ignored receipts:

- Pure tooling: `shutdown-trace-contract/2dec1d366c87463fbe9950a41d93d9c0/checks.json`.
- Full chain: `general-native-traced02-regression/chain-eb5a47f750314d7a8ec45d9e9697bfe6/native-verification.json` and `qualified-counts.json`.
- Call map: `general-chain-native-trace-02/localization.json`, `projection-map.json`, `owners.jsonl` and `phases.jsonl`; recovery images retain the selected keep-files option.
- Fresh control: `slice4be-shutdown-control/7e4818bed2eb4401ac148500852f4aa5/`; `native-projection-replay02-control/qualified-verification.json` and `fixture-verification.json`.
- Static: `general-native-live-trace-static-01/ratchet-verification.json`; all prior runtime metrics/caps unchanged, three schemas and455 PowerShell parses pass.

The existing unsaved VBA-boundary control then passes **35/35**, retaining the
original ordered live prefix on the same packages. It uses `Setup Chain`, installs
37 stage markers and records77 entries across33 distinct stages. Four instrumented
compiles pass; the processor returns, Excel closes unassisted and the delayed
audit has zero Application errors. Settings, reports and five packages are preserved.
Receipt: `projection-live-control/b362597a0dce4ffe8c9c652c83a5e4a5/qualified-verification.json`.
This control changes VBE preparation and instrumentation; its success does not
explain or repair either native failure.

Next add redacted processor counts, trigger status/error classification and
authority deltas to the fresh control to distinguish its business failures from
delayed teardown. Preserve both native failures and all original checks; no
unchanged full-chain retry or collection-based runtime repair is justified.
