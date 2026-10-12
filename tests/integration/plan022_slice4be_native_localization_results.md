# Slice 4be-A native-failure localization

Last verified: 2026-10-12 UTC. Diagnostic tooling only, under Architecture v4.11
D13 and Plan022. Earlier localization below uses frozen General02, runtime
`6f583238`, catalog31. UI01 subsequently has a clean ordinary chain recorded in
[Run results](plan022_slice4be_production_run_local_results.md). Scale01 also has
a clean ordinary chain and full reusable regression there; neither establishes
a native repair. No 4be-A completion is claimed.

## Current chain interruption and dismissal ownership

Two bounded diagnostics on unchanged `validation-run-scale-success-02` follow the
Shipping failure below. Early attachment stops before Shipping: live10/1 at
`modConfig.LoadConfig`, COM0x800AC472, with first-chance80010108/000006BA but no
Application1000/1001/1002 event. Native root `run-scale-success-native/5d02709c612f437ab3ecff5ce2452f67`,
controller `run-scale-success-regression/chain-e73f2db3f3004ac6b16111e1dafce739`,
UTC00:53:05.812--00:57:38.241. The worker exited; its pinned Excel remained hidden.
The observer detached and native close requests ended the residual process without
COM reattachment or force. Preservation passes; this is a pre-target diagnostic
failure, not Shipping RED or evidence that debugger attachment caused it.

Attachment moved to a fixed before-Sent barrier after capability validation.
The unchanged actual Shipping call returns at01:01:14.427 UTC. Diagnostic root
`shipping-boundary-native/9e82c6a12b8b46efb0633da01472328c`, controller
`run-scale-success-regression/chain-5c71c7e5f1e44270ad921a071b211c82`,
UTC00:59:01.938--01:02:27.679, retains31 executed chain checks, live48 and warehouse15
in prior order. Final static is deliberately unexecuted; authoritative controller
exit1 and StaticNotRun.DiagnosticOnly remain. One first-chance80010108 occurs, no
second-chance exception or delayed native event; shutdown and preservation qualify
through01:02:39.679 UTC. No VBA/business calls change. Observation alters timing:
this does not reproduce or repair the Shipping fault and is not chain acceptance.

The initial ad hoc transformation check had a PowerShell function-scope mistake;
corrected validation proves parsing, exact call retention, both fixed markers and
missing-anchor rejection without rerunning Excel. The outer observer incorrectly
returned zero while its controller correctly recorded exit1; it now reads the
authoritative closure, matching the established observer wrapper. Diagnostic
receipts retain both facts. No product GREEN was inferred from that outer exit.
Local ignored entry points: `reports/runtime/observe-shipping-boundary.ps1` and
`run-native-chain-at-shipping.ps1`; the calibrated observer is unchanged.

The user reported setting up Firefox and Outlook and said they may have affected
the activity. This is context, not confirmed input to Excel or a native root cause.
Keep the barrier for the next planned chain; do not repeat unchanged chains to
chase a pass. Resume independent worksheet Scale guide proof while this defect
remains open. No fresh runtime/static rebuild is required for this evidence-only update.

The D14 successful worksheet reload candidate `validation-run-scale-success-02`
has focused GREEN86 and retained enabled218/missing64/refusal162, but its ordinary
chain fails at `modTS_Shipments.RunShipmentsSentFormActionForTest`, COM0x800706BE.
Application1000 records `ntdll.dll/c0000028` at2026-10-12T00:45:36.699Z. Warehouse15
passes, live-role stops40/1 and chain5/1; both Production batches and Boxing passed
before Shipping. Controller `run-scale-success-regression/chain-0518429fa5d44e81b2063b699d3c5c5e`
retains a sanitized failed-verification receipt, exact candidate pins and preserved
settings/reports. A verified crash-reporter child closed after a normal native close
request; no prompt was observed, no COM reattachment or forced termination occurred.
Closure was assisted and the run is not acceptance. This extends the observed
failing-call set; it does not establish a different root cause or prove the new
Production helper innocent. Do not repeat unchanged chains to chase a pass.

On 2026-10-11, the ordinary Designs-disabled Scale candidate chain stopped at
`mProduction.ProductionFormTwoBatchActionReportForTest` with COM0x800706BE.
Application1000 at22:15:00.304Z records `ntdll.dll/c0000028`; Application1001
followed. Warehouse checks passed15/15, live-role35/1 and chain5/1; later stages
were not reached. The macro exercises allocation/check-in/completion, not Scale;
that does not exclude a candidate interaction or establish the native cause.
Controller `run-scale-disabled-regression/chain-4873fe873a1e4ea687996c4e3b3a453c`
retains sanitized failure, crash-reporter parent lineage and preservation receipts.
The recovery instance was closed after normal native close was requested. The
user confirmed answering/closing its Document Recovery prompt around15:18PDT.
Its selected retention option is unverified; closure is not automatic cleanup.
No COM reattachment or forced termination occurred. This run is not acceptance.

Separate actual-helper testing reproduced a tooling defect: when the supplied
process is absent, `Start-ExcelEnterDismissal` falls back to the title
`Microsoft Excel` and sends Enter to another process. The native sentinel test
`Test-LiveRoleDismissalOwner.ps1` first proves owned input works, then checks
that an absent owner cannot affect the other process. RED1/1 at
`live-role-dismissal-owner/369df1c807f34da69b481af598221ace` becomes GREEN2/2 at
`live-role-dismissal-owner/56cde10cd0fe49c8a75eaa2452cd648a`. Both use the actual
helper extracted from source, start no Excel and close the sentinel normally.
The helper now pins process creation time, stops when that owner disappears,
checks button/foreground ownership and removes the title fallback. This fixes
the reproduced input-scope defect; it neither explains the native crash nor
attributes the user-confirmed recovery closure to automation. XLAMs are unchanged.

One ordinary retry with that fix failed at a different boundary: projection
rebuild's `modProcessor.RunBatchReportForAutomation`, COM0x800706BE, with
`ntdll.dll/c0000028` at22:27:30.697Z. Controller
`run-scale-disabled-regression/chain-a781e3b11648461d9c60e2c2cd8ebd96` records
warehouse15/15, live32/1 and chain5/1, UTC22:25:02.419--22:28:05.627. Its recovery
process was identified as the crash reporter's child and closed after a normal
native close request; no recovery prompt was observed in the subsequent scan.
Closure was assisted, without COM reattachment or force. Packages, settings and
tracked reports were restored. This reproduces the earlier projection-call
signature and rules out treating the dismissal guard as a native repair. Do not
repeat ordinary chains until further localization provides a reason.

Regenerated maintenance evidence `run-scale-disabled-static-02` retains320
components,6342 procedures,138754 VBA lines,9/45 dynamic calls,192 duplicate
candidates and28 unchanged oversized caps. Three schemas validate and457
PowerShell scripts parse, including the new native ownership test.

Metadata-only native observation then reproduces the projection-call failure:
`run-scale-disabled-native/661781ea19534d0586008701647e063b`, controller
`run-scale-disabled-regression/chain-f908a084668740c494aef06c6b6cb840`.
At22:32:27.640Z an unmapped first-chance C0000005 precedes C0000028 by about100ms.
The latter first/second-chance stacks each contain32 bounded module/offset frames,
including VBE7 and OLEAUT32. The observer exited normally; business calls and VBA
were unchanged. Recovery closure was assisted by a normal native close request,
with no prompt observed afterward. No COM reattachment or force was used.
This identifies exception order, not a root cause. The existing stack filter
omits first-chance C0000005; its capture requires a calibrated test before the
next observation. The observed chain is failed diagnostic evidence, not acceptance.

The native calibration now raises and handles a software C0000005 without
invalid memory access or Office. RED46/3 at
`native-exception-calibration/ee13254f00bf4af78f9180b64098c04c` fails only the
missing access-violation stacks in Exit/Detach/FaultsOnly modes. After extending
the stack filter, GREEN49/49 at
`native-exception-calibration/dc96434e1514453f94c74e3049e3e189` retains all43 prior
checks in order. Schema, module/offset redaction and32-frame bounds apply to every
captured stack; exception handling, filter behavior and detach remain intact.
This is diagnostic tooling evidence, with no XLAM or workflow-contract change.

The next observed run does not reproduce the fatal sequence: native root
`run-scale-disabled-native/66b9fb8ec2394774bac272632d3fdf4e`, chain controller
`run-scale-disabled-regression/chain-3ae4127b27bc4b2c95a2731127d57f32`,
UTC22:34:24.563--22:42:50.757. Chain32, live48 and warehouse15 pass in exact prior
order, with preserved settings/packages/reports, unassisted closure and zero
delayed Excel failures. The native observer exits normally with no second-chance
fault. This is an observed diagnostic pass, not ordinary acceptance or a repair;
the first-chance access-violation stack was not exercised by Excel in this run.
Regenerated `run-scale-disabled-native-static-01` retains all preceding runtime
metrics and28 caps, validates three schemas inside the successful generator, and
parses457 scripts. Do not duplicate the generator's completed schema validation.

One further observation also finishes without the fatal sequence: native root
`run-scale-disabled-native/15b0e4827ae24216bd41c02baaa02542`, controller
`run-scale-disabled-regression/chain-cc4100d35a4f4ed480d5aa615dab13f6`,
UTC22:43:23.870--22:47:15.502. It retains31 chain checks,48 live and15 warehouse
checks in prior order, with clean unassisted closure and preservation. The final
static gate is explicitly unexecuted and reported as `StaticNotRun.DiagnosticOnly`;
exit1 is deliberate and does not represent a product failure. This partial
diagnostic is not full-chain acceptance. No first access-violation stack was
captured in these two passing observations; retain the calibrated observer for
the next reproduction rather than repeating chains until one passes.

## Earlier configuration-call interruption

Two ordinary Scale-candidate runs failed before the intended cases at
`TestD5Commands.SaveDirect`, returning COM0x80020009. The first failed during
read non-mutation checks (36 PASS/1 failure); the second during an unauthenticated
command (17 PASS/1 failure). In both, saved first-call evidence and native process
metadata establish that the new Excel window was a child of the original test
Excel instance. No matching Application1000/1001/1002 event was observed through
the respective delayed audit. This is not evidence of desktop lockout or a fixed
native cause. Desktop diagnostics have no error5 through the observed interval.

Controllers `production-run-local-controller/b4cffe4866e846b09234231d6eb27449` and
`production-run-local-controller/515a44f9db434068a8dc2d119713a440` retain failure,
closure and audit receipts. Workers are respectively
`slice4be-production-run-worksheet-scale/2b38cef5a3984dc88a27875e35719494` and
`slice4be-production-run-policy/8083fcde06f64ec8bf01c281bbaaf2f5`.
Both required normal assisted closure of the test child, with recovery files kept
for later viewing. No COM reattachment or forced termination was used; packages
and settings were preserved. Neither run is acceptance evidence.

The existing metadata-only native observer passed43 calibration checks at
`native-exception-calibration/ada39190cc6c46929a5b46cb40733117`, covering ordinary
handling, detach, stack bounds and redaction. It is attached to a fresh Policy
diagnostic in controller `production-run-local-controller/a55e2dba44b749499d1a8a6add98f610`.
Worker `slice4be-production-run-policy/698639e52db349aaacee7b9bee0bdb26` passes
218/218 in prior order, UTC20:56:03.978--21:02:09.403. The observer exits normally,
records one first-chance000006BA and no second-chance exceptions, and the delayed
Application audit is clean through21:02:21.403. Settings and packages are preserved.
The restart was not reproduced. Observation changes timing: this cannot establish
a repair or replace the ordinary acceptance gate. The exact failing native
boundary remains unresolved. Runtime source is unchanged by this investigation;
RUN-SCALE-01's approved runtime changes are recorded separately. A subsequent
ordinary Policy run passes218/218 with unassisted closure and zero delayed Excel
events: controller `production-run-local-controller/60b028cc94bd4c9c916980d56f23b847`,
UTC21:02:20.844--21:07:46.477; audit through21:07:58.477. This qualifies the scoped
policy gate, not a repair. The user also confirmed they did not open the earlier
unexpected Excel window, consistent with the saved test-process lineage.

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

## Fresh replay outcome and diagnostic reference cleanup

The diagnostic now emits only processor counts, allowlisted status/error classes,
authority deltas and identity booleans. Raw report tails, identifiers and error
messages are excluded. Pure tooling expands63 to **70/70** with redaction and
malformed-input checks; the initial missing-helper check was63/1, not product RED.
An exception-path guard first fails71/1, then passes72/72 after restricting cleanup
lookup to local variables: caller-owned references cannot be released accidentally.

The unchanged six replay checks pass on the same copied source fixture at
18:08:29--18:12:14Z. One event applies with exact identity, one log/applied increment,
processed status and no duplicate on replay. Excel remains210 seconds after Quit,
then exits with its worker; delayed audit18:12:43Z has zero events. The controller
still fails its explicit exit-before-worker gate. All13 source files, settings
and five packages are preserved. No assisted closure is needed.

Read-only inspection of the original generated fixture and earlier failed copy
finds the same HELD inventory lock, active with44.105 seconds remaining at the
earlier failed replay and expired before the passing replay. This supports the
three-minute lock lease as the explanation for the three failures; the earlier raw processor
report was not captured, so it is not retrospective proof of its exact return.
Inspection uses fresh Excel with events/macros disabled, no refresh or save,
unchanged hashes and a clean delayed closure audit. No lock was reset or bypassed.

An experimental diagnostic releases only its locally acquired table/worksheet
references before closing the inbox. It retains6/6 and Excel exits while the worker
remains alive, with zero delayed events and full preservation. The same cleanup is
then applied to `Test-Slice4beShutdownControl.ps1`; the actual tool repeats6/6 and
the same clean lifecycle. This corrects the diagnostic's reference lifetime. It
does not repair the separate live processor crash or identify the original
untraced full-chain combase owner. Runtime and frozen packages remain unchanged.

Exact ignored receipts:

- Pure tooling: `shutdown-trace-contract/c4b7c574a09e4c3e8645c400dcfdeb02/checks.json`.
- Before cleanup: `native-projection-replay03-control/qualified-verification.json`; `slice4be-shutdown-control/27eca37b39c74994be7b4cb3c7140ed8/projection-outcome.json`.
- Read-only lock state: `native-projection-lock-read-01/lock-state.json`, `preservation.json`, `verification.json`.
- Experimental cleanup: `native-projection-release-control-01/instrumentation.json`, `qualified-verification.json`.
- Corrected diagnostic: `native-projection-replay04-control/qualified-verification.json`; `slice4be-shutdown-control/d8f932dce8564b648303e0f1f397e2bf/`.

The first full-chain deleted-table-release experiment stops before projection at
a trace append sharing violation (chain5/1; no fresh live gate credited). Excel
closes normally with preservation and zero native events. It does not test the
release hypothesis. A real file-contention test reproduces72/2; bounded retries
for Win32 sharing/lock violations32/33 produce **74/74**, preserving all72 checks.
Other write failures still throw. Writers serialize complete metadata-only lines;
there is no runtime change or product RED.
Final static evidence at `general-native-outcome-static-03/ratchet-verification.json`
retains all metrics/28 caps; three schemas and455 PowerShell scripts validate.
`native-outcome-tooling-verification.json` verifies all63 prior pure checks remain
the exact ordered prefix of74 GREEN checks.

Receipts: `general-native-release01-regression/chain-dac91b5d5fbf4b0fbbb3184952fee31a/native-verification.json`;
`shutdown-trace-contract/64eef894a1024311baf1dc66deeca5f9/checks.json` (caller-scope RED),
`shutdown-trace-contract/9d0573aff77e4e6fa1ef13a6ed1eb46f/checks.json` (scope GREEN),
`shutdown-trace-contract/470463514dc4417da44f48977f0b923d/checks.json` (contention RED),
`shutdown-trace-contract/60beb25ad6224670b65aa3661400764b/checks.json` (74 GREEN).

## Deleted-table cleanup and ordinary-chain limit

The corrected experimental chain runs18:19:30--18:25:12Z and passes **32/32,
live48/48, warehouse15/15**, retaining every original ordered check. The processor
returns after deletion and all three rebuild/authority assertions pass. All six
owned Excel instances close without assistance; delayed audit has zero Excel
failure events. Two instances had remained after Quit until worker teardown, so
this is not proof of prompt per-procedure shutdown. Settings, reports and packages
are preserved. Receipt: `general-native-release02-regression/chain-0d52431b24aa4140809a0cff9a43b203/native-verification.json`
and `qualified-counts.json`; trace `general-chain-native-release-02/`.

The regular live validator now releases and clears its two deleted table wrappers
immediately after Delete. All143 business calls retain exact order and text
(`live-validator-cleanup-source-verification.json`); runtime/packages are unchanged.
Pure checks remain74/74. The ordinary untraced chain then fails **before this
cleanup is reached**, at Receiving ConfirmWrites: chain5 PASS/1 FAIL,
live14 PASS/1 FAIL, warehouse15/15;27 chain and34 live checks are unreached.
The reported macro is `modTS_Received.RunReceivingConfirmWritesFormActionForTest`,
HRESULT `0x80020009`. No corresponding Application1000/1001/1002 Excel event is
found through12 seconds after closure. Its child Excel process requires native
CloseMainWindow; no COM reattachment or recovery-file deletion occurs. The child's
precise role is unproved. This is an unresolved automation failure, not a proven
native crash or evidence about the unreached projection change.

Ordinary-chain receipt: `general-native-release-final-regression/chain-5d34bc36c71c43dab886cd1bc2d35212/qualified-verification.json`,
`qualified-counts.json` and `assisted-close.json`. Final corrected diagnostic
replay retains6/6, closes while its worker lives and preserves all13 source files,
settings and packages with zero delayed events: `native-projection-replay05-control/qualified-verification.json`.
Final source pure checks: `shutdown-trace-contract/621b3b12a7da4d23a5593064bcc4b8aa/checks.json`.
Final static: `general-native-outcome-static-04/ratchet-verification.json`, with
unchanged metrics/28 caps, three valid schemas and455 parsed PowerShell scripts.

The existing untraced `BeforeProjection`/`Setup Chain` control then passes31/31,
the exact ordered prefix of48, with explicit General02 pins, no VBA instrumentation,
normal closure, full preservation and zero delayed events. Receiving succeeds;
17 later live checks remain unexecuted. Receipt:
`projection-live-control/6fc9ba39362e4332a3c4750018f04d09/qualified-verification.json`.
The14/1 failure is not reproduced or explained. Continue independent Settings
regressions and observation coverage; retain ordinary-chain validation as an open
final gate. No universal native repair or 4be-A completion is claimed.
This same14/1 Receiving macro/HRESULT pattern is already documented in
`plan022_slice4be_settings_editor_activity_results.md`; do not treat it as a new
failure class or infer its cause from a subsequent successful run.
