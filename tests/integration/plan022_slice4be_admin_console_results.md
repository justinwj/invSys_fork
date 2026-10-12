# Plan022 Admin Console recording

4be-A, Architecture v4.11 D18 discovered-control refinement; catalog36.
Eligible NAS recording/guide coverage is verified on unpromoted console-03,
with the explicit guide-evidence reuse below. This is not4be-A acceptance.
D18-LOCAL-01 remains proposed and unapproved.

The generated `btnAdminOpen` Ribbon callback retains ADMIN_MAINT and dispatches
`modAdmin.Admin_Click`. REQUESTED precedes the original owner; OPENED means the
console actually opened, otherwise FAILED. Effects are Unknown because opening
prepares local projections. Empty references and fixed messages must not claim
processing or Domain application. The wrapper captures the original workbook and
discovered configuration path before optional recording; an empty captured path
stays empty. Existing service callers retain discovery. Session/storage failures
must preserve the owning action and truthful evidence.

## Focused RED, 2026-10-12

`Test-Slice4beAdminConsole.ps1 -DeployRoot deploy/validation-admin-verify-03 -Phase RED`:
**55 PASS / 9 expected FAIL**,64 checks. Controller
`reports/runtime/admin-console-controller/e1a9a99186574778b75ec88ddadff9ec`; worker
`reports/runtime/slice4be-admin-console/ba18ba81c4e740d0a7899422f905b46c`.
UTC12:39:03.4372116 to12:40:46.3085539; clean delayed native audit through
12:40:58.3085539. Five instrumented compiles, exact console/JSON, compiled pins,
settings/package/tooling preservation and normal Excel exit pass.

An Office.IRibbonControl fixture invokes the actual generated Admin dispatcher.
Owner-entry and successful-presentation probes observe the existing owner without
replacing its result. A protected console sheet produces a real opening refusal.
The nine REDs are missing registration and four recording assertions for each
successful and protected opening. Business execution, actual reader/signed-out
Ribbon denial, policy suppression, custom workbook values and saved warehouse
files all pass. Denied Ribbon entry creates no fabricated callback action.
Only message presentation is intercepted; this does not establish native input
or guide/visible acceptance. Raw artifacts remain ignored.

## Candidate01 qualification

`deploy/validation-admin-console-01`, pinned in
`reports/runtime/admin-console-build-01/compiled-pins.json`: five cold compiles
and Operations startup pass. Compared with verify-03, four existing components
change (`modAdmin`, `modAdminConsole`, `modActivityCatalog`, `modEvaluationMatches`),
two are added (`modAdminConsoleActivity`, `modAdminConsoleCodes`), and323 prior
components are identical. Settings and the frozen prior candidate are preserved;
Excel exits normally; no native failures through12:43:38.2069736 UTC. Unpromoted.

The console's existing opening body moves behind a typed public delegate, keeping
the oversized console module below its previous size. Its observer follows that
owner body in the disposable package; the operator still uses the original
Ribbon/callback. Context/storage, guide, visible and current-package regression
gates remain pending.

Candidate01's first GREEN attempt is **63 PASS / 1 FAIL**, not acceptance:
`admin-console-controller/8ba417f32edf41c4838a2bbc9fe0347d`, worker
`slice4be-admin-console/4c0e6a675ddd45af82aa5e45ee91bee6` under `reports/runtime/`.
All64 prior identities, five compiles, console/JSON, pins and preservation pass;
normal closure and no native failures through12:46:11.1883667 UTC. The sole
failure is `AdminConsole.Opened.ActualOwnerPresentation`: records correctly say
OPENED, but the callback returns with a different surface selected.

An added read-only focus observer reruns the same64 checks on candidate01:
controller `admin-console-controller/9f2583ea02f04382a6c484ff73352e9a`, also63/1.
It confirms owner=True, presented-at-owner=True, selected-at-return=False for
the successful recorded case; suppressed cases retain presentation. Optional
outcome recording changes focus after the owner. Five compiles, preservation,
normal closure and clean delayed audit through12:48:55.2419230 UTC pass. This is
a behavioral RED under existing captured-workbook/presentation requirements.
The next focused gate adds exact returned-surface checks for all seven cases.

Expanded candidate01 diagnostic: **67/4**,71 checks, controller
`admin-console-controller/a9b037ec1afb4d42ad26697f50df9f69`; clean audit through
12:50:58.6675423 UTC. In addition to the original presentation failure, returned
surface checks fail for Opened, Protected and Reader. The unchanged verify-03
comparison is61/10, controller
`admin-console-controller/9d5c20dd1eea4c8f95f3fd4f7c41cba6`; clean audit through
12:53:48.3027896 UTC. Its Opened/Protected return surfaces pass; Reader also fails
there, alongside the nine expected absent-record REDs. All runs preserve exact
prior identities, five compiles, settings/packages/tooling and normal closure.

The Reader focus change occurs in the permission gate before the console owner
is invoked, including on the unchanged package. It is not a console regression
or proof of a reachable authorized action. Retain it as opt-in
`-PreDispatchSurfaceDiagnostic` (RED only), without claiming its acceptance.
The default console gate retains all64 original checks plus the five exact
returned-surface checks whose owner actually runs:69 total, the first69 checks
of the71 diagnostic. That prefix supplies66/3 behavioral RED on candidate01.
Existing Reader/SignedOut authority, zero owner invocation, no fabricated action
and workbook/warehouse preservation checks remain in the default gate.

Candidate02 restores the exact still-live sheet around optional observation in
Admin, preserving owner presentation on success and failure. It never reopens a
workbook or asks headless Core to select UI. Five cold compiles, Operations startup,
the same four changed/two added components and323 identical prior components pass.
Frozen verify-03 and settings are preserved; normal closure and clean native audit
through12:55:12.5086522 UTC. Pins: `admin-console-build-02/compiled-pins.json` under
`reports/runtime/`. Qualified **69/69 GREEN**, controller
`admin-console-controller/542b90d4f6844eeba143d338dc827b9e`, worker
`slice4be-admin-console/56fe63fbd6794cbb80c39a303541015b`, under `reports/runtime/`.
Exact64 prior prefix, five compiles, console/JSON, compiled pins, preservation and
normal closure pass; no native failures through12:58:13.0045205 UTC. All five
owner-executed cases retain the required surface. Broader acceptance is pending.

The first context/storage extension stops93/2, controller
`admin-console-controller/cbfe7019c9cd48e7a7edd28b097220cf`. The69 prefix and
workbook/sign-out/target-switch cases pass, but setting a shared Core runtime
override to an empty directory prevents the Ribbon gate from reaching the intended
empty-scope fault. This is a fixture failure, not product RED or acceptance.
Normal closure restores settings and preserves packages/tooling. The corrected
fixture supplies an actual empty directory only to the console's scan-root reader;
it does not change the live warehouse/session, fabricate a read result or force
owner success. Product source and candidate02 remain unchanged for the rerun.

Corrected context/storage gate: **118/118 GREEN**, exact69 prefix, controller
`admin-console-controller/307eb2aa74074f2da8fc02004742eb6c`, worker
`slice4be-admin-console/0faba160cd2846b0ba4704f1d2d6405b`, under `reports/runtime/`.
Five compiles, console/JSON, compiled pins, preservation and normal closure pass;
the delayed native audit is clean through13:08:33.8431096 UTC. Workbook switching,
sign-out, target changes, captured empty scope, and real attempt/outcome write
failures preserve one owner execution, its selected surface and original scope.
Evidence stays attributed to the original context; tracking failures produce one
truthful warning without inventing a completion or altering saved authority.

Static01 finds336 components/6414 procedures/139759 lines, unchanged dynamic9/45
and28 module caps, but198 duplicate candidates versus the accepted194. This is
not a passing maintenance gate. Two observation bodies duplicate Verify Add-ins;
two literal-erased catalog factories repeat its record construction. Consolidate
the observations behind one typed Admin helper and the two fixed definitions in
one Core catalog module, retaining exact historical definitions and messages.
The live candidate02 guide gate remains frozen; the refactored source requires a
new compiled candidate and both affected callback regressions before acceptance.

Static02 after consolidation passes:336 components/6410 procedures/139721 lines,
194 duplicate candidates, all28 prior caps without growth, dynamic9/45, three
schemas and490 PowerShell parsers. No new exception. `modAdminActionObservation`
owns the shared optional begin/finish behavior; `modAdminActionCodes` retains
the two exact fixed definitions behind version-gated catalog lookup. This is
source/maintenance evidence, not compiled runtime acceptance.

## Independent recording and guide proof on candidate02

`-CheckGuidePaths` passes **188/188**, exact118 prefix, controller
`admin-console-controller/68d1641b103649838e30c81b9b8f5673`, worker
`slice4be-admin-console/2d23318bf1484d29894fd5a13beba6e8`, under `reports/runtime/`.
UTC13:11:05.4386253 to13:43:59.5743546; clean delayed native audit through
13:44:11.5743546. Five instrumented compiles, exact log/JSON, compiled pins,
settings/packages/tooling preservation and normal Excel closure pass.

Three independent real callback journeys supply Source, Observed and Interrupted
recordings. The interrupted second outcome uses a real write failure; its owner
still opens once. Actual Start/Stop, immutable journal chains, exact publication,
guide authoring and reader/evaluation handlers preserve the source/observed
distinction. Successful evidence concludes two openings; missing completion
evidence stays Incomplete. Neither borrows source success or claims Domain effects.
All three views preserve evidence. Inventory and the custom-column canary remain.

Six `admin-console-guide-*.png` captures were reviewed: successful How-To,
Diagnostic, Compare and scrolled conclusion; interrupted Compare and conclusion.
Multiline guidance, full provenance, matched/unmatched steps and saved conclusions
are readable. Captures and raw reports remain ignored. These are packaged-handler
and visible-form results; they do not claim physical mouse-input acceptance.

The maintenance refactor needs new-candidate callback proof. Reuse this guide
evidence only with compiled identity of its guide/evaluation/view components and
fresh exact recording/boundary checks; report reused evidence separately from reruns.

## Consolidated candidate03

Five cold compiles and Operations startup pass; `admin-console-build-03` records
compiled pins, restored settings/frozen verify-03 and a clean delayed native audit
through13:45:30.5401025 UTC. Relative to verify-03, five existing components change,
three are added, the old `modAdminVerifyCodes` is replaced, and321 remain identical.
Relative to console-02, only the catalog and two Admin observation callers change;
the two specific code factories are replaced by the shared factory/helper. All324
other components, including guide/evaluation/view implementations, remain identical.

Fresh **118/118** on03 retains the exact02 check order: controller
`admin-console-controller/65f4fad0483942f889bb33f1984b1352`, worker
`slice4be-admin-console/6d2c713875594d2a84790129c7711e5b`, under `reports/runtime/`.
Five instrumented compiles, exact log/JSON, compiled pins, preservation, normal
closure and delayed zero native audit through13:50:07.6440112 UTC pass.
Final static03 retains static02's metrics and passes all three schemas/490 parsers.
Prior guide188/six captures are reused with this component comparison and fresh
recording/boundary proof; they are not represented as a second188 run on03.

Fresh Verify Add-ins **130/130**, exact prior order: controller
`admin-verify-controller/786fb7a7e5c642f69aeb19a5f58df654`, worker
`slice4be-admin-verify/3886d80c3903496984ec76fd6b518e1e`. Five instrumented compiles,
current03 pins, exact log/JSON, preservation and normal closure pass; delayed
native audit is clean through13:53:01.3786654 UTC.

Fresh Reconcile **258/258**, exact258 prefix of accepted360: controller
`admin-reconcile-controller/b686a1170e584c9ea77e8932522b7d8c`, worker
`slice4be-admin-reconcile/98c61d32068f4e02a106136bfd4e40fb`. Five instrumented
compiles, current03 pins, exact log/JSON, preservation and normal closure pass;
delayed native audit is clean through13:56:21.7026366 UTC. Current minimum/restored
history and typed-confirmation captures were reviewed; identities and preview fit.

Verify202 and Reconcile360 guide evidence remain reused baselines, not fresh full
runs. Their guide/view/Domain implementations are unchanged. Source comparison
also proves `modEvaluationMatches` differs from the prior committed classifier by
only the explicit Console/OPENED case; fresh130/258 protect previous recording
and conclusion behavior. `admin-console-build-03/prior-guide-reuse-check.json`
records that comparison and the exact prior258 prefix. All paths above are under
ignored `reports/runtime/`; no raw runtime data is committed.

Current03 regressions are qualified with exact prior order, compiled pins,
preservation and clean delayed native audits:

- Settings202: controller `admin-console-regression/settings-5e48f767977340e08dc57c95b86da49b`,
  worker `slice4be-tracking-settings/3de2100da70140178be99b80893aefee`;
  audit through14:01:17.2997896 UTC. Five instrumented compiles and preference
  restart pass;10 references release without failure and Excel exits unassisted.
- Smoke86: `admin-console-regression/smoke-53b810fc2ca24c049846806d61d3be04`;
  both exits are unassisted, delayed audit through14:01:54.6790869 UTC.
- Full chain32/live-role48/Create15:
  `admin-console-regression/chain-17b34cea28654d9ba14bc7e9e7c98a9d`;
  normal closed/restored state and zero native events through14:07:45.8673635 UTC.
  The two buffered cleanup messages were completed inter-stage waits; no agent
  close, restart, recovery attachment or process termination was used.

All paths are under `reports/runtime/`. Static03 and the exact compiled-component
comparisons above complete this control's scoped gates. D18-LOCAL-01 and remaining
Admin reachability keep4be-A open. No station deployment is promoted.
