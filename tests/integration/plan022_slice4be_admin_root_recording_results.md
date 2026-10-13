# Add Warehouse Root recording

4be-A / D18 discovered-control refinement,2026-10-13 Pacific. **Scoped-qualified.**
Unpromoted candidate `deploy/validation-root-recording-04`;4be-A remains incomplete.
Owner feedback prerequisite D18-ROOT-01 is approved/qualified in code61b6df8d and
docs7cb25a0; see `plan022_slice4be_admin_root_write_results.md`. Preserve that gate.

Catalog40/162 adds `ADMIN_WAREHOUSE_ROOT_ADD` under existing D18 semantic inheritance.
The specification, Plan022 and Controls1.570 define cancellation, confirmed or
unconfirmed persistence with successful or failed refresh, and unexpected failure.
Only REMEMBERED_REFRESHED concludes CommandCompleted. No path payload, source
reference, permission change or replay adapter is added.

## Qualified recording RED

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File tests/tooling/Test-Slice4beAdminDirectoryColumns.ps1 -DeployRoot deploy/validation-root-feedback-01 -Phase RED -RootRecording
```

Controller `admin-directory-columns-controller/2d10bdd6f58d441c88a2ebbe2f9928ea`:
**168 PASS/29 expected FAIL**, missing catalog definition and four observation/
context/payload/conclusion checks in each of seven eligible owner cases. All113
nonvisual owner checks remain intact. Five compiles, exact log, compiled/tooling
pins, restoration and normal closure pass; native audit clean through15:57:24.526UTC.
The local ignored `verify-root-recording-red.ps1` records the exact failure set.

## First candidate and excluded run

Unpromoted `validation-root-recording-01`; build `root-recording-build-01` passes
five cold compiles and Operations cold start. Two added Admin modules, five changed
components and331 unchanged compiled components match the intended scope. Frozen
owner candidate/settings preserved; native audit clean through16:00:16.925UTC.

First candidate run, controller `e925ee584f9c4d1386ebec42bbdcb5c2`, worker
`slice4be-admin-directory-columns/ba8363b64e724264a217b92985c61a6e`: **194/10**,
excluded. All recording correlation/effect/privacy/conclusion and suppression/
denial cases pass, as do native warning checks. Nine strict owner file-set checks
fail after command-default activity writes; the owner-only fixture now explicitly
disables recording before its snapshot, retaining the original strict assertion.
The worksheet capture encountered a Shell popup covering the content. No error5
or native Excel failure: pins/restoration/normal closure pass and audit is clean
through16:07:46.578UTC. Do not count this as a completed GREEN gate.

## Optional setup failure and focused GREEN

Frozen recording-01 setup-fault RED: controller
`30278f7201d04d6e92405d1f865e06d0`, worker `eabc27e05fa84e9299cdc1104a5bd9b6`,
**200/3**. The reached optional discovery-capture fault prevented the original
prompt/write/refresh. Only owner counts, presentation and feedback checks fail;
five compiles, pins, restoration and normal closure pass, native audit clean
through16:15:10.437UTC.

Candidate `validation-root-recording-02` catches that optional setup failure,
preserves ordinary owner work and reports one tracking notice. Build02 passes five
cold compiles and Operations cold start; only Admin Console/root observer changed
from01, with336 other compiled components identical. Audit clean through
16:18:38.300UTC. Focused GREEN **210/0**, controller
`8dc761e59f3245c78e9eeb9717d754c6`, worker `2d101e6f0be64a118e02e829b213b57c`:
all203 setup-RED checks retained, three reviewed screenshots, five compiles,
exact log, pins, restoration and normal closure pass. Native audit clean through
16:27:08.071UTC. This is a focused gate, not recording coverage completion.

## Expanded fault and guide qualification

Active-recording setup fault on frozen02: controller
`974cc981798b4b039058dc34e3cea23e`, worker `119094cb90bc4251a6feab56839ee08d`,
**RED207/2**. The actual callback runs its owner correctly, but capture stays active
without a durable incomplete closure. Only NoLongerActive/DurableIncomplete fail;
five compiles, exact log, pins, restoration and normal closure pass, with native
audit clean through16:37:03.002UTC. Candidate03 adds a primitive Core interruption
for the captured context only. Focused **GREEN217/0**, controller
`74a7390d25384acb8c33a6a7f6b920ba`, worker `b7eaf1ff0c634b079e00a76a9e0d4909`,
retains all209 RED checks and proves unrelated/empty contexts cannot close capture.
Five compiles, three reviewed images, pins, restoration and normal closure pass;
native audit clean through16:45:40.756UTC. Build03 has336 unchanged components,
five cold compiles and Operations cold start, native audit through16:38:10.828UTC.
No fabricated action and no owner rollback are introduced.

Intermediate static03:345 components/6438 procedures/140271 lines, dynamic9/45,
28 caps preserved. Duplicate-body candidates rise196 ->197 because ContextClosed
and the new ObservationUnavailable wrapper differ only in their fixed reason.
Do not accept that regression: consolidate into the existing primitive closure
with an explicit optional tracking-unavailable flag, retaining VIEWER_CLOSED by
default; rerun focused/regression/static gates on the final candidate.

Expanded frozen03 guide-form **RED279/1**, controller
`6c918edda01b4519a103205084cd7174`, worker `e8889ed291804bdbb24e63905cc7c50f`:
only AllRootOutcomesOffered.0 fails. Five context/storage boundaries, independent
source/observed/interrupted recording lifecycles, publication and immutable prior
history pass. Five compiles, exact log, pins, restoration and normal closure pass;
native audit clean through16:58:14.054UTC.

Candidate04 adds the four editor choices and consolidates capture closure. Build04
passes five cold compiles/Operations cold start,335 other compiled components
unchanged, native audit clean through17:00:49.228UTC. Initial expanded run
`d525e5c5fcc64ddeabf6b52809f8b4b0` / `761fb2803dca4fab806fbb8415862d5c` is **excluded**:
four instrumented projects compiled, then the Admin test probe failed compilation.
ProcStartLine included leading separators; inserting at+1 put executable statements
outside the capture function. Use the unique declaration line instead. The exact
owned compile dialog required acknowledgment (WM_COMMAND for its verified unique
OK; background BM_CLICK did not dismiss it). No recovery attachment or forced
termination; settings/packages/tooling preserved, native audit clean through
17:05:35.742UTC. This is neither product RED nor GREEN.

Rerun04 **focused GREEN370/0**, controller `9e918c298d134c91b471079f0afe6680`, worker
`af6b2195a2bd4444bef67c60bff2600f`, retains all280 guide-RED checks. It covers unexpected
prompt failure, actual post-write remembered roots with frozen ambient/open-workbook
scope, and separate Concluded/Incomplete/Failed guide results, including a fully
recorded unconfirmed preference write. All11 captures reviewed: custom columns,
two native warnings, four successful guide views, two interrupted and two failed
views. Five instrumented compiles, exact logs, pins, settings/packages preservation
and unassisted closure pass; native audit clean through17:23:29.290UTC.

Static04:345 components/6437 procedures/140271 lines; two added modules/five net
procedures/94 lines versus the owner checkpoint. All28 module caps pass, dynamic
calls remain9/45 and duplicate candidates return to196. Three schemas and all520
PowerShell scripts pass. No static exception is requested.

The applicable current-candidate regressions below qualify this recording increment.
Native InputBox/Ribbon input and root replay are not claimed.

## Current-candidate regressions

Packaged smoke **86/86**, exact prior order: controller
`root-recording-regression-04/smoke-d91d9eb995644c3097d2f487d3fca5d3`.
Both Excel exits are unassisted; compiled pins, settings/packages/tracked reports
and delayed native audit pass. Run17:24:03.233--17:24:35.842UTC.

Full chain **32/32**, live roles **48/48**, Create Warehouse **15/15**, exact prior
orders: `root-recording-regression-04/chain-2c9bcfe4385d4024be1dcfb38a8c0dc3`.
Run17:25:26.734--17:31:34.760UTC; normal closure, preservation and clean native
audit through17:31:46.760UTC. Buffered cleanup waits completed without intervention.

Compiled comparison with the qualified owner checkpoint finds two added Admin
modules, seven changed components and329 unchanged components. The changed Core
components are ActionRecording, ActivityCatalog, AdminActionCodes,
EvaluationMatches and ExpectationDraft; Admin changes are modAdmin/modAdminConsole.
All Operations and Domain component bodies are unchanged. Existing local-recording,
Users, Seed, Reconcile and reusable Production receipts remain scoped prior evidence,
not newly executed full gates. Fresh Directory/guide, Settings and default Viewer
closure checks protect the shared dependencies changed here.

Directory/guide **219/219**, exact prior order: controller
`admin-directory-columns-controller/b47d5c6af75f4528b7ad2b6ee2e1cdf6`, worker
`slice4be-admin-directory-columns/e814dee5bdc34122947bd22731572cf7`.
Five instrumented compiles, exact log/JSON, compiled/tooling pins, preservation
and normal closure pass; native audit clean through17:41:13.468UTC. All seven
captures reviewed, including successful/interrupted independent guide results.
The existing catalog39 definition and catalog1-38 preservation checks remain;
records now assert the candidate's actual catalog version40.

Initial Settings regression is excluded: controller
`root-recording-regression-04/settings-d904e02b5aac4091b71da9713b177259`, worker
`slice4be-tracking-settings/a196a08268b04e1884b6056a0d8f0807`,27 PASS/two FAIL
(EditorLoaded and its harness exception). The test still expected catalog39/161;
update its exact expectation to40/162 and rerun, retaining all202 prior checks.
Five compiles, preservation and normal closure pass; native audit clean through
17:42:34.740UTC. No product change or product RED is claimed for this fixture fix.

Settings rerun **202/202**, exact prior order: controller
`root-recording-regression-04/settings-d1ec2a6237264028af51b34ecdbfcd7e`, worker
`slice4be-tracking-settings/58236421d16847be9b189fe7cc7f53d8`. Five instrumented
compiles, exact log/JSON, compiled pins, preservation and normal closure pass;
preference restart is unassisted with all10 owned references released. Native
audit clean through17:48:12.048UTC. No runtime change was needed.

Recording regression preflight `recording-3858e67e70034dc195ea83b160a96d05` is
excluded before Excel/checks: the compile guard runs before the implied published
reader flag is expanded. Pass both existing flags explicitly; do not change the
guard. Settings/packages were preserved and Excel remained closed.

Recording attempt `recording-f37fd9fc9e574176b71afb3fe1212d97`, worker
`slice4be-viewer-published-read/4f99fe4a0fef421a9c88d459d8a48fde`, is also excluded:
55 PASS/one fixture exception. Its strict Save Value-only fixture encountered
the now-supported Config selection observations (two REQUESTED/SELECTED pairs).
Use the existing SaveValueOnly policy option at the two lifecycle setup calls;
retain the exact two-record/journal assertions and every existing action handler.
No product change or observation filtering. Five compiles, preservation and normal
closure pass; native audit clean through17:51:31.392UTC.

Recording lifecycle **80/80**, retaining all64 baseline checks in order plus16
existing compile/layout/read-preservation checks: controller
`root-recording-regression-04/recording-17c06586f0de49f79758ed127430d5e9`, worker
`slice4be-viewer-published-read/dbd850c597904f328190ab8f1e2cb5e7`. Actual Viewer
closure now additionally requires VIEWER_CLOSED; optional setup failure remains
separately protected as TRACKING_UNAVAILABLE in focused370. Five compiles, exact
log/JSON, compiled pins, preservation and unassisted closure pass; native audit
clean through17:56:39.389UTC. No product changes followed the370-check candidate.

The scoped Add Warehouse Root increment is complete. Comprehensive Admin recording,
remaining native/human acceptance and final early B0 accounting remain open.
Next audit Create New Warehouse through its actual callback/form handlers before
defining observations; distinguish local bootstrap from initial publication.
