# Slice4be-A: Shipment Reconcile identity prerequisite

2026-10-12. **Identity/authority172 GREEN; combined native179/2.**
Approved D8-RECON-01 is implemented; long-history presentation and4be-A remain incomplete.
See the final sections for current evidence.

## Historical initial RED

The implementation/evidence section below supersedes the historical RED status.
Initial candidate `deploy/validation-admin-entry-02`; expanded RED uses
`deploy/validation-admin-auth-read-01`, whose Reconcile/Core/Domain code is unchanged.
Both retain catalog33/151. No Reconcile implementation has changed.
Governing contract: v4.11 D14 exact identity/unknown columns, headless owners and
read preservation. The user approved D14-RECON-01 on2026-10-12 after this RED.

Run:
```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File tests/tooling/Test-Slice4beAdminReconcileIdentity.ps1 -DeployRoot deploy/validation-admin-entry-02 -Phase RED
```

Qualified result: **15 passed,3 failed**, five packaged compiles. Admin-generated
training warehouse and Admin demo CSV upload create two opaque keys sharing SKU.
The actual Domain bridge applies one shipment and one later adjustment to one key.
The unchanged public `OpenShipmentReconcileTool` callback receives deterministic
selection/narrative answers through test-only presentation seams. No native-input
claim; no lookup, queue, payload or business result is stubbed.

Failures:
- `Reconcile.Payload.ExactSelectedSystemKey`: selected key is absent.
- `Reconcile.Payload.NoLegacyRow`: queued correction contains legacy Row.
- `Reconcile.Preservation.InventoryNotDirtied`: clean borrowed inventory becomes dirty.

The callback confirms one queued correction with the expected single-event delta
and evidence. Both quantities, inventory file bytes, a saved unknown-column canary
and the other warehouse Config remain unchanged. Domain processing was not run.
Core's resolver enters `OpenOrCreateCanonicalInventoryWorkbookLocal` and schema
maintenance even for a borrowed workbook; this is a suspected dirty-state cause,
not yet an isolated causal proof. Do not infer custom-column loss: the final
VBA-staged canary passed before and after the callback.

Ignored receipts:
- Controller `reports/runtime/admin-reconcile-controller/98f7f7b2ba024611b27f0004edf6fd4a`.
- Worker `reports/runtime/slice4be-admin-reconcile/3a6a8caff1b24490abeef3a809f092cb`.
- `qualified-verification.json`: exact ordered log/JSON agreement, five compiles,
  candidate/settings preservation, normal Excel closure, no native failure events
  through03:57:02.8767575 UTC.
- `reports/runtime/admin-reconcile-red-static-01/ratchet-verification.json`:
  all28 caps,326 components/6371 procedures/139227 lines, dynamic9/45,
  duplicate candidates192 unchanged; three schemas and471 PowerShell files pass.

Earlier runs are harness diagnostics, not RED: unsupported gate; open-workbook
hash sharing; PowerShell5 nested JSON array inspection; and COM canary setup.
The hash-failure run needed an owned WM_CLOSE after Quit; that run does not qualify
for normal shutdown. Final run needed no assistance. Use shared-read streams,
direct ConvertFrom-Json assignment and actual VBA canary staging; do not repeat
the failed approaches or interpret their preservation failures as product loss.

## Approved quantity/history and Domain RED

`Slice4beAdminReconcileHistory.ps1` extends the same gate under approved
D14-RECON-01. Qualified **31 PASS / 24 expected FAIL**, retaining all original
18 identities, their order and outcomes. Five instrumented compiles, exact
log/JSON agreement, package/tooling/settings preservation, normal unassisted
shutdown and a clean delayed Application audit pass.

- Controller: `reports/runtime/admin-reconcile-controller/19fc5155c0db4587837b7b3d34d7b6dd`.
- Worker: `reports/runtime/slice4be-admin-reconcile/b7f66bcb810f4fd686f12d65472f5dc6/red.json`.
- UTC: 04:32:01.7739619–04:32:47.7929839; audit through 04:32:59.7929839.
- Run the command above with `-DeployRoot deploy/validation-admin-auth-read-01`.

Owner-generated fixtures independently assert quantities before the callback.
Selecting the older of two shipments must compare entity A's current21 against
its after-event15, yielding -6; later changes to same-SKU entity B are excluded.
A later event contains two lines for A and one for B. Its selected A balance is18,
current A22 and B55, so the correction is -4. This protects all lines at the
selected AppliedSeq, exact event selection and independence from another entity.

The 24 failures comprise the original three, five in each history case (exact
event/key payload, entity-only delta, key visible in both prompts, current
attributes, and nonmutating inspection), two queued-application failures, and
nine Domain guard failures. Applying the actual queued payload reports APPLIED
but does not produce the exact entity quantity and clears custom-column values.
Current Location/Condition and duplicate-event idempotency pass. Domain also
accepts a wrong entity's shipment, a non-SHIP event reference and a positive
correction, changing quantities/history; missing narrative is correctly refused.
These are real callback/queue/Domain calls, not mocked business outcomes. Direct
Domain application supplements the callback; full inbox-processor acceptance
and native-dialog input are still separate gates.

Canary attribution is explicit: fixture event application itself clears the
manually staged projection values. Each inspection case therefore restages and
proves distinct key-bound saved canaries and sorts rows by quantity immediately
before the callback. Inspection preserves
those values; actual queued application separately exposes their loss.
`RewriteEntityProjectionTable` clears and rebuilds managed fields
without restoring custom values; the required preservation fix remains open.

Three preliminary attempts are not the final RED: untyped adjustment fixtures
retained factory `IoType=CREATE` (15/5, including a fixture exception); invalid VBA
object-return indexing (4/1 compile failure); and missing canary preconditions
after fixture events (24/24). The compile dialog required verified native OK
dispatch before the original cleanup finished. No force-kill or COM reattachment;
all controllers restored settings and preserved packages/tooling. Use explicit
SHIP/ADJUST fixture types and a local object variable before assigning a function
return. Final RED closes without assistance. The earlier qualified55-check run
(`admin-reconcile-controller/0159fc605ee945fcb2040682b7689e17` under runtime)
has identical check order/outcomes; final canaries additionally protect mapping
by key across sorting instead of using one repeated value.

Maintenance: `reports/runtime/admin-reconcile-history-red-static-01` holds all28
caps and unchanged326 components/6371 procedures/139226 lines, dynamic9/45 and192
duplicate candidates. Three schemas validate and473 PowerShell scripts parse.

The implementation below resolves these focused failures under the approved
contract, preserving custom values by System_Key through projection rebuild.
Keep closed/missing/invalid authority, cancellation, permissions, native/layout,
full processor and relevant regression gates open until separately proved.

## Approved implementation and qualified GREEN

New, unpromoted candidate `deploy/validation-admin-reconcile-01` implements
D14-RECON-01. Admin captures warehouse/root, uses exact event/key selection and
queues only the confirmed negative delta with narrative. Core provides a read-only
lease over the existing captured authority; the new headless Domain module owns
history calculation and event/key validation. Invalid wrong-entity/non-SHIP/
positive/missing-narrative corrections are refused. Current attributes are
preserved; a separate projection helper restores unknown columns by exact key.
Reviewed caller searches preceded removal of obsolete private SKU/ROW helpers.
Existing unrelated payload compatibility paths are not acceptance evidence.

Run the focused command above with `-DeployRoot deploy/validation-admin-reconcile-01`
and `-Phase GREEN`. All **55/55** checks pass in the exact RED
order, including actual queued application and duplicate-event idempotency.
Prompts still use deterministic test seams; native UI acceptance is not implied.

- Build: `reports/runtime/admin-reconcile-build-01`; five cold compiles and
  Operations cold start pass. Four changed/two added/315 unchanged components;
  five package pins recorded. Prior candidate and local settings preserved.
- Focused controller: `reports/runtime/admin-reconcile-controller/ce61a0f0e95945e3bab03c65531698cd`.
- Worker: `reports/runtime/slice4be-admin-reconcile/c93d48330ba94846bae8f11dac40f314/green.json`.
- UTC 04:40:51.7525561–04:41:36.2118577; delayed audit through04:41:48.2118577.
  Five instrumented compiles, exact log/receipt agreement, package/tooling/settings
  preservation and normal unassisted shutdown pass; zero native failures.
- Smoke: `reports/runtime/admin-reconcile-regression/smoke-cd6638d9a252442abc198d28c4e30172`.
  Exact prior86 checks pass, both native exits unassisted, settings/packages/tracked
  reports preserved and zero delayed native failures. UTC04:42:19.8786641–04:42:42.5677651.
- Static: `reports/runtime/admin-reconcile-green-static-01/ratchet-verification.json`.
  328 components/6379 procedures/139222 lines (four fewer total); all28 caps hold,
  InventoryApply2882→2849 lines. Dynamic9/45 and duplicate192 counts unchanged;
  three schemas and473 PowerShell scripts validate.

No deployment promotion, new recording ID, complete Reconcile acceptance or
4be-A closure. Next prove authority/permission/cancellation edges and native
dialog layout/input, then integrate the reachable control's recording contract.

## Source-harness dependency correction

The first chain attempt (`admin-reconcile-regression/chain-b04b3da00be9476ca384ba3ccf6348d5`
under runtime) stops at4/1 before live roles: the Create Warehouse source harness
omits the new projection helper, producing a compile error. An owned VBE capture
identifies the missing module. This is a harness failure, not behavioral RED.
Owned dialog dismissal and VBE Reset allowed original cleanup; packages/settings/
tracked reports were restored. An attempted native close was refused by Excel;
no force termination or Excel COM reattachment was used. Cursor target and restore
were verified for Reset. This attempt does not qualify for normal shutdown.

All16 explicit source-harness import lists that include InventoryApply now also
include its two new Domain dependencies. Each change adds only those two imports;
all scripts parse. No product contract or package changes follow from this repair.

Fresh chain is qualified GREEN: **32 chain /48 live-role /15 Create Warehouse**
checks in exact prior order, on the same five package hashes. Controller:
`reports/runtime/admin-reconcile-regression/chain-9fa6e9d72f104e2f905474e3766860a7`.
UTC04:50:56.2396123–04:56:46.9596530; delayed audit through04:56:58.9596530.
No native failures; Excel exits without assistance and settings/packages/tracked
reports are preserved. The final static audit takes several minutes while hidden
Excel processes remain; do not mistake active audit work for stuck shutdown.
This qualifies the candidate's existing full-chain regression, not all4be-A gates
or a root-cause resolution of earlier intermittent Shipping/native failures.
Legacy source tests with SKU/ROW fixtures and broader Reconcile native/role/
authority cases are not claimed as covered by this run.

## Authority preservation and interactive authentication follow-up

`Slice4beAdminReconcileAuthority.ps1` extends the same callback test on unchanged
`validation-admin-reconcile-01`. No runtime implementation/package changes.
Generated fixture changes and credential material stay private; reports contain
only check identities/counts/booleans. Only prompt answers/notices and an unsaved
Domain read-lease observer are instrumented; lookup, queue and authorization run
through their actual owners.

Final **128 PASS /2 behavioral FAIL**,130 checks, retains the original55 checks
in order. Passes cover closed read-only lease/closure, missing file not recreated,
invalid schema and duplicate key refusal, exact memory/formula/shape and saved-state
preservation of borrowed clean/dirty workbooks, cancellation, fractional selection,
reader denial, inventory bytes and unrelated warehouse Config. Exact fixed refusal
messages distinguish intended rejection from a caught generic exception.

The final fixture explicitly sets a signed-out Reader user-ID hint, which is denied
because it lacks ADMIN_MAINT. That alone is insufficient: the next case signs out, sets only the
local user-ID hint to a generated provisioned Admin, independently asserts
`modAuth.IsSignedIn=False`, then submits through the actual callback. It queues an
event and reports queue success. Failing checks:
- `Reconcile.Authority.SignedOutStaleIdentity.ExpectedCallbackConclusion`
- `Reconcile.Authority.SignedOutStaleIdentity.NoQueuedEvent`

Cause: Admin uses `ResolveCurrentUserId` → `QueuePayloadEvent` → `QueueEventCore`,
which calls `HasProvisionedCapabilityForSystem` without authenticating the live
interactive actor. Do not change the generic headless API to satisfy this UI test.
v4.11 D8-RECON-01 proposes a captured interactive session/user/target Core gate,
preserving Admin local target policy, read-only inspection and processor behavior.
Approved2026-10-12; no authentication correction is implemented. Align the
ordinary signed-out notice assertion with the approved interactive gate before
implementation; do not preserve the current fallback merely to retain that text.

- Controller: `reports/runtime/admin-reconcile-controller/8946a114524e4cd6b61462e637735fae`.
- Worker: `reports/runtime/slice4be-admin-reconcile/7bdf6fa0cdcd452f839e88dd56b04231/red.json`.
- UTC05:09:35.9621962–05:10:41.6647610; clean delayed audit through05:10:53.6647610.
  Five instrumented compiles, exact log/receipt agreement, normal unassisted closure,
  settings/tooling/package preservation; no native failures.

Two preliminary authority runs passed112, the second with exact refusal notices.
The123/1 permission diagnostic assumed a signed-out-session notice, but the old
API actually checks capability for a fallback identity. It is not the final RED;
the stale provisioned-hint case above proves the actual queued-write failure.
Reuse prior cold compile/smoke86/chain32-live48-create15 evidence for the unchanged
candidate; no new native UI, comprehensive role/binding or4be-A acceptance follows.
The earlier130-check run (`711215bbddf6424b87d3a90452c3d23e` controller) has identical
check order/outcomes; the final fixture pins both signed-out hints to avoid any
dependency on the Windows account name. The controller now defaults to the same
current candidate used by the explicit deployment argument in these runs.
Static `reports/runtime/admin-reconcile-authority-static-02` holds all28 caps,
unchanged328/6379/139222 component/procedure/line counts, dynamic9/45 and duplicates192;
three schemas and474 PowerShell scripts validate. No runtime rebuild was needed.

## Native presentation RED and authentication approval

The user approved D8-RECON-01 on2026-10-12. Record the captured live session/user,
warehouse/station and ADMIN_MAINT Core gate without changing Admin local targets
or provisioned headless processing. Extend mid-dialog mutation RED before that fix.

`-NativePresentation` adds20 real owner-applied SHIP entries, invokes the actual
packaged callback and forwards its selection seam to native VBA.InputBox. The
observer identifies the exact generated Excel process and dialog, reads its real
Static control, measures wrapped text against client bounds, captures only the
owned foreground window, and dispatches the native Cancel button. Reports expose
booleans/counts only; screenshots contain disposable fixture data and stay ignored.

Qualified **134 PASS/5 FAIL**,139 checks, retains all130 previous outcomes:
- two existing signed-out stale-identity failures;
- `Reconcile.Native.LastChoiceAndExactKeyPresent`;
- `Reconcile.Native.SelectionTextFitsVisibleDialog`;
- `Reconcile.Native.HistoryTimestampReadable`.

Prompt length is1023 characters. The viewed capture ends partway through choice7;
choice20 is absent, visible text is clipped, and timestamps are numeric serials.
Exact first key, native Cancel, unchanged queue count/inventory bytes/quantities/
custom values/saved state pass. This does not cover native confirmation/entry.

Controller `reports/runtime/admin-reconcile-controller/4f7357a16e3d4ca89321327abf02d6f7`;
worker `reports/runtime/slice4be-admin-reconcile/a73a7b2a2b75436b9790572808f749b6`.
UTC05:20:53.3901647–05:22:07.9668259; clean delayed audit through05:22:19.9668259.
All five instrumented compiles, log/JSON identity and normal preserved closure pass.
Static `admin-reconcile-native-static-01`: unchanged328/6379/139222, dynamic9/45,
duplicates192, all28 caps, three schemas and475 PowerShell parsers pass.

The first observer attempt failed because a helper was dot-sourced only inside
an installer function; move it to the script scope. That is not behavioral RED.
The next138-check capture exposed the date problem; the final139 adds its assertion.

## Approved authentication and timestamp corrections

D8-RECON-01 was approved in chat and synchronized in v4.11/Plan022/Controls before
implementation. Expanded callback RED on `validation-admin-reconcile-date-01` is
**160/12**,172 checks. It proves sign-out during selection/confirmation, re-sign-in
as the same Admin, switching to Reader, and switching warehouse through real Core
owners. Mutation preconditions pass. It also proves one permitted local-fallback
submission. Generated secrets stay in private transient fixture state, never reports.

RED controller `admin-reconcile-controller/d20e882850aa447a82ff94d2bcd761d6`;
worker `slice4be-admin-reconcile/6da0b92b03fb40329c75135b7bbdb655`, under reports/runtime.
UTC05:31:59.5491947–05:33:32.8488052; clean audit through05:33:44.8488052.
All original55 remain GREEN. The ordinary signed-out notice assertion intentionally
changes from capability denial to the approved sign-in refusal; no GREEN regression
is hidden. Five compiles and preserved normal closure pass.

`Core.modAdminReconcileAccess` captures transient session/user/target context and
validates it with live `modAuth.CanPerform` before Admin's queue helper proceeds.
Admin captures the authenticated user, not a local/Windows/Office hint. The existing
headless writer/processor API is unchanged. The first access candidate passed170/172:
`modRoleUiAccess` also enforces NAS for ADMIN_MAINT and wrongly rejected local use.
Do not reuse that wrapper here; direct live Core capability authorization preserves
D8-RECON-01. The explicit Phase6 source import list includes the new dependency.

Current unpromoted candidate **`deploy/validation-admin-reconcile-access-02`** passes
all **172/172** ordered identity/authority checks. The combined native gate is
**179 PASS/2 FAIL**: only LastChoiceAndExactKeyPresent and SelectionTextFitsVisibleDialog
remain RED. Numeric Domain dates now render through the existing ISO-like display
format in Admin. The reviewed physical capture shows readable timestamps; native
Cancel queues nothing and preserves quantities, custom values, saved state and bytes.
Two earlier date-only captures were refused foreground activation; neither was
accepted as visible evidence. The final capture verifies foreground before/after.

- Controller: `reports/runtime/admin-reconcile-controller/2c63b04cfb2646f799198b0b1c20620e`.
- Worker: `reports/runtime/slice4be-admin-reconcile/cbd21ccdd85b4634acd0d83340611f72`.
- UTC05:38:03.1461052–05:39:46.4077465; clean delayed audit through05:39:58.4077465.
- Build: `reports/runtime/admin-reconcile-access-build-02`; five cold compiles and
  Operations startup pass. Relative to the date candidate, only Admin Reconcile
  changes and the Core access helper is added;320 components are unchanged.
- Smoke: `admin-reconcile-access-regression/smoke-0ef63223a1a745be81f84cf7f5f786b2`
  under runtime; all86 prior ordered checks, both unassisted exits, preservation
  and clean delayed native audit pass.
- Static: `reports/runtime/admin-reconcile-access-static-02`;329 components,
  6381 procedures,139265 lines; all28 caps, unchanged dynamic9/45 and duplicates192,
  three schemas and475 PowerShell parsers pass.

No deployment promotion, new tracking ID or4be-A closure. Native history selection
needs complete readable/reachable choices; confirmation/entry and comprehensive
recording/guide proof remain open. Do not shorten exact keys or silently reduce20
choices to make the native RED pass.

Final same-candidate chain: **32 chain/48 live-role/15 Create Warehouse**, exact
prior order, unassisted normal closure, settings/packages/tracked reports preserved
and zero delayed native failures. Controller under runtime:
`admin-reconcile-access-regression/chain-81a8cecc8e994d6bbf4be47c1a61ab5f`.
UTC05:40:59.8755927–05:46:37.1591580; audit through05:46:49.1591580.
This is regression evidence, not complete4be-A or full Release1 user acceptance.
