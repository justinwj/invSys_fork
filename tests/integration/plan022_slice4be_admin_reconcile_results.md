# Slice4be-A: Shipment Reconcile identity prerequisite

2026-10-12. **191/191 GREEN: identity/authority, complete history, native selection and layout.**
Native confirmation/window-close coverage, recording and4be-A remain incomplete.
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

## Scrollable selection form GREEN

The pre-implementation native179/2 RED above protects the replacement of the
1023-character InputBox. Admin now owns `frmShipmentReconcileSelection`: complete
locked/wrapped history, numbered entry, Inspect and Cancel, using the existing
callback validation and shared Core anchor manager. Default720x480, minimum540x360;
grow900x600, maximize and restore retain reachable actions and full history.
Set the caption in Initialize: the FRX-less packaged designer initially retained
`UserForm1`. No quantity, identity, authorization or recording rule changes.

Candidate **`deploy/validation-admin-reconcile-selection-02`** passes **191/191**.
All181 prior checks retain order, all179 prior GREEN remain GREEN, and both native
history failures pass. The original172 identity/authority prefix is unchanged.
Native Ctrl+Home/End reaches entry20; Escape invokes Cancel without queuing.
Native Tab/20/Enter invokes Inspect; a read-only observer verifies the exact
selected event/key reaches Domain and the existing no-higher-quantity conclusion
queues nothing. Inventory bytes, quantities, custom values and saved state hold.
Six foreground captures were directly reviewed, including readable entry20 at
minimum/default/restored sizes and the complete history at larger/maximized sizes.
The conclusion message is observed at the existing notice seam, not a native
confirmation screenshot. Native window-close is not separately accepted here.

- Controller: `reports/runtime/admin-reconcile-controller/825b1a529c504fc6bd54f0b8f3734765`.
- Worker: `reports/runtime/slice4be-admin-reconcile/4397cbda0d054a6f8920d6d8d0f5368d`.
- UTC06:24:13.5415385--06:26:09.6691404; clean delayed audit through06:26:21.6691404.
- Five instrumented compiles, exact log/JSON correspondence, normal closure and
  settings/package/tooling preservation pass. Prior191 GREEN controller
  `0931ca9cb13a4a9891dfc1a3e666ad34` is superseded only for stronger resized captures.
- Build `reports/runtime/admin-reconcile-selection-build-02`: five cold compiles,
  Operations startup, only the intended Admin module change/new form,321 unchanged
  components, preserved prior candidate and clean delayed native audit.
- Smoke `admin-reconcile-selection-regression/smoke-42cbea0cf15b4a3e9e3996b2649f6959`
  under runtime:86 retained ordered checks, both unassisted exits, preservation and
  clean delayed native audit.
- Static `reports/runtime/admin-reconcile-selection-static-03`:330 components,
  6390 procedures,139369 lines, all28 module caps, unchanged dynamic9/45, three
  schemas and476 PowerShell parsers pass. **Reviewed exception:** duplicates193
  (+1) is the one-line `UserForm_Layout` event adapter shared with `frmItemSearch`;
  both delegate layout to the existing anchor manager. Retain required event
  entries; do not alter scanner rules or add duplicate layout implementation.

Diagnostic limits: hidden Excel suffered intermittent `EXCEL.EXE` access violations
before native input. Traced controller `8dbd4bcc085f4725855ae48cbaba27ff` wrote only
ObserverReady before the crash. Startup visibility alone did not survive hidden
fixture windows (`af848eed21194f07bf7cf375865da1c6`). The native harness now keeps an
unsaved disposable operator workbook visible and refuses entry if that prerequisite
fails. Both qualified191 runs close normally without native failures. This matches
operator context; it does not claim to repair Excel's hidden-host fault. Assisted
recovery runs are diagnostic only and never accepted as GREEN.

Harness corrections are not product RED: caret position3210 versus value length3229
reflects19 CRLF separators; compare native position to the same control's clamped
end only after capture. Traverse Home then End after resizing to prove viewport
reachability. Keep test assertions outside returned receipt objects and parenthesize
the Boolean Start-Job argument so Cancel cannot become Inspect. No new tracking ID,
deployment promotion or complete4be-A acceptance is claimed.

Final same-candidate regression:32 full-chain/48 live-role/15 Create Warehouse
checks retain exact prior order. Controller under runtime:
`admin-reconcile-selection-regression/chain-98c7d4114a4048b6997d451769713736`.
UTC06:26:33.2352910--06:32:19.1695957; zero delayed native failures through
06:32:31.1695957, normal closure and settings/packages/tracked reports preserved.

## Native confirmation and window-close GREEN

The unchanged `validation-admin-reconcile-selection-02` passes **203/203**,
retaining the exact191-check prefix. This extends acceptance under D14-RECON-01
and D8-RECON-01; no product code, package, wording or architectural contract changes.
The two preliminary readability failures were observer defects, not behavioral RED.

Native WM_CLOSE reaches the selection form's existing close handler and queues
nothing. After an owner-applied adjustment, native selection/Inspect opens the
actual confirmation with the exact selected EventID/System_Key, current4,
after-event1 and delta-3. Native Escape cancels; Unicode keyboard entry followed
by Enter submits one exact correction with the required synthetic narrative and
no legacy identity field. Neither action applies inventory: bytes, both entity
quantities, custom values and saved state remain unchanged. The existing notice
seam verifies the callback conclusion; final notices are not native captures.

The two reviewed foreground captures show complete blank and typed confirmation
states. Exact prompt text is checked separately. Its270 characters measure150px
inside a150px-high client area using the actual font inside Excel. A bounded,
test-only timer observes that dialog and is stopped after observation/cancellation;
it never supplies input. A GDI font handle cannot be used from the external helper
process. Native edit readback uses WM_GETTEXT, because GetWindowText cannot retrieve
another process's edit contents. Missing measurements or mismatched input fail the
harness; submission is withheld when native input readback differs.

- Controller: `reports/runtime/admin-reconcile-controller/933d47cb6a244f2b8a2ad59ea39de8c2`.
- Worker: `reports/runtime/slice4be-admin-reconcile/0ed7dfc863ca47038e274b7da114e850`.
- UTC06:51:19.4366473--06:53:27.8039640; clean delayed audit through06:53:39.8039640.
- All five instrumented compiles, exact log/JSON agreement,191-check prefix,
  settings/package/tooling preservation and normal Excel closure pass.
- Reuse the same-candidate cold compile,86 smoke,32 chain/48 live-role/15 Create
  Warehouse and layout evidence above; no new package was built or promoted.
- Static `reports/runtime/admin-reconcile-confirmation-static-02`: unchanged330
  components/6390 procedures/139369 lines,28 caps, dynamic9/45 and duplicates193;
  all three JSON schemas and478 PowerShell parsers pass. The prior reviewed
  one-line layout-adapter exception remains unchanged.

Recording/publication/guide coverage remains open. This is not4be-A acceptance.

## Reconcile recording and binding gate

D18's recorded control map precedes implementation: catalog34 adds Open, Inspect
and Cancel under ADMIN_SHIPMENT_RECONCILE, preserving all151 prior definitions.
Admin owns an invocation-scoped observer; Core owns policy, storage and conclusions.
The existing native narrative response belongs to Inspect. No entered values,
paths, inventory keys or quantities enter activity. CONFIRMED means one new queued
correction with an Inventory/Submitted reference, never Domain application or the
older corrected shipment. FAILED references, if present, remain singular/Unknown.

Qualified D13 progression (all paths below are under ignored `reports/runtime`):

- Absent-record RED:207/11 of218 on unchanged selection-02; original203 all GREEN.
  Controller `admin-reconcile-controller/a2e35baa35dc46b1b362a4d18cdb0256`, worker
  `slice4be-admin-reconcile/6f41e470aee44ad08f0406ccf37a6b4c`.
- Missing sign-out warning RED:223/1 of224 on recording-03; controller
  `admin-reconcile-controller/924d329b49364b1ca25f2d45acb81513`, worker
  `slice4be-admin-reconcile/17e6d7aa3ea842649770f527f571fefd`.
  The earlier214/4 run exposed an observer conflating tracking and business notices;
  that harness defect is not behavioral RED. Separate observations preserve both.
- Reference guard RED:241/3 of244 on recording-04. Core incorrectly accepted
  multiple CONFIRMED/FAILED references and a FAILED/Submitted reference.
  Controller `admin-reconcile-controller/7ad13190888345d3848c140e173acc8f`, worker
  `slice4be-admin-reconcile/f779b379d6ba464193ca42e1b4908470`.
  Recording-05 passes244/244; controller
  `admin-reconcile-controller/588a10b0aba64a958e0621ff199e687d`, worker
  `slice4be-admin-reconcile/3695df0eba66416880a4ab4aacbf9265`.
- Entry-binding RED:246/2 of248 on recording-05, retaining the exact244 prefix.
  A one-shot test-only hook performs real sign-out/re-authentication as optional
  Core recording begins. The callback wrongly captures that later session and
  queues a correction. Controller
  `admin-reconcile-controller/6640e35597e84e46b0bb2da19a3aa583`, worker
  `slice4be-admin-reconcile/4b13f8f524d244e7826b1a60b8dfa3e7`.
  Clean delayed audit through07:42:55.1338934 UTC. Business target/user/context
  capture now precedes optional recording, preserving approved D8-RECON-01.

The first entry test incorrectly required InvokeCallback's successful-queue
Boolean while expecting no queue. Its247/1 candidate06 result is not GREEN.
Correct the assertion to require False plus unchanged queue count; the qualified
246/2 RED above reruns that corrected assertion against candidate05.

The244 gate observes five actual native callbacks producing exactly20 records:
five Open/OPENED pairs, two Cancel/CANCELLED pairs, and Inspect NO_CORRECTION,
CANCELLED and CONFIRMED pairs. It checks unique correlation, original context,
fixed messages/effects, exact new-event references and CommandCompleted polarity.
Five mid-dialog context changes show one tracking warning each. Global recording
off, all three controls off and current-user recording off each leave the business
command working once, with no activity or false warning and unchanged inventory.
These checks do not yet prove storage failure, recording publication or guide use.

Final recording-06 passes **248/248** with the exact244-check prefix retained.
Controller `admin-reconcile-controller/7ddc1113764344a99d3dc0510a692a71`, worker
`slice4be-admin-reconcile/46d64de022df47908b3fffa27fac4ef9`, UTC07:42:55.6245589--
07:45:48.3790984; zero delayed native failures through07:46:00.3790984. All five
instrumented compiles, exact log/JSON correspondence, preservation and normal
closure pass. Entry reauthentication refuses with the existing exact message,
no queued event, no later-context activity and unchanged inventory.
Final foreground captures were reviewed directly: minimum-size history retains
entry20 and reachable controls; the typed native confirmation displays the full
selected event/key, quantities, delta and narrative. Other sizes and native
cancellation retain the passing measurements and handler checks.

Build `admin-reconcile-recording-build-06` cold-compiles all five packages and
passes Operations startup: five intended components changed, two added,318
unchanged. Prior candidate/settings are preserved; clean delayed audit through
07:35:56.5107972. The candidate is not promoted.

Static `admin-reconcile-recording-static-04`:332 components,6399 procedures,
139542 lines; all28 size caps, unchanged dynamic9/45, three schemas and479 parsers
pass. **Reviewed exception:** duplicates194 (+1) is the literal-erased three-ID
ControlIds list shared by separate Reconcile and Production recipe-order registries.
Their fixed identities differ; merging the registries would obscure ownership.
No duplicate business logic is accepted. Prior one-line layout-adapter exception
remains; no scanner rule or deletion authority changes.

Same-candidate smoke retains86/86 in exact prior order, with both unassisted exits:
`admin-reconcile-recording-regression/smoke-52989636304f4be29249cab44c6a0027`,
UTC07:46:00.7179731--07:46:21.2101993; delayed native audit is clean.
Settings retains202/202 in exact prior order, including policy/detail/preference,
restart, Operations isolation/binding and Admin Close:
`admin-reconcile-recording-regression/settings-d0dfdc5bdbf649debc236c3fdd405dc1`,
UTC07:46:53.6726159--07:51:31.5088476; clean audit through07:51:43.5088476.
All five instrumented compiles and settings/package/report preservation pass.
Settings' fixed catalog expectation changes33/151 to specified34/154; no check
is removed. Earlier scope-specific tests retain historical candidate expectations;
update their version assertion explicitly when rerunning them on a later catalog.

Build diagnostics: two hidden builds crashed in VBE7.DLL with native access
violation c0000005 before Operations construction. Neither is GREEN or desktop
error5. A disposable visible host succeeded; canonical `build-xlam.ps1` now offers
opt-in `-VisibleHost`, closes its unsaved host and retains the default hidden mode.
This does not claim to repair Excel. Raw diagnostic artifacts remain ignored.

First recording-06 full-chain run is not accepted: controller
`admin-reconcile-recording-regression/chain-4b98c7ad3e994ef992f1ae3d71c1e494`.
At07:53:58.8468184 UTC Excel crashed in ntdll.dll, c0000028, offset12d2f, while
`modTS_Shipments.RunShipmentsSentFormActionForTest` was active; the harness received
0x800706BE. Live checks stopped40/1, Create15/15, chain5/1. Root cause is unresolved;
this is neither product RED nor desktop error5. The replacement window matched a
generated fixture and closed normally through WM_CLOSE, without COM attachment
or force-kill. The waiting controller then restored settings/packages/tracked
reports; assisted cleanup is not GREEN. A separate fresh run is required.

Fresh same-candidate full-chain GREEN:
`admin-reconcile-recording-regression/chain-ce50c7639c4046caaeb6c85697667bdd`,
UTC07:57:28.0771557--08:02:59.4228247. All32 chain,48 live-role and15 Create Warehouse
checks retain the exact prior order. Normal unassisted completion, restoration,
package/report preservation and zero delayed native failures through08:03:11.4228247
pass. This qualifies the separate clean run; it does not diagnose or repair the
earlier intermittent Excel crash.

Local-target recording eligibility, active capture/publication/guide proof and
remaining Admin coverage stay open. No local business permission restriction,
recording exclusion, package promotion or complete4be-A acceptance is claimed.

## Recording storage failure and guide acceptance extension — 2026-10-12

Acceptance tests extend the unchanged, unpromoted `validation-admin-reconcile-recording-06`.
No runtime behavior or architecture changes; existing implementation RED/GREEN
remains above. New passing coverage does not manufacture another product RED.

Storage failure: **258/258 GREEN**, exact prior248-check prefix, five instrumented
compiles and normal preserved closure. Controller
`admin-reconcile-controller/f9e3e8ba5c6f4a8a8b2cdf3a8972f397`; worker
`slice4be-admin-reconcile/d294321895b0422283eab2479ba31d8b`.
UTC08:12:21.7209335–08:15:16.8407147; clean delayed native audit through08:15:28.8407147.
The disposable probe changes only the actual activity-store pending path so its
real file-open fails on a nonexistent directory. REQUESTED and CONFIRMED failure
each preserve one business execution, exactly one fixed tracking warning,
previous records, inventory bytes/quantities, exact keys and custom values.
Missing completion records never become invented success.

Harness corrections, neither product RED nor acceptance:

- Controller `b05fd9cb6dee471789648e2dbe4cb80e` ended257/1 because the deterministic
  selection seam does not show the form, so it cannot emit OPENED. Corrected
  expectation requires only the two actual REQUESTED records, then reran GREEN.
- Controller `045bbd1e354f4a199e1793bbefd46168` ended287/1 at guide authoring.
  Its new fixture omitted the explicit ACTION_PATH_MAINT grant. Recording and
  exact publication passed; guide acceptance requires the corrected rerun.
  Admin bootstrap alone must not imply guide-maintenance authority.
- Controller `9864c4c59f194763bc2f742fcea1a154`, worker
  `slice4be-admin-reconcile/734b1b8e00964250b6a0426ce9b2f975`, ended327/2. The interrupted run had a saved
  REQUESTED but no CONFIRMED outcome: zero missing steps, one terminal outcome
  mismatch and overall Incomplete/CAPTURE_INCOMPLETE. The reused Settings oracle
  wrongly expected an absent action. The final preservation check also depended
  on a caller-defined Retained helper. Corrected the Reconcile-only expectation
  and made that preservation helper local; no product change. All other view
  checks passed, but the run is not GREEN. Normal preservation and a clean
  delayed native audit through08:44:42.0301671 pass.

Static `admin-reconcile-guide-static-01` is regenerated and verified: all metrics
match recording-static-04 exactly (332 components,6399 procedures,139542 lines,
dynamic9/45,duplicates194), all28 caps and all three schemas pass. Final parsing
after harness corrections passes480 PowerShell files with zero errors. No new
bloat exception. Build, prior full-chain and current gate pins match all five
XLAMs; retain the same-candidate cold compile,86 smoke,202 Settings and
32 chain/48 live-role/15 Create Warehouse evidence above.

Final **330/330 GREEN**: controller
`admin-reconcile-controller/b984a1b8640d41e5a480b385a1acb998`; worker
`slice4be-admin-reconcile/6919b5a7216f4500bb462a24b0fe04f7`.
UTC08:44:58.0335704–09:02:46.9402452; clean delayed native audit through09:02:58.9402452.
Exact258-check prefix, log/JSON agreement, five instrumented compiles, preserved
settings/packages/tooling and normal Excel closure pass. Run with
`Test-Slice4beAdminReconcileIdentity.ps1 -DeployRoot deploy/validation-admin-reconcile-recording-06 -Phase GREEN -NativePresentation -CheckRecording -CheckGuidePaths`.

Three independent native callback/confirmation journeys generate their own
immutable journals and exact new correction references. Actual Admin publication
preserves every observation. Packaged guide controls author two ordered steps
from Source, then evaluate distinct Observed and Interrupted runs. Their saved
results are Concluded/COMMAND_COMPLETED and Incomplete/CAPTURE_INCOMPLETE, bound
to the exact guide, observed journal and publication. Neither claims Domain
application or borrows source-run success. The interrupted terminal attempt has
an unmatched outcome, not an absent action. Original journals/activity/publication,
inventory bytes, exact keys, custom values and operator canary are preserved.

All three presentation modes retain identical evidence and pass layout checks.
Directly reviewed six final `reconcile-guide-*.png` captures in the worker:
Observed How-To/Diagnostic/Compare/conclusion and Interrupted Compare/conclusion.
Full provenance, multiline instructions, saved conclusions and scrolling are
readable. Raw generated fixture captures remain ignored.

This completes the storage/recording/publication/command-conclusion extension,
not full Reconcile or4be-A acceptance. SourceEventsApplied guide proof,
local-target recording eligibility and remaining Admin coverage stay open.
