# Slice4be-A: Shipment Reconcile identity prerequisite

2026-10-12. **RED established; no runtime fix or 4be-A acceptance.**
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

Next: implement the exact-key read/queue path and Domain guards under the approved
contract, preserving custom values by System_Key through projection rebuild.
Keep closed/missing/invalid authority, cancellation, permissions, native/layout,
full processor and relevant regression gates open until separately proved.
