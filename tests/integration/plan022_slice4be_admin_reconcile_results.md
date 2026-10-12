# Slice4be-A: Shipment Reconcile identity prerequisite

2026-10-12. **RED established; no runtime fix or 4be-A acceptance.**
Candidate `deploy/validation-admin-entry-02` unchanged, catalog33/151.
Governing contract: v4.11 D14 exact identity/unknown columns, headless owners and
read preservation. D14-RECON-01 quantity clarification is proposed, not approved.

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

Next: approve D14-RECON-01, add protecting repeated/multi-entity and Domain guard
cases, then correct the exact-key/read path. Retain current RED and all prior
GREEN gates; packaged/native/role/application/layout and regression proof remain.
