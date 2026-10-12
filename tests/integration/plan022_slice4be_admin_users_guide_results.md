# Slice 4be-A: Users continuous recording and independent guide

Last verified2026-10-12 UTC. **GREEN129/129;4be-A remains incomplete.**
Unpromoted candidate: `deploy/validation-admin-users-context-01`.
This adds test coverage under existing D18/D8-USERS-01, with no production,
permission, catalog or wording change. The first focused run passed; no behavioral
RED or product correction is claimed.

`Slice4beAdminUsersGuide.ps1` records actual Create/Update, Deactivate and Delete
form handlers in three independent sequences. Each command enters Core and saves
once; fresh read-only Auth checks prove all nine business results. The final run
injects a real activity-file Open failure for Delete's SAVED observation after its
business save. It retains one fixed notice, truthful incomplete capture and the
two earlier successes. Test seams observe persistence and handle confirmation;
they do not supply command results or fabricate activity.

Actual Viewer Start/Stop, publication, guide authoring, expected-step selection,
guide pairing, Evaluate and presentation controls establish:

- Exact action/context records, immutable journal chains and prior activity,
  no private target-account/credential payload, and exact publication membership.
- A three-step SAVED guide with exact source provenance; explicit Admin guide
  permission and Reader authoring refusal.
- Independent successful evidence concludes; interrupted evidence stays Incomplete,
  matches only its first two successes and leaves the attempted Delete unmatched.
  Both results bind the exact guide, observed journal and publication. Neither
  borrows source-run success or asserts inventory Domain application.
- How-To, Diagnostic and Compare preserve the same evidence. Six reviewed captures
  show readable multiline instructions, provenance and saved conclusions; scrolling
  exposes the successful matches and interrupted terminal outcome mismatch.
- Original Auth, other warehouse authority, inventory and saved host bytes survive.

## Qualification

Controller `reports/runtime/auth-read-controller/4472ac183f7b4613b8531f92200d7ca8`;
worker `reports/runtime/slice4be-auth-read/39f007c143a440b18034dd924510417c`.
Run21:16:48.4117457-21:25:49.4623907UTC,129PASS/0FAIL, five instrumented compiles,
exact log/JSON agreement, settings/package/tooling preservation and unassisted
Excel closure. Delayed native audit through21:26:01.4623907UTC found no failures.
The seventh capture retains the shared Settings Save regression.

`reports/runtime/admin-users-guide-static-01`:342 components/6428 procedures/
140068 lines; all28 size caps, dynamic9/45 and196 duplicate groups unchanged.
Three schemas and all507 tools/tests PowerShell files validate.

Five package hashes exactly match the prior493-check Users qualification. Product
source is unchanged. Reuse its five cold compiles,493 ordered regression checks,
chain32/live-role48/Create15 and normal closure/native evidence, as detailed in
`plan022_slice4be_admin_users_context_results.md`; these were not rerun here.
The new opt-in mode preserves the existing default authority/recording gate.

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File tests/tooling/Test-Slice4beAuthReadOnly.ps1 -DeployRoot deploy/validation-admin-users-context-01 -Phase GREEN -AdminRefresh -AdminWriteDiagnostic -RequireWriteAuthority -CoreAuthority -UserRecording -UserGuideOnly -CaptureEvidence
```

Runtime evidence and generated credentials remain ignored. Ordinary native Users
warning/confirmation, existing form clipping, other controls and further context/
reentrancy boundaries remain open. This is recording/guide evaluation, not replay.
Seed recording boundaries and independent guides, broader4be-A coverage and early
B0 still require completion. D18-LOCAL-01 remains unresolved.
