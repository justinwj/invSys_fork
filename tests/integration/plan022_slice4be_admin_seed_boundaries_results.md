# Slice 4be-A: Seed recording failure boundaries

Last verified2026-10-12 UTC. **GREEN181/181;4be-A remains incomplete.**
Test-only extension under existing v4.11 D18 Demo Seed observation requirements.
Candidate `deploy/validation-admin-users-context-01` is unchanged and unpromoted.
No product correction or meaningful behavioral RED is claimed.

`Slice4beAdminSeedBoundaries.ps1` exercises the actual Seed callback, owner, queue
and processor. Four real activity-file Open failures cover REQUESTED and terminal
storage for uploaded inventory creation and built-in no-op Seed. Optional recording
never repeats or prevents the owner: each callback invokes it once, retains its
success and shows one fixed sanitized notice. Honest partial evidence, unchanged
prior activity, both Auth authorities and no pending artifacts are checked.

Sign-out inside BeginAction and FinishAction is tested on the existing no-op path:
one owner invocation, no inventory/inbox mutation, one notice, and no stale terminal
attribution. Reentering BeginAction during the outer REQUESTED write is quiet and
produces no nested action; the outer command still applies once and retains its
correlated pair and exact source event reference. These are specific boundaries,
not acceptance of every actor/warehouse change or business authorization scenario.

Three uploaded commands are independently checked from persisted inventory:
close only the generated fixture's unsaved projection, hash, reopen read-only,
verify the exact entity quantity/nonempty System_Key and exact APPLIED receipt,
close without saving and verify identical bytes. Required identity uniqueness
and one newly queued/applied event are also checked through the retained owner probe.

## Qualification

Controller `reports/runtime/admin-seed-outcome-controller/e7347e7f85594ac7b0c7187e3137e9ab`;
worker `reports/runtime/slice4be-admin-seed-outcome/3751bfc5015f4b72ab12f11f50989b86`.
Run21:40:26.8709882-21:44:05.9066578UTC,181PASS/0FAIL, retaining all114 earlier Seed
checks and all177 checks from the intermediate measurement run in order.
Five instrumented compiles, exact log/JSON, settings/package/tooling preservation
and normal unassisted Excel closure pass. Delayed native audit through
21:44:25.9066578UTC has zero Excel failure events.

Static `reports/runtime/admin-seed-boundaries-static-01`:342 components/6428
procedures/140068 lines;28 caps,dynamic9/45 and196 duplicate groups unchanged.
Three schemas and all508 tools/tests PowerShell files validate. Five package hashes
match the prior Users493 qualification. Reuse its cold compiles and chain32/live48/
Create15 evidence from the Users context receipt; no product/package change requires
a repeated chain. The new boundary option leaves existing default gates unchanged.

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File tests/tooling/Test-Slice4beAdminSeedOutcome.ps1 -DeployRoot deploy/validation-admin-users-context-01 -Phase GREEN -WorkflowEdges -Recording -RecordingBoundaries
```

## Excluded diagnostic attempts

Initial probe setup failed before behavior (`e873564c7f94426a8680973fa6f37722`),
then a PowerShell case-insensitive variable collision inserted an Admin probe
into Core (`92413239655240559396c2902cee6938`). Its owned compile dialog required
normal foreground/close dismissal before cleanup; it is not unassisted evidence.
Renaming the probe parameter fixed the test. These are not product RED.

The intermediate174/3 run (`2b76aed12a1f4815ad1c97486535dbf3`) queried the live
projection after publication could dirty it. That cannot establish saved receipts.
The first fresh-read retry (`05e0ca3f533746c4924307f4e74cc258`,117/1) tried hashing
before Excel released the file. Final181 uses fresh read-only disk evidence and
retains all prior behavioral assertions. All attempts restored settings/packages;
none required product changes. Runtime reports and credentials remain ignored.

Ordinary native tracking-notice presentation is not proven by the explicit notice
seam. Seed continuous recording, publication, independent guides, remaining Demo
controls, broader context cases and early B0/4be-A completion remain open.
