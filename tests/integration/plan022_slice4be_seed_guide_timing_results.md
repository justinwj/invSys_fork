# Slice 4be-A: Seed Create Guide responsiveness RED

Verified 2026-10-12 Pacific. No runtime change or performance GREEN. The approved
D18-LOCAL-01 eligibility change is separate and remains unimplemented.

## Focused evidence

On unchanged `deploy/validation-admin-users-context-01`, run:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File tests/tooling/Test-Slice4beAdminSeedOutcome.ps1 -DeployRoot deploy/validation-admin-users-context-01 -Phase GREEN -WorkflowEdges -Recording -RecordingBoundaries -GuidePaths -GuideTiming
```

`-Phase GREEN` retains the existing functional assertions. This is a behavioral
responsiveness RED: **202 PASS / 1 FAIL**, solely
`SeedTiming.CreateGuideWithinFiveSeconds`. All prior181 checks remain in order.
The five-second limit is an engineering regression budget for this generated
two-step fixture, not a warehouse-wide service-level contract or permission waiver.

Controller `reports/runtime/admin-seed-outcome-controller/32b4e020c39644e8bf599ecad5c2c413`;
worker `reports/runtime/slice4be-admin-seed-outcome/736aec8b08904ab9832fc7b4a4702ddf`.
Run23:46:43–23:55:51UTC; clean delayed native audit through23:56:11UTC.
Five instrumented compiles, saved disposable package copies, source publication,
two observed steps, unchanged source evidence/authorities/host, restored settings,
original package/tooling pins and normal Excel closure pass. Exact log/JSON agree.

Actual packaged `btnCreateGuide` handler elapsed **118,234ms**. Fixed-label traces
within the exact Operations callback interval show:

| Core procedure | Calls | Inclusive milliseconds |
| --- | ---: | ---: |
| modActivityPolicy.ReadPolicy | 16 | 116251 |
| modTrackingPolicyModel.Project | 16 | 104982 |
| modConfig.LoadConfig | 16 | 5579 |
| modRecordingReader.ReadRun | 8 | 392 |

Nested totals must not be added. Cross-XLAM Operations durations include Core
work. Policy projection dominates; source inspection finds repeated worksheet
cell/header reads inside current-control/history loops. A per-call array snapshot
is a proposed optimization, not yet implemented. Preserve fresh policy reads,
validation, comparison/order semantics, session binding and unknown columns.
Configuration loading alone exceeds five seconds across16 calls; projection
optimization might not meet the budget. Do not silently relax the test threshold.

## Diagnostic and exclusions

Initial successful profile: controller `df265c17d2b54deea73bab884f9eae2a`, worker
`133095c25ba84ba49e621ca6474dd64d`;202/202 before adding the timing assertion,
118963ms. Prior181 retained, restoration/normal closure pass; delayed native audit
clean through23:45:22UTC. This established slowness, not performance acceptance.

Earlier controller `c2bdbc93f5ac4d89944b1316819907c0`, worker
`6e229a0e36c64562917f71fd30b60976`, failed setup because tracing requires writable
disposable copies. It restored/closed normally. Not behavioral RED. The timing
mode now uses existing SavedCopies isolation and saves compiled probes before
fixtures; original packages remain untouched.

Opt-in timing keeps the normal283-check Seed guide path intact and records only
one source sequence for profiling. It does not replace independent guide
conclusion, full regression or visible acceptance evidence. Prior GREEN283 remains
qualified by `plan022_slice4be_admin_seed_guide_results.md` on these unchanged
packages. No new cold compile, station deployment or full4be-A acceptance claimed.

Static `reports/runtime/seed-guide-timing-static-01` matches the prior Seed guide
baseline:342 components,6428 procedures,140068 lines,28 unchanged module caps,
dynamic calls9/45 and196 duplicate groups. Three JSON schemas and all511
PowerShell parsers pass. Unrelated documentation changes retain their prior hashes.
