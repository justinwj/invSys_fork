# Create New Warehouse authority and reopened publication

4be-A / Admin prerequisite,2026-10-13. **D8-CREATE-01 is approved, implemented and
scoped-qualified.** Unpromoted candidate:
`deploy/validation-create-authority-03`; recording catalog remains40/162.
Create recording, guide conclusions and comprehensive4be-A acceptance remain open.

## Contract and implementation

The real Admin Ribbon/form handlers call typed, headless Core commands. Each
write checks the live session generation and ADMIN_MAINT at the captured source
warehouse/station. Sign-out, new session, changed warehouse, reader rights and
persisted revocation refuse writes; a local user-ID hint cannot replace sign-in.
Authorization is repeated after read-only publication validation. No confirmation
exists in these current handlers; any future confirmation must precede that check.

Local creation and publication retain separate results. **Publish Existing
Warehouse** opens the recovery path in a new form/session, reads saved identity
from the selected existing runtime and ignores new-warehouse entry fields. It
makes zero bootstrap calls and preserves local Config/Auth/inventory bytes.
Missing/invalid runtime refuses without repair. Ordinary Create still rejects
duplicates. Downstream diagnostic sinks receive fixed text during these commands,
not entered paths/fields. No new recording IDs or replay authority are introduced.

The first candidate exposed an existing side effect: inventory-path resolution
selected the destination interactively during bootstrap, signing out the source.
The inventory bridge and processor now prefer their explicit transaction root
before asking for an interactive target. This preserves the approved source
context without reauthentication or restoring a stale session. Existing headless
bootstrap remains separate. Headless test follow-up Seed explicitly binds its
own generated runtime rather than relying on that removed selection side effect.

Summary formatting moved unchanged into an Admin helper to keep the Create form
small. The new footer button remains inside the form at minimum/expanded sizes.

## D13 evidence

All roots below are ignored `reports/runtime/`; no runtime payload is committed.

- Historical owner RED45/2: controller
  `admin-create-owner-controller/bc3531d7e5b04df3af169524d0c2b3f1`.
- Expanded RED63/23 before product edits: controller
  `admin-create-owner-controller/54c4bc6167674e66a36d24891740e3a8`.
  Repeated with counters at actual Core calls:
  `admin-create-owner-controller/0a637996150d49f48a314969490389ee`, worker
  `slice4be-admin-create-owner/24468886a8e949fcaab7333015217ae3`.
- Diagnostic sink RED106/4 on candidate02: controller
  `admin-create-owner-controller/abd69e0b5eae4355810c82d3cf6cacff`, worker
  `slice4be-admin-create-owner/3b0f1dc78a2c440fb0dc9a02974e265e`.
  Observation at the actual file write exposed entered identity in PerfBegin/Mark
  run IDs. Moving redaction into the shared writer protects every performance entry.
- Final focused GREEN110/110 on candidate03: controller
  `admin-create-owner-controller/1defae167b5a4a8282fc2ecb56d7ec36`, worker
  `slice4be-admin-create-owner/fb1b42520e7c4242bf50338c074af48c`.
  Every original45 pass and all86 expanded RED check identities are now GREEN;
  `create-authority-check-comparison-02.json` and `-03.json` record preservation.
  Five instrumented compiles pass. Packages/tooling/settings are preserved;
  Excel exits normally. Delayed native audit through22:02:21.454UTC finds zero
  Excel error/hang events.
  Three reviewed final-candidate captures show denied retry and successful
  reopened publication at minimum/expanded sizes.
- Purpose/layout GREEN36/36:
  `warehouse-purpose/f21f240cf9a44fabadba3af4375e69fc`, with two reviewed captures.
  All35 prior checks remain; one added check establishes a real source Admin.
  The disposable fixture credentials are never emitted and are removed after
  Excel exits. Source initialization reuses the existing credential helper.
- Candidate03 passes Operations cold start and all five cold compiles:
  `create-authority-compile-03.log`, `create-authority-compiled-03.json`.
  Compared with qualified root-recording-04,333 compiled components are unchanged,
  five are changed and three are added. Operations and both Domains are unchanged.

Focused coverage includes actual callback entry, Operational/Training creation,
exact generated Receiving opening, duplicate/invalid/cancel refusal, successful
publication, real partial publication failure and retry, four stale-context retry
cases, reopened saved identity, direct Core source-vs-destination authorization,
context loss after the read, missing/invalid runtime, saved rights revocation,
source-session preservation and non-vacuous diagnostic redaction.

Reproduction:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File tests/tooling/Test-Slice4beAdminCreateWarehouseOwner.ps1 -DeployRoot deploy/validation-create-authority-03 -Phase GREEN -CaptureEvidence
powershell.exe -NoProfile -ExecutionPolicy Bypass -File tests/tooling/Test-Slice4beWarehousePurpose.ps1 -DeployRoot deploy/validation-create-authority-03 -Phase GREEN -CaptureEvidence -CheckLayout
```

## Qualification and exclusions

Final candidate03 regression:

- Chain32/32, live-role48/48 and Create15/15:
  `create-authority-regression-03/chain-81338c93b9b8465b95e60449d6f0f17f`.
- Packaged smoke86/86:
  `create-authority-regression-03/smoke-d82ab71cf100403b80aaa68c38594d45`.
- Recording/Viewer80/80:
  `create-authority-regression-03/recording-fc267ad13bfc47b5970fc985f1c07d91`.

Chain, smoke and recording retain the preceding qualified check identities/results, with
normal unassisted closure and preserved packages/settings/tracked reports.
The delayed final audit through22:15:04.999UTC finds zero Excel1000/1001/1002
events and Excel closed (`create-authority-native-audit-02.json`). Passive desktop
monitoring reports no error5; visible captures passed their input guards. Earlier
owned-window cleanup is excluded below, not attributed to the final chain.
Static `create-authority-maintenance-04` reports348 components,6445 procedures,
140438 lines; literal/dynamic calls9/45 and duplicate groups196 are unchanged.
All28 oversized modules stay within their existing limits. The Create form shrinks
664 ->588 lines; new helpers meet module/procedure limits. Three JSON schemas and
524 PowerShell parses pass. No code deletion or growth exception is introduced.
Final diff cleanup removes one trailing blank line from the presentation helper
(final source140437 lines); no executable code changed after these gates.

Candidate01's49/9 runs localized lost source context; they are not acceptance.
Candidate02's104/2 attempt had an adapter error: a bare handler name before a
colon became a VBA label. Explicit `Call` invokes the actual reopened-publication
handler. Candidate02's later110 observed only Diagnostic/End call sites and did
not establish complete log redaction; the sink RED and candidate03 GREEN supersede it.

The first chain controller `chain-3f836ba1bff4423dabdea56b8a644b46` failed in the
source integration harness, whose explicit imports omitted the new Core command
and dependencies. Native inspection identified the undefined module in modPerfLog.
The owned compile/quit dialogs were acknowledged and VBA Reset allowed cleanup;
no force kill or recovery COM attachment was used. This assisted attempt is excluded.
The import list is corrected. The next chain exposed headless Seed fixtures relying
on implicit target selection; both fixtures now bind their generated runtime.
No inventory assertion or authorization guard was relaxed.

Manual/native Ribbon clicking, ordinary modal interactions and folder-helper
acceptance are not established by these programmatic-handler captures. They remain
part of the wider release acceptance;4be-A and early Receiving B0 are not complete.
