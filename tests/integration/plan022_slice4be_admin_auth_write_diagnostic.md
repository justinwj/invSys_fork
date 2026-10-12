# 4be-A Users & Roles write-authority diagnostic

Last verified2026-10-12 UTC, unchanged `deploy/validation-detail-glyph-01`.
**Original diagnostic only; the approved correction is tracked below. No new ControlId.**

The earlier D8-A read correction explicitly left write authorization unresolved.
Before recording Users & Roles actions, this diagnostic exercises the real
Create/Update, Deactivate and Delete handlers after a live context change.
It uses Admin-generated disposable authority, opens the form while the original
warehouse has a signed-in Admin, retains that selected Auth file, then changes
the live session. Delete confirmation uses the existing fixed test prompt seam;
ordinary modal/Ribbon launch and native confirmation are not established here.

## Observed result

All12 commands persist their requested change. Core's existing capability query
denies ADMIN_MAINT for the original warehouse in the nine non-baseline cases:

| Context before write | Original-target Admin capability | Save / Deactivate / Delete persisted |
|---|---|---|
| Original signed-in Admin | Allowed | All three |
| Signed out | Denied | All three |
| Signed-in Reader | Denied | All three |
| Signed-in Admin of another warehouse | Denied | All three against the original Auth file |

Source path: `frmCreateDeleteUser` handlers -> `SaveUserAndRolesForm`,
`SetUserStatusForm`, `DeleteUserForm` -> `OpenSelectedAuthWorkbookForm(..., True)`
-> direct table mutation and `wb.Save`, with no live Core authorization boundary.
Core recognizing the lost authority does not prevent these form-owned writes.

**139/139 diagnostic/fixture checks pass**, retaining the original111 in order;
this is not139 successful authorization checks. The28 additions establish case
preconditions, the authorized baseline, other-warehouse preservation and fixture
restoration. The12 observations separately record whether writes persisted.
No new behavior is called RED/GREEN before the write contract is resolved.

Five instrumented compiles pass. Excel closes normally; settings, package/tooling
pins and generated Auth bytes are restored. The other warehouse's Auth remains
unchanged in every case. Delayed Application audit through16:31:46.5417308UTC:
zero failures. All496 PowerShell files parse; product source/packages are unchanged,
so the preceding glyph static baseline remains applicable at its recorded scope.

- Controller: `reports/runtime/auth-read-controller/ed969a94324644e095de4da7072147c7`.
- Worker: `reports/runtime/slice4be-auth-read/1cf426d8c184480a91e6028236bc6264`.
- Qualification: controller `qualified-verification.json`.
- Observations: worker `admin-auth-write-diagnostic.json`; fixed case/action names
  and Boolean fields only. No credentials, hashes, account payloads or paths.
- Command: `Test-Slice4beAuthReadOnly.ps1 -DeployRoot deploy/validation-detail-glyph-01
  -Phase GREEN -AdminRefresh -AdminWriteDiagnostic`.

## Approved decision and next test

**D8-USERS-01**, approved by the user on2026-10-12 and recorded in v4.11,
Plan022 and the controls catalog, moves the
three writes into headless Core commands with live captured session/target/Admin
checks, exact trusted Auth selection and protection of unrelated edits. It explicitly
replaces the form-owned write/schema-ensure path; ordinary reads, processor
authorization and explicit provisioning remain separate. Implementation is under test.

The first authority gate on unchanged seed-recording-02 gives **139 PASS/9 expected
FAIL**, retaining all139 prior checks. Each unauthorized handler must leave both
Auth bytes and requested account state unchanged; all nine still persist. Five
instrumented compiles, restoration and normal closure pass; delayed native audit
through18:17:23.7118998UTC is clean. Controller
`reports/runtime/auth-read-controller/37a36cb08c654390861aaae3c2beb13b`, worker
`reports/runtime/slice4be-auth-read/3fe92208742846e59ffbb0504cc3fd73`;
controller `authority-red-verification.json` qualifies the RED. The Phase flag
remains GREEN for the retained read contract, not for the new denied-write assertions.
Command adds `-RequireWriteAuthority` to the earlier wrapper switches.

Protect the new direct Core boundary and authorized persistence,
post-confirmation context loss, missing/invalid/read-only/dirty authority, unknown
columns and existing read/provisioning behavior. Recording/guide coverage follows
the qualified command contract. Full4be-A and the other two proposals remain open.

## Implementation checkpoint

Unpromoted `deploy/validation-admin-users-authority-01`: actual-handler
**GREEN148/148**, retaining all139 prior checks in order. Nine unauthorized
commands preserve Auth bytes and requested account state; all three authorized
commands persist. Five instrumented compiles, fixture/settings/package/tooling
preservation and normal closure pass; delayed native audit through18:27:33.0715296UTC
is clean. Controller `auth-read-controller/e83d09ec669e44e49bb6759d88cd561f`, worker
`slice4be-auth-read/c40e2be8409b49ac906e8fd7f8af803b`, both under `reports/runtime/`.
Controller `authority-green-verification.json` qualifies this scope only.

Build `reports/runtime/admin-users-authority-build-01/verification.json` confirms
five cold compiles/Operations cold start, preservation and clean delayed audit.
New Core `modAuthUserCommands` owns the three commands; Admin delegates through
typed primitive APIs.331 compiled components are unchanged. Seed observer differs
only by the two previously documented trailing newlines; executable text/literals
match. This first checkpoint alone does not establish full D8-USERS-01 or4be-A acceptance.

Final static02:341 components/6427 procedures/139981 lines;28 existing caps do not
grow, dynamic counts remain9/45, and all three schemas/503 PowerShell parses pass.
**Bounded duplicate
exception:**195 groups versus194 reflects only DeactivateUser/DeleteUser's distinct
one-line command adapters (`duplicate:b510b269f2d40420:DeactivateUser+DeleteUser`).
The scanner erases their differing operation literals. Retain these narrow typed
entry points; authority and mutation logic are implemented once in Execute. This
review does not waive future duplication, module-size or dynamic-call limits.

Expanded-test instrumentation initially failed before behavioral execution;
these compile/anchor failures are harness failures, not product RED. Native owned
compile dialogs were dismissed; controllers retained restoration until Excel
closed. Product source and the first148-check qualification were unchanged.

Expanded Core GREEN250 retains148; final visible/data **GREEN259** retains250.
These exercise all three direct Core commands after sign-out, Reader/warehouse/
session change, wrong Auth selection, saved grant removal despite a stale allowed
cache, dirty/read-only borrowed authority, missing files and invalid schema.
Delete's actual confirmation seam signs out before the command. Real BeforeSave
cancellation returns failure for all three commands, leaves the borrowed workbook
dirty and preserves disk bytes; no rollback is claimed. Authorized saves preserve
unknown columns, shifted managed columns, dates and an unchanged credential;
fresh disk inspection verifies a new user's exact capability scope. Fixtures restore.

- Core250: controller `auth-read-controller/e9c4b70530c1470e8aef2252a41ef323`,
  worker `slice4be-auth-read/4664e0f86f094a17a273a49db5113047`.
- Visible259: controller `auth-read-controller/d83395dd10694cb294d92b31c2ad08c9`,
  worker `slice4be-auth-read/010a2b2a9b7b43a8bd9271943d792240`.
- Both are under `reports/runtime/`, with controller `qualified-verification.json`.
  Five instrumented compiles, preservation and normal closure pass; final delayed
  native audit through18:54:35.3119085UTC is clean.
- Final command adds `-CoreAuthority -CaptureEvidence` to the148-check command.
  Four captures are reviewed in worker `visible-review.json`: healthy/invalid Auth,
  the complete session-change refusal and retained Settings Save. Credential fields
  are blank. Modeless display is a test seam; ordinary modal launch, existing path
  and account-copy-button clipping, and comprehensive form acceptance remain open.
- Smoke86: `admin-users-authority-regression/smoke-dcc7222b620b406d883ee18c66a493ff`
  retains the prior86 checks, both unassisted exits, preservation and clean audit.

Full chain32/live-role48/Create Warehouse15 retain all prior ordered checks:
`reports/runtime/admin-users-authority-regression/chain-1b9bb6b512f54beab17ffd847af178d1`.
Its `verification.json` confirms preservation, normal closure and zero native
failures through19:01:29.4473393UTC. All inter-stage waits complete unassisted;
no forced termination, recovery attachment, source-loading registry change or
desktop rebinding was used. The command correction is qualified. Users action
recording, comprehensive form acceptance and4be-A remain open.
