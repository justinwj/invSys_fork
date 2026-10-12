# 4be-A Users & Roles write-authority diagnostic

Last verified2026-10-12 UTC, unchanged `deploy/validation-detail-glyph-01`.
**Diagnostic only; no authorization acceptance, product change or new ControlId.**

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

## Required decision and next test

Proposed **D8-USERS-01** in v4.11(line4991), Plan022 and Controls1.550 moves the
three writes into headless Core commands with live captured session/target/Admin
checks, exact trusted Auth selection and protection of unrelated edits. It explicitly
replaces the form-owned write/schema-ensure path; ordinary reads, processor
authorization and explicit provisioning remain separate. Approval is pending.

After approval, first add/run actual-handler denial assertions for these cases
and direct Core boundary tests. Then implement and protect authorized persistence,
post-confirmation context loss, missing/invalid/read-only/dirty authority, unknown
columns and existing read/provisioning behavior. Recording/guide coverage follows
the qualified command contract. Full4be-A and the other two proposals remain open.
