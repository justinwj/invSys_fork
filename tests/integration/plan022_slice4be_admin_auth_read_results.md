# Slice 4be-A: Users & Roles Auth read separation

Last verified: 2026-10-12 UTC. Governing contract: approved v4.11 D8-A.
This prerequisite does not establish Admin action recording or close 4be-A.

## Protecting test and RED

`Test-Slice4beAuthReadOnly.ps1 -AdminRefresh` adds the actual packaged Users &
Roles Refresh, user-selection, Save, Deactivate and Delete handlers to the
existing 82-check Auth/Config/Production suite. Fixtures originate in Admin
Generate Warehouse. Unsaved instrumentation observes workbook ownership,
read-only/saved state and schema-maintenance calls; it does not replace the
opener or schema implementation. Delete confirmation is supplied at the prompt
seam, so this is not native confirmation acceptance. No credentials are logged.

Frozen `deploy/validation-admin-entry-02`: **103 PASS / 8 expected FAIL**.
Closed healthy Auth opens writable, becomes dirty and enters schema maintenance;
two invalid schemas are repaired rather than refused; selection also enters
maintenance. Borrowed dirty state/custom values, file bytes, missing-source
refusal with another warehouse open, and explicit writes pass. All original
82 checks remain GREEN. Five instrumented compiles, restored settings, unchanged
packages/tooling, normal Excel closure and delayed native audit pass.

- Controller: `reports/runtime/auth-read-controller/d77e337ef2ff4287a7f746e12ef2ebf8`.
- Worker: `reports/runtime/slice4be-auth-read/a98d77585f3b4d539a62e202e102239a/red.json`.
- UTC: 04:07:40.3697579–04:10:07.0111886; audit through 04:10:19.0111886.
- Earlier narrower RED: 97 PASS / 7 expected FAIL, controller
  `auth-read-controller/79f1bc58405248d582b285590c3141b0` under runtime.

## Implementation and acceptance scope

Only `frmCreateDeleteUser` changes runtime behavior. Refresh and selection open
closed Auth read-only, validate required schema without repair and preserve
borrowed workbooks. Invalid schema displays “Auth workbook has missing required
tables or columns.” Save, Deactivate and Delete explicitly retain the write path.
Failure cleanup closes only the workbook opened by this operation.

Candidate `deploy/validation-admin-auth-read-01` passes five cold compiles and
Operations cold start; only `invSys.Admin.xlam|frmCreateDeleteUser` changes, with
318 other compiled components unchanged. The previous candidate is preserved.
Build receipts: `reports/runtime/admin-auth-read-build-01`.

Focused GREEN: **111/111**, retaining the exact RED order and all 82 prior GREEN
identities. Five instrumented compiles, settings/package/tooling preservation,
normal shutdown and zero delayed Application failures pass.

- Controller: `reports/runtime/auth-read-controller/e2f2c4f03eb947d2887621e2281882bc`.
- Worker: `reports/runtime/slice4be-auth-read/7638fb8409f44da18fdd89cdc5ed617a/green.json`.
- UTC: 04:11:50.1051964–04:14:26.9639466; audit through 04:14:38.9639466.

Visible GREEN with `-CaptureEvidence`: **111/111** in identical order, five
compiles and normal preserved shutdown; no delayed native failures.

- Controller: `reports/runtime/auth-read-controller/4a2f92d05a3a46a1af0990f429939719`.
- Worker: `reports/runtime/slice4be-auth-read/635eea5ab3d344aa846618f49e8eb2fd/green.json`.
- UTC: 04:14:57.0631013–04:17:31.1347652; audit through 04:17:43.1347652.

Three local captures are reviewed: healthy Users & Roles, invalid-schema refusal
with empty user list, and the retained Settings Save result. The new refusal is
fully visible. Credential controls are blank; generated fixture values appear
only in ignored captures. Form methods exercise real handlers; modeless display
is a test presentation seam, not proof of the ordinary modal launch. One owned
caption activation verifies cursor arrival before clicking; the other two forms
already have focus. No broader mouse-restoration or desktop-handoff claim.
Existing long-path and account-copy-button clipping remains open layout work.

Maintenance receipts: `reports/runtime/admin-auth-read-static-01` and the final
`reports/runtime/admin-auth-read-static-02` refresh.
All 28 size caps hold; the form shrinks 1395 to 1394 lines. Dynamic calls remain
9 literal/45 unresolved and duplicate candidates remain 192. Three schemas and
472 PowerShell scripts validate. No promotion,
broader Users & Roles authorization change, full-chain acceptance or 4be-A
completion is claimed.
