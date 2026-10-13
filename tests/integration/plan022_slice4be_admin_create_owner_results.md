# Create New Warehouse owner diagnostic

4be-A / Admin prerequisite,2026-10-13. **D8-CREATE-01 is proposed; not approved.**
Production code and unpromoted `deploy/validation-root-recording-04` are unchanged,
catalog40/162. This is owner characterization and authorization RED, not recording
or4be-A acceptance.

## Finding and protecting test

The actual generated `btnAdminCreateWarehouse` Ribbon callback opens the actual
`frmCreateWarehouse` under an authenticated Admin. The test signs out via Core,
proves `IsSignedIn=False`, and invokes the form's own `btnOK_Click`. A new Config
is saved and the form reports local completion. The two failing assertions are
`Create.SignedOut.DoesNotCreateRuntime` and `Create.SignedOut.NoCompletionClaim`.
Ribbon authorization alone does not protect the later form write. Admin calls
the separately available headless bootstrap without a captured live-session gate.

The normative station-setup paragraph requires signed-in ADMIN_MAINT; D8-A keeps
explicit provisioning separate and D8-USERS-01 concerns Users only. Proposed
D8-CREATE-01 defines the interactive captured-context boundary for approval.
No production authorization or bootstrap behavior has been changed.

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File tests/tooling/Test-Slice4beAdminCreateWarehouseOwner.ps1 -DeployRoot deploy/validation-root-recording-04 -Phase RED
```

Test adapters are in-memory only: an Office.IRibbonControl selects the real
callback, its existing Show becomes modeless to allow subsequent handler calls,
and fixed facts/counters expose outcomes. Bootstrap/publication results are real.
The partial-publication case uses a regular file inside the generated destination
to block its warehouse folder; retry uses a fresh generated destination. No ACL
changes, operational workbooks or fabricated owner returns are involved. Credentials
stay in the shared harness's disposable fixtures and are never emitted.

## Scope and evidence

Initial valid worker `slice4be-admin-create-owner/c6c219f5afad4fc0b39e11c1630d72cf`
finishes **45 PASS / 2 expected FAIL**, including all five instrumented compiles.
Passed behavior: Operational and Training local creation with saved purpose;
Config/Auth/Inventory/inbox artifacts; opening the exact generated Receiving
workbook; duplicate refusal with Config bytes preserved; invalid input and Cancel
without provisioning; complete publication; retained local success after failed
publication; retry without rebootstrap or changing local Config bytes; unchanged
source/other Config bytes. The signed-out precondition itself passes.

Closure-qualified controller `admin-create-owner-controller/bc3531d7e5b04df3af169524d0c2b3f1`,
worker `slice4be-admin-create-owner/226ea78b479541ec9fed157b8f988dfc`, repeats **45/2**.
All47 log check IDs/results exactly match JSON. UTC18:11:46.444--18:13:46.253;
settings, five package pins and tooling pins are preserved; Excel closes normally.
Delayed native audit through18:14:06.275UTC finds zero Excel error/hang events and
no Excel process. No forced termination or recovery attachment was used.
Ignored evidence roots are under `reports/runtime/`; no runtime payload is committed.

All522 PowerShell files parse. The desktop monitor runs18:06:00.552--18:14:37.806UTC,
103 samples with zero cursor/desktop/capture/error5 failures, and stops normally.
Separate `navigation-cursor/61123cc91dd8430080b6a60216ca5842` verifies actual
SetCursorPos movement to the exact target and original-position restoration.
This is input availability evidence, not native Create form acceptance.

Regenerated `create-owner-maintenance-01` retains345 components,6437 procedures,
140271 lines, dynamic calls9/45 and duplicate candidates196. Production source
has no diff, so all28 existing oversized-module limits remain unchanged. The
reviewed-backlog schema passes; no static exception or code deletion is introduced.

Excluded harness attempts: an unsupported visible-Excel switch stopped before
Excel launch; worker `10b882866f3041fab0d08a31eb8d0b07` stopped before compile because
the counter adapter expected an assignment instead of the actual If-call boundary.
The adapter was corrected without product edits. Neither attempt is product RED.

Not yet covered: recording/catalog/policy/guide conclusions, native Ribbon clicks,
ordinary modal interaction, folder-helper dialogs, context/right-change cases,
post-sign-out publication retry or direct interactive Core commands. Existing
full-chain/live-role/static/layout/visible results remain scoped prior evidence;
they are not fresh acceptance of this diagnostic or a new implementation.
