# Slice4be remaining Admin coverage

Source/package audit,2026-10-12, code `a0491583`; candidate
`deploy/validation-admin-entry-02`, catalog33/151. **Not runtime acceptance.**
This map discovers existing entries under D18; it approves no new ControlId,
outcome, permission, exclusion, replay adapter or business rule.

Continuation: [Shipment Reconcile evidence](plan022_slice4be_admin_reconcile_results.md)
now passes330/330, retaining the258-check recording prefix under approved D8-RECON-01.
Scrollable history, native entry20/Inspect/Escape, resize, window-close and
confirmation Cancel/typed Submit pass. Catalog34 Open/Inspect/Cancel recording,
suppression, exact references and original-session binding pass. Actual storage
failure preserves one command and one warning without invented completion.
Independent native recording, exact publication, guide authoring and saved
Concluded/Incomplete CommandCompleted views pass. SourceEventsApplied guide
proof and local-target recording eligibility remain open.
Earlier31/24 RED covered repeated/multi-entity history,
actual queued application and Domain association/quantity guards.
The user approved D14-RECON-01 in v4.11 on2026-10-12. The
source audit below remains useful; its original no-behavioral-RED status is superseded
for this callback only. The subsequent [Auth read correction](plan022_slice4be_admin_auth_read_results.md)
passes 111 packaged checks under D8-A with three reviewed captures. Full layout,
write authorization and remaining tracking coverage are still open.

The packaged Admin Ribbon contains exactly14 Actions buttons. All use
`RibbonOnActionAdmin` and `RibbonRequiredCapabilityGetEnabledAdmin`; their IDs
match `tools/build-xlam.ps1`. The generator checks `ADMIN_MAINT` before each
macro. Many macros do not repeat that check. Test the actual dispatch/handler
path; direct service calls alone do not prove operator authorization or capture.
Ignored verification: `reports/runtime/admin-remaining-source-audit-01/ribbon-verification.json`.

## Actions map

Except for Settings, complete recording/publication/guide proof remains pending;
Reconcile has the scoped command-conclusion gate described above.
Form subactions below are source findings, not a complete live-control census.
Do not count typing, initialization, programmatic selection or internal refresh
as deliberate actions merely because an event procedure exists.

| Packaged button | Public macro | Existing owner/surface and result to protect |
|---|---|---|
| `btnAdminOpen` | `modAdmin.Admin_Click` | `modAdminConsole.OpenAdminConsole`: schema/local projection preparation, refresh, then AdminConsole sheet activation. Its Boolean can be False; the macro currently discards it. Presentation is not processing or Domain application. |
| `btnAdminUsers` | `modAdmin.Open_CreateDeleteUser` | `frmCreateDeleteUser`: server/root discovery, warehouse/user selection, refresh, PIN controls, clear, save/deactivate/delete, Close. Mutation and read boundaries require separate review; see Auth finding below. Never record credentials or user-field values. |
| `btnAdminSettings` | `modAdmin.Open_Settings` | Existing catalog33 Open/five tabs plus catalog32 Close; launcher98, native177 and continuous recording/guide125 have scoped evidence. General/editor/policy controls retain their separate records. |
| `btnAdminWarehouses` | `modAdmin.Open_WarehouseDirectory` | Optional root prompt, `modAdminConsole.OpenWarehouseDirectory`, directory refresh/activation and existing success/failure notice. Do not log discovered paths or assume directory viewing changes the trusted target. |
| `btnAdminWarehouseRoot` | `modAdmin.Add_WarehouseDirectoryRoot` | Explicit root prompt and `RememberWarehouseScanRoot`; local remembered-root change, then directory presentation. Cancellation, remembering and refresh are distinct facts; no new warehouse is created. |
| `btnAdminCreateWarehouse` | `modAdmin.Open_CreateWarehouse` | `frmCreateWarehouse`: folder helpers, Operational/Training choice, validation/bootstrap, optional/retried initial publication, Open Receiving and Close. Local bootstrap can succeed before publication fails. No existing trusted runtime means no warehouse activity store. |
| `btnAdminSetupTesterStation` | `modAdmin.Admin_SetupTesterStation_Click` | `frmSetupTesterStation`: path helpers, `modTesterSetup.SetupTesterStation`, open generated Receiving workbook, delete generated setup, cancel. Capture only eligible current-context actions; never use existing operational workbooks as disposable fixtures. |
| `btnAdminAddInventoryItem` | `modAdmin.Add_InventoryItem` | Captures exact warehouse/station/user and operator workbook, then modal `frmAddInventoryItem`: Add/Edit selection, custom fields, sheet workbench create/upload, accept/delete/cancel. Owner command results occur after form acceptance; retain existing System_Key, unknown columns and workflow gates. |
| `btnAdminSeedInventory` | `modAdmin.Seed_DemoInventory` | `frmSeedInventory` selects Seed/Delete/Upload Data Set/Delete Data Set; callback confirmations and `modAdminInventorySeed` results follow. Inbox repair has its own form handler. CSV definition changes are not inventory changes. Existing automation result text is not shared activity proof. |
| `btnAdminShipmentReconcile` | `modAdminShipmentReconcile.OpenShipmentReconcileTool` | `frmShipmentReconcileSelection`/narrative -> `QueueAdminShipmentReconcileEvent` -> Core live-context gate -> provisioned queue API. D14-RECON-01/D8-RECON-01/D18 pass330 checks: native handlers, catalog34 recording, suppression, references, binding, storage failure and independent recording/publication/CommandCompleted guide proof. SourceEventsApplied guide proof and local recording eligibility remain open. |
| `btnAdminHqAggregator` | `modAdmin.Admin_AggregateGlobalSnapshot_Click` | `frmAggregationSources`: discover/add/connect, source selection, `RunHQAggregationFromSourceSet`, Close. Explicit source-set aggregation remains advisory. Connection credentials and source paths are excluded from activity payloads. |
| `btnAdminDesignLifecycle` | `modAdminDesignLifecycle.Admin_DesignLifecycle_Click` | `frmAdminDesignLifecycle`: refresh, constructed Import Legacy Recipes, Release selected, Obsolete selected, Close. `ExecuteDesignLifecycleCommand` separates queued event, processing and resulting status. A form open is not a release; recipe import is not authority to migrate legacy inventory. |
| `btnAdminVerifyAddinsPublished` | `modAdmin.Verify_AddinsPublished` | `modAddinsPublish.VerifyAddinsPublished` checks required files/nonzero sizes and shows existing result. Verification is not PublishAddins. Capture fixed result facts, not full diagnostic paths. |
| `btnAdminRetireMigrateWarehouse` | `modAdmin.Admin_RetireMigrateWarehouse_Click` | `frmRetireMigrateWarehouse`: source/target/root selection, connection, archive destination, mode, re-auth, confirmation, Back/Cancel/Close. `HandleConfirmOk` performs separate archive/migrate/retire/delete stages; partial success must not become rollback or whole-operation success. |

## Shared session entries and old console functions

The same packaged Session group also declares `btnAdminServerSession`,
`btnAdminCurrentUser`, and `ddAdminWarehouseTarget`. These are outside the14
Actions count. Their Core session/target handlers need shared Operations/Admin
accounting. D18 excludes pre-sign-in actions without a trusted warehouse actor;
it does not blanket-exclude all signed-in session actions. Target changes end the
original recording; later target/user identities cannot be assigned retroactively.

`modAdminConsole` still exposes RunProcessorFromConsole, BreakInventoryLock,
ReissuePoisonEvent, GenerateInventorySnapshot and PublishWarehouseArtifacts.
The current Actions Ribbon does not directly declare them, and
`ClearAdminConsoleActionButtons` deletes old `btnAdmin` shapes. Tests, processor
internals and scheduled automation still call some owners. Complete reachability
review before classifying a deliberate entry, compatibility-only API or exclusion;
neither public visibility nor a unit-test call proves a user-side control.

## Shipment Reconcile prerequisite: scoped correction

The historical source finding below is superseded by approved D14-RECON-01 and
candidate `validation-admin-reconcile-01`: focused31/24 RED →55/55 GREEN, including
exact history/queue/Domain application and unknown-column preservation. See
[Reconcile evidence](plan022_slice4be_admin_reconcile_results.md). Recording,
native dialogs, role and authority edge gates remain open.

The subsequent130-check authority gate retains those55, passes closed/missing/
invalid/borrowed-edit/cancellation and reader-denial cases, but queues from a
signed-out stale provisioned user-ID hint:128 pass/2 fail. D8-RECON-01 proposes
the interactive Core authentication boundary; approved2026-10-12 in v4.11.
Do not silently change headless processor authorization or Admin local-target policy.

Current `validation-admin-reconcile-access-02` fixes that authentication path:
172/172 pass, including live sign-out/re-sign-in/user/target changes and permitted
local fallback. Combined native179/2 leaves only history truncation/clipping RED;
this is not complete Reconcile recording or4be-A acceptance.

### Historical identity conflict

Normative v4.11 D14 requires opaque exact `System_Key`, prohibits managed `ROW`
and prohibits identifying one entity by SKU. Current source instead:

1. `BuildRecentShipmentSentRows` selects EventID/SKU without System_Key.
2. `DetectNasIncreaseAfterShipEvent` uses `tblSkuBalance` and SKU history.
3. `ResolveRowForSku` reads `tblSkuCatalog.ROW` and chooses by SKU/ITEM_CODE.
4. `QueueAdminShipmentReconcileEvent` calls `CreatePayloadItem`, which emits `Row`
   and SKU, rather than `CreateInventoryEntityPayloadItem` with System_Key.

The packaged Ribbon exposes this macro. This is a confirmed source conflict,
**not a recorded behavioral RED**. No runtime correction has been attempted.
Do not preserve ROW as a compatibility field, derive a key from SKU/position,
silently remove the reachable button, or count its current behavior as acceptance.

Next test-first prerequisite: use an Admin-generated training warehouse and
owner-created inventory/shipment history, invoke the actual packaged correction
callback, and verify that its queued correction identifies only the selected
shipment entity by exact System_Key. Include another entity sharing the SKU,
preserve both keys/custom fields and saved unrelated authority, and distinguish
callback refusal from missing fixtures or a broken dialog driver. A direct queue
service test may supplement this, not replace the actual handler. Resolve any
underspecified correction semantics in v4.11 before changing them.

## Auth refresh prerequisite and separate write review

`frmCreateDeleteUser.mBtnRefreshUsers_Click` calls `RefreshUsersForm`, which calls
`OpenSelectedAuthWorkbookForm`. An already-open Auth workbook is borrowed, but a
closed one in candidate02 is opened writable and passed to `modAuth.EnsureAuthSchema`.
Packaged RED subsequently proves dirtying and invalid-schema repair; candidate
`validation-admin-auth-read-01` separates ordinary reads from explicit writes and
passes 111 checks. Exact evidence and remaining visible gates are in the
[Auth read result](plan022_slice4be_admin_auth_read_results.md). The new refusal
message is readable; existing long-path/account-copy clipping remains open.

The actual Refresh and selection handlers now protect complete/incomplete/missing
Auth, borrowed edits and other-warehouse refusal. Save/deactivate/delete retain
their explicit write behavior. These writes still need current-context/capability
review; a prior Ribbon gate is not evidence
of a later write gate. This audit neither changes Auth write ownership nor grants
a new permission exception. Any material Auth contract change needs an explicit
architecture decision and approval before implementation.
