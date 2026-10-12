# Slice4be remaining Admin coverage

Source/package audit,2026-10-12, code `a0491583`; candidate
`deploy/validation-admin-entry-02`, catalog33/151. **Not runtime acceptance.**
This map discovers existing entries under D18; it approves no new ControlId,
outcome, permission, exclusion, replay adapter or business rule.

The packaged Admin Ribbon contains exactly14 Actions buttons. All use
`RibbonOnActionAdmin` and `RibbonRequiredCapabilityGetEnabledAdmin`; their IDs
match `tools/build-xlam.ps1`. The generator checks `ADMIN_MAINT` before each
macro. Many macros do not repeat that check. Test the actual dispatch/handler
path; direct service calls alone do not prove operator authorization or capture.
Ignored verification: `reports/runtime/admin-remaining-source-audit-01/ribbon-verification.json`.

## Actions map

All rows except Settings remain pending shared recording/publication/guide proof.
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
| `btnAdminShipmentReconcile` | `modAdminShipmentReconcile.OpenShipmentReconcileTool` | Native selection/narrative dialogs -> `QueueAdminShipmentReconcileEvent` -> `modRoleEventWriter.QueuePayloadEvent`. Current identity path conflicts with D14; do not accept or add a ROW-compatible observation contract. |
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

## Confirmed identity conflict: Shipment Reconcile

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

## Auth refresh finding requiring a focused test

`frmCreateDeleteUser.mBtnRefreshUsers_Click` calls `RefreshUsersForm`, which calls
`OpenSelectedAuthWorkbookForm`. An already-open Auth workbook is borrowed, but a
closed one is opened writable and passed to `modAuth.EnsureAuthSchema`. Thus a read
path enters provisioning/schema maintenance. Its exact mutation/closure effects
have not been measured. The same helper is reused by write actions, so replacing
it indiscriminately could alter authorized maintenance.

Reconcile this path with D8-A Auth read/provisioning separation using the actual
Refresh handler on generated complete and incomplete Auth workbooks. Verify no
create/repair/format/dirty/save, preservation of caller-owned open books, missing
authority refusal and current warehouse/rights behavior. Form save/deactivate/delete
also need current-context/capability review; a prior Ribbon gate is not evidence
of a later write gate. This audit neither changes Auth write ownership nor grants
a new permission exception. Any material Auth contract change needs an explicit
architecture decision and approval before implementation.
