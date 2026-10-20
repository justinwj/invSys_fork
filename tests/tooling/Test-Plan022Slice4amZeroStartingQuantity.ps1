[CmdletBinding()]
param(
    [string]$RepoRoot = (Split-Path -Parent (Split-Path -Parent $PSScriptRoot))
)

$ErrorActionPreference = "Stop"
$repo = (Resolve-Path -LiteralPath $RepoRoot).Path
$docs = (Resolve-Path -LiteralPath (Join-Path $repo "..\invSys_docs")).Path

function Read-Source([string]$relativePath) {
    Get-Content -LiteralPath (Join-Path $repo $relativePath) -Raw
}

$spec = Get-Content -LiteralPath (Join-Path $docs "0 plan docs\xlam_invSys\invSys-Design-v4.11.md") -Raw
$plan = Get-Content -LiteralPath (Join-Path $docs "expert guidance docs\022 Deployed Operations Launcher and NAS Runtime Stabilization Plan.md") -Raw
$controls = Get-Content -LiteralPath (Join-Path $docs "0 plan docs\xlam_invSys\invSys-Controls-v1.md") -Raw
$form = Read-Source "src\Admin\Forms\frmAddInventoryItem.frm"
$admin = Read-Source "src\Admin\Modules\modAdmin.bas"
$createWire = Read-Source "src\Core\Modules\modInventoryCreateWire.bas"
$worksheet = Read-Source "src\Admin\Modules\modAdminInventoryWorksheet.bas"
$apply = Read-Source "src\InventoryDomain\Modules\modInventoryApply.bas"
$queries = Read-Source "src\InventoryDomain\Modules\modInventoryQueries.bas"
$itemSearch = Read-Source "src\Core\ClassModules\cDynItemSearch.cls"
$viewer = Read-Source "src\Core\Modules\modInventoryViewerData.bas"
$validator = Read-Source "tools\validate_phase6_packaged_xlams.ps1"

$addFunction = [regex]::Match(
    $admin,
    'Public Function AddInventoryItemForWarehouse(?s:.*?)End Function'
).Value
$createPayload = [regex]::Match(
    $admin,
    'Private Function BuildInventoryCreatePayloadAdmin(?s:.*?)End Function'
).Value
$wirePayload = [regex]::Match(
    $createWire,
    'Public Function BuildPayload(?s:.*?)End Function'
).Value
$formContract = [regex]::Match(
    $form,
    'Public Function TestAddItemVisibilityDropdownContract(?s:.*?)End Function'
).Value
$worksheetValidation = [regex]::Match(
    $worksheet,
    'Private Function ValidateInventoryWorksheetRow(?s:.*?)End Function'
).Value
$entityQuery = [regex]::Match(
    $queries,
    'Public Function ListAvailableInventoryEntities(?s:.*?)End Function'
).Value
$pickerQuery = [regex]::Match(
    $queries,
    'Public Function ListInventoryPickerItems(?s:.*?)End Function'
).Value

$checks = [ordered]@{}
$checks["Docs.Slice4amContract"] =
    $spec.Contains("counted Starting Qty is required to be numeric and may be zero or greater") -and
    $plan.Contains("Slice 4am -- Zero starting quantity for managed inventory creation") -and
    $controls.Contains("numeric Starting Qty of zero or greater")
$checks["Form.RealAddAcceptsZeroRejectsNegative"] =
    $formContract.Contains('mTxtQty.Value = "0"') -and
    $formContract.Contains('mTxtQty.Value = "-1"') -and
    $formContract.Contains("mBtnOK_Click") -and
    $formContract.Contains("ZeroQtyAccepted=") -and
    $formContract.Contains("NegativeRejected=") -and
    $form.Contains('If StartingQty < 0 Then')
$checks["Admin.ServiceAllowsZeroRejectsNegative"] =
    $addFunction.Contains('ElseIf qty < 0 Then') -and
    $addFunction.Contains('Starting quantity cannot be negative.') -and
    $addFunction.Contains('payloadJson = BuildInventoryCreatePayloadAdmin(systemKey, sku, itemName, uom, locationVal, qty,') -and
    $createPayload.Contains('modInventoryCreateWire.BuildPayload(') -and
    $createPayload.Contains('systemKey, sku, qty, locationVal, "GOOD",') -and
    $wirePayload.Contains('modRoleEventWriter.CreateInventoryEntityPayloadItem(') -and
    $wirePayload.Contains('systemKey, sku, qty, location, condition, attributesJson, note)') -and
    $wirePayload.Contains('modRoleEventWriter.BuildPayloadJsonFromCollection(items)')
$checks["Worksheet.CountedAddAcceptsExplicitZero"] =
    $worksheet.Contains('"COUNTED", 0, "CLEARVIEW"') -and
    $worksheetValidation.Contains('actionName = "ADD" And (Not hasQty Or qty < 0)') -and
    $worksheetValidation.Contains('COUNTED ADD requires an explicit nonnegative Quantity.') -and
    $worksheet.Contains("ZeroCounted=")
$checks["Domain.CountedZeroCreatesActiveExactEntity"] =
    $apply.Contains("Public Function InventoryZeroCreateContractForAutomation") -and
    $apply.Contains('EVENT_TYPE_INVENTORY_CREATE And qty = 0 Then GoTo QtyAccepted') -and
    $apply.Contains('IIf(qtyOnHand >= 0 Or (entity.Exists("NonCounted")') -and
    $apply.Contains('availableRows = modInventoryQueries.ListInventoryPickerItems("", wb)') -and
    $apply.Contains("ZeroEntityActive=") -and
    $apply.Contains("NegativeRejected=")
$checks["ManagedPickerRetainsZeroWithoutRunAllocation"] =
    $pickerQuery.Contains('qtyOnHand = InventoryQuerySkuBalance(loBalance, sku)') -and
    $pickerQuery.Contains('ResolveInventoryQueryRepresentativeSystemKey(loEntities, sku)') -and
    $pickerQuery.Contains('If catalogState = "RETIRED" Then GoTo NextCatalogRow') -and
    $itemSearch -match 'LoadProcessManagedInventoryItems(?s).*?ListInventoryPickerItemsBridge' -and
    $entityQuery.Contains('If qtyOnHand <= 0 And Not nonCounted Then GoTo NextEntity') -and
    $entityQuery.Contains('StrComp(inventoryState, "ACTIVE"')
$checks["ViewerShowsActiveZeroOmitsRetired"] =
    $viewer.Contains('If inventoryState = "RETIRED" Then GoTo ContinueRow') -and
    $viewer.Contains('If CDbl(quantities(key)) >= 0 Then visibleCount = visibleCount + 1') -and
    $viewer.Contains('If CDbl(quantities(key)) < 0 Then GoTo ContinueDisplayGroup')
$checks["Packaged.ZeroQuantityEvidence"] =
    $validator.Contains("Admin.InventoryZeroStartingQuantity") -and
    $validator.Contains("InventoryDomain.InventoryZeroCreate") -and
    $validator.Contains("ZeroQtyAccepted=True") -and
    $validator.Contains("ZeroEntityActive=True") -and
    $validator.Contains("NegativeRejected=True")

$passed = 0
$red = 0
foreach ($entry in $checks.GetEnumerator()) {
    if ($entry.Value) {
        Write-Host ("PASS " + $entry.Key)
        $passed++
    }
    else {
        Write-Host ("RED  " + $entry.Key)
        $red++
    }
}

Write-Host ("PLAN022_SLICE4AM_SOURCE passed={0} red={1} total={2}" -f $passed, $red, $checks.Count)
if ($red -gt 0) { exit 1 }
