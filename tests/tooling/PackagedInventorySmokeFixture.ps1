# Disposable Admin-generated fixture and unsaved test-only form adapters.
# No authorization handler is replaced. Credentials remain in memory except for
# the generated fixture's Auth store, exactly as in the existing fixture factory.
function Get-PackagedFixtureHash {
    param([string]$Path)
    $stream=[IO.File]::Open($Path,[IO.FileMode]::Open,[IO.FileAccess]::Read,([IO.FileShare]::ReadWrite -bor [IO.FileShare]::Delete))
    try {
        $sha=[Security.Cryptography.SHA256]::Create()
        try {[BitConverter]::ToString($sha.ComputeHash($stream)).Replace('-','')}
        finally {$sha.Dispose()}
    } finally {$stream.Dispose()}
}

function Invoke-PackagedFixtureMacro {
    param($Excel,[string]$Macro,[object[]]$Arguments=@())
    Assert-PackagedDesktop
    # Explicit COM arities preserve primitive Variant types through Excel.Run.
    switch($Arguments.Count){
        0 {$Excel.Run($Macro)}
        1 {$Excel.Run($Macro,$Arguments[0])}
        2 {$Excel.Run($Macro,$Arguments[0],$Arguments[1])}
        3 {$Excel.Run($Macro,$Arguments[0],$Arguments[1],$Arguments[2])}
        4 {$Excel.Run($Macro,$Arguments[0],$Arguments[1],$Arguments[2],$Arguments[3])}
        6 {$Excel.Run($Macro,$Arguments[0],$Arguments[1],$Arguments[2],$Arguments[3],$Arguments[4],$Arguments[5])}
        default {throw 'Unsupported fixture macro arity.'}
    }
}

function Initialize-PackagedInventorySmokeFixture {
    param($Excel,$AdminPackage,[string]$Root,[bool]$BindContext)
    $fixtureRoot=Join-Path $Root 'admin-inventory-fixture'
    $share=Join-Path $fixtureRoot 'published'
    New-Item -ItemType Directory -Path $share -Force|Out-Null
    [void](Invoke-PackagedFixtureMacro $Excel "'invSys.Core.xlam'!modRuntimeWorkbooks.SetCoreDataRootOverride" @($fixtureRoot))
    $created=[bool](Invoke-PackagedFixtureMacro $Excel "'invSys.Admin.xlam'!modAdminConsole.BootstrapWarehouseLocalAdmin" @('WH-TEST','Packaged Inventory fixture','S1','admin',$fixtureRoot,$share))
    if(-not $created){throw 'Admin-generated Inventory fixture failed.'}
    $secret=[guid]::NewGuid().ToString('N')
    [double]$acc=5381
    for($i=0;$i -lt $secret.Length;$i++){$acc=($acc*33+[int][char]$secret[$i]+$i+1)%2147483647}
    $credential='{0:X8}' -f [int]$acc
    Assert-PackagedDesktop
    $auth=$Excel.Workbooks.Open((Join-Path $fixtureRoot 'WH-TEST.invSys.Auth.xlsb'),0,$false)
    try {
        $users=$null
        foreach($sheet in $auth.Worksheets){foreach($table in $sheet.ListObjects){if($table.Name -ceq 'tblUsers'){$users=$table}}}
        if($null -eq $users){throw 'Generated Auth users table unavailable.'}
        $users.ListColumns.Item('PinHash').Range.NumberFormat='@'
        $matched=0
        foreach($row in $users.ListRows){
            if([string]$row.Range.Cells.Item(1,$users.ListColumns.Item('UserId').Index).Value2 -ceq 'admin'){
                $row.Range.Cells.Item(1,$users.ListColumns.Item('PinHash').Index).Value2=$credential
                if([string]$row.Range.Cells.Item(1,$users.ListColumns.Item('PinHash').Index).Value2 -cne $credential){throw 'Fixture credential did not round-trip.'}
                $matched++
            }
        }
        if($matched -ne 1){throw 'Generated fixture Admin is ambiguous.'}
        $auth.Save()
    } finally {$auth.Close($false)}
    $module=$AdminPackage.VBProject.VBComponents.Add(1); $module.Name='TestPackagedInventoryFixture'
    $module.CodeModule.AddFromString(@'
Option Explicit
Private mBind As Boolean, mContext As String
Public Sub SetBinding(ByVal enabled As Boolean)
    mBind = enabled
End Sub
Public Function Seed(ByVal csvPath As String) As Boolean
    Dim report As String
    Seed = modAdminInventorySeed.UploadDemoInventoryForWarehouse("WH-TEST", "S1", "admin", csvPath, report)
End Function
Public Function ContextReady(ByVal hostName As String) As Boolean
    Dim context As String
    Application.Workbooks(hostName).Activate
    context = modInventoryCommands.Capture(hostName)
    ContextReady = (context <> "")
    modInventoryCommands.Release context
End Function
Private Sub Prepare(ByVal hostName As String)
    Dim book As Workbook
    Set book = Application.Workbooks(hostName)
    book.Activate
    mContext = ""
    If mBind Then mContext = modInventoryCommands.Capture(hostName)
    frmAddInventoryItem.SetOperatorWorkbook book
    frmAddInventoryItem.SetInteractiveContext mContext
End Sub
Public Function Add(ByVal hostName As String) As String
    Prepare hostName
    Add = modAdmin.InventoryAddVisibilityDropdownContractForAutomation()
    modInventoryCommands.Release mContext
End Function
Public Function Delete(ByVal hostName As String) As String
    Prepare hostName
    Delete = modAdmin.InventoryDeleteContractForAutomation()
    modInventoryCommands.Release mContext
End Function
Public Function Worksheet(ByVal hostName As String) As String
    Prepare hostName
    Worksheet = modAdmin.InventoryWorksheetContractForAutomation(Application.Workbooks(hostName))
    modInventoryCommands.Release mContext
End Function
Public Sub ReopenForUpload(ByVal form As frmAddInventoryItem, ByVal book As Workbook)
    If Not mBind Then Exit Sub
    modInventoryCommands.Release mContext
    mContext = modInventoryCommands.Capture(book.Name)
    form.SetInteractiveContext mContext
End Sub
'@)
    # These edits are confined to packaged TEST helpers. The owning handlers,
    # Core guards and command implementations remain byte-for-byte unchanged.
    $form=$AdminPackage.VBProject.VBComponents.Item('frmAddInventoryItem').CodeModule
    $line=$form.ProcStartLine('TestInventoryWorksheetActionContract',0); $count=$form.ProcCountLines('TestInventoryWorksheetActionContract',0)
    $body=$form.Lines($line,$count)
    $boundary='    testStage = "UploadHandler"'
    if([regex]::Matches($body,[regex]::Escape($boundary)).Count -ne 1){throw 'Worksheet test-helper boundary changed.'}
    $body=$body.Replace('"EXISTING-SKU"','"DEMO-EXISTING"').Replace($boundary,('    TestPackagedInventoryFixture.ReopenForUpload Me, operatorWb'+"`r`n"+$boundary))
    $form.DeleteLines($line,$count); $form.InsertLines($line,$body)
    $worksheet=$AdminPackage.VBProject.VBComponents.Item('modAdminInventoryWorksheet').CodeModule
    foreach($name in @('PopulateInventoryWorksheetContractRowsForTest','InventoryWorksheetEvidenceForTest')){
        $line=$worksheet.ProcStartLine($name,0); $count=$worksheet.ProcCountLines($name,0)
        $body=$worksheet.Lines($line,$count)
        if(-not $body.Contains('"EXISTING-SKU"')){throw 'Worksheet test fixture identity boundary changed.'}
        $worksheet.DeleteLines($line,$count); $worksheet.InsertLines($line,$body.Replace('"EXISTING-SKU"','"DEMO-EXISTING"'))
    }
    $csv=Join-Path $fixtureRoot 'seed.csv'
    [IO.File]::WriteAllText($csv,"ITEM_CODE,ITEM,QTY,UOM,LOCATION,Condition`r`nDEMO-EXISTING,Existing Item,25,EA,CLEARVIEW,GOOD`r`n")
    # Establish runtime binding after installing all unsaved VBA adapters.
    [void](Invoke-PackagedFixtureMacro $Excel "'invSys.Core.xlam'!modRuntimeWorkbooks.SetCoreDataRootOverride" @($fixtureRoot))
    $selected=[string](Invoke-PackagedFixtureMacro $Excel "'invSys.Core.xlam'!modNasConnection.SelectWarehouseTargetForAutomation" @($fixtureRoot,$fixtureRoot,'S1',$false))
    if(-not $selected.StartsWith('OK|')){throw 'Fixture target selection failed.'}
    $seeded=[bool](Invoke-PackagedFixtureMacro $Excel "'invSys.Admin.xlam'!TestPackagedInventoryFixture.Seed" @($csv))
    if(-not $seeded){throw 'Admin demo fixture seed failed.'}
    $inventoryPath=Join-Path $fixtureRoot 'WH-TEST.invSys.Data.Inventory.xlsb'
    $openInventory=@($Excel.Workbooks|Where-Object {[string]$_.FullName -ceq $inventoryPath})
    [pscustomobject]@{OpenCount=$openInventory.Count;Saved=if($openInventory.Count -eq 1){[bool]$openInventory[0].Saved}else{$null};ReadOnly=if($openInventory.Count -eq 1){[bool]$openInventory[0].ReadOnly}else{$null}}|ConvertTo-Json|Set-Content (Join-Path $Root 'inventory-fixture-open-state.json')
    $beforeRead=Get-PackagedFixtureHash $inventoryPath
    foreach($owned in $openInventory){$owned.Close($false)}
    Assert-PackagedDesktop
    $persisted=$Excel.Workbooks.Open($inventoryPath,0,$true)
    $entityReady=$false; $catalogReady=$false; $appliedReady=$false; $logReady=$false
    try {
        $tables=@{}
        foreach($sheet in $persisted.Worksheets){foreach($table in $sheet.ListObjects){$tables[[string]$table.Name]=$table}}
        $entities=$tables['tblInventoryEntities']; $catalog=$tables['tblSkuCatalog']
        $applied=$tables['tblAppliedEvents']; $log=$tables['tblInventoryLog']
        $keys=@(); $events=@()
        foreach($row in $entities.ListRows){
            if([string]$row.Range.Cells.Item(1,$entities.ListColumns.Item('SKU').Index).Value2 -ceq 'DEMO-EXISTING'){
                $key=[string]$row.Range.Cells.Item(1,$entities.ListColumns.Item('System_Key').Index).Value2
                if($key -ne '' -and [double]$row.Range.Cells.Item(1,$entities.ListColumns.Item('QtyOnHand').Index).Value2 -eq 25){$keys+=,$key}
            }
        }
        $entityReady=$keys.Count -eq 1
        foreach($row in $catalog.ListRows){if([string]$row.Range.Cells.Item(1,$catalog.ListColumns.Item('SKU').Index).Value2 -ceq 'DEMO-EXISTING'){$catalogReady=$true}}
        foreach($row in $log.ListRows){
            if($entityReady -and [string]$row.Range.Cells.Item(1,$log.ListColumns.Item('System_Key').Index).Value2 -ceq $keys[0] -and [double]$row.Range.Cells.Item(1,$log.ListColumns.Item('QtyDelta').Index).Value2 -eq 25){
                $events+=,[string]$row.Range.Cells.Item(1,$log.ListColumns.Item('EventID').Index).Value2
            }
        }
        $logReady=$events.Count -eq 1
        foreach($row in $applied.ListRows){
            if($logReady -and [string]$row.Range.Cells.Item(1,$applied.ListColumns.Item('EventID').Index).Value2 -ceq $events[0] -and [string]$row.Range.Cells.Item(1,$applied.ListColumns.Item('Status').Index).Value2 -ceq 'APPLIED'){$appliedReady=$true}
        }
    } finally {$persisted.Close($false)}
    $unchanged=(Get-PackagedFixtureHash $inventoryPath) -ceq $beforeRead
    [pscustomobject]@{EntityReady=$entityReady;CatalogReady=$catalogReady;LogReady=$logReady;AppliedReady=$appliedReady;ReadOnlyBytesUnchanged=$unchanged}|ConvertTo-Json|Set-Content (Join-Path $Root 'inventory-fixture-persistence.json')
    if(-not($entityReady -and $catalogReady -and $logReady -and $appliedReady -and $unchanged)){throw 'Seed persistence was not independently confirmed; not meaningful RED.'}
    Assert-PackagedDesktop
    $hostBook=$Excel.Workbooks.Add()
    $hostBook.SaveAs((Join-Path $fixtureRoot 'Admin.Inventory.Smoke.xlsx'),51)
    $signed=[string](Invoke-PackagedFixtureMacro $Excel "'invSys.Core.xlam'!modAuth.SignInCurrentTargetForAutomation" @('admin',$secret,''))
    $secret=$null; $credential=$null; $acc=0
    if(-not $signed.StartsWith('OK|')){throw 'Fixture Admin sign-in failed.'}
    $ready=[bool](Invoke-PackagedFixtureMacro $Excel "'invSys.Admin.xlam'!TestPackagedInventoryFixture.ContextReady" @([string]$hostBook.Name))
    if(-not $ready){throw 'Live captured Admin context unavailable; not meaningful RED.'}
    [void](Invoke-PackagedFixtureMacro $Excel "'invSys.Admin.xlam'!TestPackagedInventoryFixture.SetBinding" @($BindContext))
    $sourcePins=@()
    foreach($kind in @('Config','Auth','Data.Inventory')){
        $path=Join-Path $fixtureRoot ('WH-TEST.invSys.'+$kind+'.xlsb')
        $sourcePins+=[pscustomobject]@{Kind=$kind;Path=$path;Hash=(Get-PackagedFixtureHash $path)}
    }
    [pscustomobject]@{HostName=[string]$hostBook.Name;ContextReady=$ready;SourcePins=$sourcePins}
}

function Assert-PackagedInventorySmokeFixture {
    param($Fixture,[string]$EvidencePath)
    $facts=@($Fixture.SourcePins|ForEach-Object {[pscustomobject]@{Store=$_.Kind;BytesUnchanged=(Get-PackagedFixtureHash $_.Path) -ceq $_.Hash}})
    $facts|ConvertTo-Json|Set-Content -LiteralPath $EvidencePath
    if(@($facts|Where-Object {-not $_.BytesUnchanged}).Count){throw 'Packaged Inventory fixture source changed; diagnose before acceptance.'}
}
