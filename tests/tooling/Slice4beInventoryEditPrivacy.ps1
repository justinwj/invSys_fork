# Supplemental projection/service checks, not packaged role UI acceptance.
# Only the generated warehouse is altered. A private-table canary is diagnostic
# metadata, never synthetic business inventory or an accepted command receipt.
function Install-InventoryEditPrivacyProbe {
    $probe=$packages['invSys.Admin.xlam'].VBProject.VBComponents.Add(1)
    $probe.Name='TestInventoryEditPublication'
    $probe.CodeModule.AddFromString(@'
Option Explicit
Public Function Publish(ByVal warehouse As String, ByVal hostName As String, ByVal sourcePath As String) As Boolean
    Dim report As String, host As Workbook, source As Workbook
    On Error GoTo Finished
    Set host = Application.Workbooks(hostName)
    If Not modRoleWorkbookSurfaces.EnsureAdminLegacyWorkbookSurface(host, report) Then Exit Function
    host.Save
    Set source = Application.Workbooks.Open(sourcePath, UpdateLinks:=0, ReadOnly:=True)
    Publish = modAdminConsole.PublishWarehouseArtifacts("config-admin", warehouse, source, host, report)
Finished:
    On Error Resume Next
    If Not source Is Nothing Then source.Close False
End Function
'@)
    $probe=$packages['invSys.Core.xlam'].VBProject.VBComponents.Add(1)
    $probe.Name='TestInventoryEditPrivacy'
    $probe.CodeModule.AddFromString(@'
Option Explicit
Public Function RefreshRole(ByVal role As String, ByVal bookName As String, ByVal warehouse As String) As Boolean
    Dim book As Workbook, report As String, ready As Boolean
    Set book = Application.Workbooks(bookName)
    Select Case role
        Case "Receiving": ready = modRoleWorkbookSurfaces.EnsureReceivingWorkbookSurface(book, report)
        Case "Shipping": ready = modRoleWorkbookSurfaces.EnsureShippingWorkbookSurface(book, report)
        Case "Production": ready = modRoleWorkbookSurfaces.EnsureProductionWorkbookSurface(book, report)
    End Select
    If Not ready Then Exit Function
    RefreshRole = modOperatorReadModel.RefreshInventoryReadModelFromLocalForWorkbook(book, warehouse, report)
    book.Save
End Function
Public Function AggregateD17(ByVal snapshot As String, ByVal output As String) As Boolean
    Dim report As String
    AggregateD17 = modHqAggregator.GenerateGlobalSnapshotFromFiles(snapshot, output, report)
End Function
Public Function AggregateCaptured(ByVal hostName As String, ByVal warehouse As String, ByVal config As String, ByVal snapshot As String) As Boolean
    Dim context As String, report As String, outcome As String
    On Error GoTo Finished
    context = modAggregationCommands.CaptureForm(hostName)
    If context = "" Then Exit Function
    If modAggregationCommands.RegisterSource(context, warehouse, config, snapshot) = "" Then GoTo Finished
    AggregateCaptured = modAggregationCommands.Aggregate(context, snapshot, report, outcome)
Finished:
    If context <> "" Then modAggregationCommands.ReleaseForm context
End Function
'@)
}
function Test-InventoryPrivateCanary([string]$Value,[string]$Canary) {
    return (($Canary -ne '' -and $Value.IndexOf($Canary,[StringComparison]::OrdinalIgnoreCase) -ge 0) -or
        $Value.IndexOf('PrivateReceiptCanary',[StringComparison]::OrdinalIgnoreCase) -ge 0)
}
function Get-InventoryProjectionPrivacyFacts([string]$Path,[string]$TableName,[string]$Key,[string]$Canary) {
    $hash=Get-AdminInventoryEntryHash $Path
    $book=$excel.Workbooks.Open($Path,0,$true)
    $result=[pscustomobject]@{PrivateTable=$false;CanaryFound=$false;ExactKeyRows=0;Preserved=$false}
    try{
        foreach($sheet in $book.Worksheets){
            $values=$sheet.UsedRange.Value2
            foreach($value in $values){
                if(Test-InventoryPrivateCanary ([string]$value) $Canary){$result.CanaryFound=$true}
            }
            foreach($formula in $sheet.UsedRange.Formula){
                if(Test-InventoryPrivateCanary ([string]$formula) $Canary){$result.CanaryFound=$true}
            }
            foreach($table in $sheet.ListObjects){
                if($table.Name -ieq 'tblInventoryEditReceipts'){$result.PrivateTable=$true}
                if($table.Name -ieq $TableName){
                    foreach($row in @(Get-InventoryEditRows $table)){
                        if((Get-InventoryEditCell $row $table 'System_Key') -ceq $Key){$result.ExactKeyRows++}
                    }
                }
            }
        }
    }finally{$book.Close($false)}
    $result.Preserved=((Get-AdminInventoryEntryHash $Path) -ceq $hash)
    return $result
}
function Test-InventoryProjectionOutput([string]$Label,[string]$Path,[string]$TableName,[string]$Key,[string]$Canary) {
    Check ($Label+'.ArtifactExists') (Test-Path -LiteralPath $Path)
    if(-not(Test-Path -LiteralPath $Path)){throw 'Owned projection output unavailable; not privacy RED.'}
    $facts=Get-InventoryProjectionPrivacyFacts $Path $TableName $Key $Canary
    Check ($Label+'.ExactBusinessEntityPresent') ($facts.ExactKeyRows -eq 1)
    Check ($Label+'.PrivateReceiptTableExcluded') (-not $facts.PrivateTable)
    Check ($Label+'.PrivateCanaryExcluded') (-not $facts.CanaryFound)
    Check ($Label+'.InspectionPreservedBytes') $facts.Preserved
}
function Test-InventoryEditPrivacy($Fixture,$HostBook,[string]$Key) {
    SelectTarget $Fixture
    $sourcePath=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Data.Inventory.xlsb')
    $canary='OWNED-PRIVATE-RECEIPT-CANARY'
    $source=$excel.Workbooks.Open($sourcePath,0,$false)
    try{
        $tables=@($source.Worksheets|ForEach-Object {$_.ListObjects}|Where-Object {$_.Name -ieq 'tblInventoryEditReceipts'})
        if($tables.Count -gt 1){throw 'Ambiguous owned private table fixture.'}
        if($tables.Count -eq 0){
            # The pre-implementation source has no receipt table. This deliberately
            # incomplete canary table tests exclusion, not valid receipt creation.
            $sheet=$source.Worksheets.Add();$sheet.Name='OwnedPrivateReceipt'
            $sheet.Cells.Item(1,1).Value2='PrivateReceiptCanary'
            $sheet.Cells.Item(1,2).Value2='UntrustedFixtureMarker'
            $sheet.Cells.Item(2,1).Value2=$canary
            $sheet.Cells.Item(2,2).Value2='NOT-A-COMMAND-RECEIPT'
            $private=$sheet.ListObjects.Add(1,$sheet.Range('A1:B2'),$null,1)
            $private.Name='tblInventoryEditReceipts'
        }else{
            $private=$tables[0]
            $column=$private.ListColumns.Add();$column.Name='PrivateReceiptCanary'
            if($private.ListRows.Count -eq 0){throw 'Owning EDIT did not produce a receipt row; canary fixture unavailable.'}
            $column.DataBodyRange.Value2=$canary
        }
        $source.Save()
    }finally{$source.Close($false)}
    $before=Get-AdminInventoryEntryHash $sourcePath
    $canaryProof=Get-InventoryProjectionPrivacyFacts $sourcePath 'tblInventoryEntities' $Key $canary
    $ready=$canaryProof.PrivateTable -and $canaryProof.CanaryFound -and $canaryProof.ExactKeyRows -eq 1 -and $canaryProof.Preserved
    Check 'InventoryEditPrivacy.OwnedCanaryAndRealEntityPresent' $ready
    if(-not $ready){throw 'Private projection canary prerequisite unavailable; not RED.'}
    $published=[bool](Run 'invSys.Admin.xlam' 'TestInventoryEditPublication.Publish' @($Fixture.Warehouse,$HostBook.Name,$sourcePath))
    Check 'InventoryEditPrivacy.AdminPublicationCompleted' $published
    if(-not $published){throw 'Existing publication control failed; privacy outputs unavailable.'}
    $name=$Fixture.Warehouse+'.invSys.Snapshot.Inventory.xlsb'
    $local=Join-Path $Fixture.Root $name
    $share=$Fixture.Root+'-share'
    $remote=Join-Path (Join-Path $share 'Snapshots') $name
    Test-InventoryProjectionOutput 'InventoryEditPrivacy.LocalSnapshot' $local 'tblInventorySnapshot' $Key $canary
    Test-InventoryProjectionOutput 'InventoryEditPrivacy.PublishedSnapshot' $remote 'tblInventorySnapshot' $Key $canary
    foreach($role in @('Receiving','Shipping','Production')){
        $book=$excel.Workbooks.Add()
        $path=Join-Path $runRoot ('private-projection-'+$role+'.xlsx');$book.SaveAs($path,51)
        try{
            $refreshed=[bool](Run 'invSys.Core.xlam' 'TestInventoryEditPrivacy.RefreshRole' @($role,$book.Name,$Fixture.Warehouse))
            Check ('InventoryEditPrivacy.'+$role+'.RefreshCompleted') $refreshed
        }finally{$book.Close($false)}
        Test-InventoryProjectionOutput ('InventoryEditPrivacy.'+$role) $path 'invSys' $Key $canary
    }
    $d17=Join-Path $runRoot 'private-projection-d17.xlsb'
    $aggregated=[bool](Run 'invSys.Core.xlam' 'TestInventoryEditPrivacy.AggregateD17' @($remote,$d17))
    Check 'InventoryEditPrivacy.D17.Completed' $aggregated
    Test-InventoryProjectionOutput 'InventoryEditPrivacy.D17' $d17 'tblGlobalInventorySnapshot' $Key $canary
    SelectTarget $Fixture
    $HostBook.Activate();$HostBook.Save()
    $aggregated=[bool](Run 'invSys.Core.xlam' 'TestInventoryEditPrivacy.AggregateCaptured' @($HostBook.Name,$Fixture.Warehouse,$Fixture.Config,$remote))
    Check 'InventoryEditPrivacy.D8CapturedAggregation.Completed' $aggregated
    $d8=Join-Path (Join-Path $share 'Global') 'invSys.Global.InventorySnapshot.xlsb'
    Test-InventoryProjectionOutput 'InventoryEditPrivacy.D8CapturedAggregation' $d8 'tblGlobalInventorySnapshot' $Key $canary
    Check 'InventoryEditPrivacy.CanonicalBytesUnchangedByAllProjections' ((Get-AdminInventoryEntryHash $sourcePath) -ceq $before)
}
