# Existing demo lifecycle regressions on the Admin-generated outcome fixture.
# Callback arguments/results and workbook fingerprints remain in memory.
function Install-AdminSeedWorkflowProbe($Module) {
    $Module.AddFromString(@'
Public Function InvokeWorkflow(ByVal action As String, ByVal path As String, ByVal confirmed As Boolean) As String
    Dim result As String
    result = modAdmin.RunDemoInventoryActionCallbackForAutomation(mWarehouse, "S1", "config-admin", action, path, confirmed)
    InvokeWorkflow = Left$(result, InStr(1, result & "|", "|", vbBinaryCompare) - 1)
End Function

Public Function WorkflowFingerprint() As String
    Dim inventory As Workbook, names As Variant, name As Variant, lo As ListObject
    Dim row As Long, col As Long, text As String, valueText As String
    Set inventory = modInventoryDomainBridge.ResolveInventoryWorkbookBridge(mWarehouse)
    names = Array("tblInventoryEntities", "tblAppliedEvents", "tblInventoryLog")
    For Each name In names
        Set lo = FindTable(inventory, CStr(name))
        For row = 1 To lo.Range.Rows.Count
            For col = 1 To lo.ListColumns.Count
                valueText = CStr(lo.Range.Cells(row, col).Value2)
                text = text & CStr(Len(valueText)) & ":" & valueText
            Next col
        Next row
    Next name
    Set lo = FindTable(mInbox, "tblInboxProd")
    WorkflowFingerprint = CStr(lo.ListRows.Count) & "|" & text
End Function

Public Function ReceiptSchemaChecks() As String
    Dim wb As Workbook, lo As ListObject, added As ListRow, column As ListColumn
    Dim row As Long, matched As Long, count As Long, values As Variant
    Dim idName As String, statusName As String, oldStatus As String
    Dim exact As Boolean, duplicate As Boolean, pending As Boolean, absent As Boolean
    Dim normalized As Boolean, ambiguous As Boolean, missingTable As Boolean, borrowed As Boolean, owned As Boolean
    Set wb = modInventoryDomainBridge.ResolveInventoryWorkbookBridge(mWarehouse)
    wb.Save
    Set lo = FindTable(wb, "tblAppliedEvents")
    For row = 1 To lo.ListRows.Count
        If StrComp(CStr(Value(lo, row, "EventID")), mSeedEvent, vbBinaryCompare) = 0 Then matched = row
    Next row
    If matched = 0 Then Err.Raise 5, , "Saved receipt fixture unavailable"
    exact = modInventoryDomainBridge.InventoryEventAppliedBridge(mWarehouse, mSeedEvent)
    values = lo.ListRows(matched).Range.Value2
    Set added = lo.ListRows.Add: added.Range.Value2 = values: wb.Save
    duplicate = Not modInventoryDomainBridge.InventoryEventAppliedBridge(mWarehouse, mSeedEvent)
    added.Delete: wb.Save
    oldStatus = CStr(Value(lo, matched, "Status"))
    lo.ListColumns("Status").DataBodyRange.Cells(matched, 1).Value2 = "NEW": wb.Save
    pending = Not modInventoryDomainBridge.InventoryEventAppliedBridge(mWarehouse, mSeedEvent)
    lo.ListColumns("Status").DataBodyRange.Cells(matched, 1).Value2 = oldStatus
    idName = lo.ListColumns("EventID").Name: statusName = lo.ListColumns("Status").Name
    lo.ListColumns(idName).Name = "ReceiptReference": wb.Save
    absent = Not modInventoryDomainBridge.InventoryEventAppliedBridge(mWarehouse, mSeedEvent)
    absent = absent And wb.Saved And lo.ListColumns("ReceiptReference").Name = "ReceiptReference"
    lo.ListColumns("ReceiptReference").Name = " eventid "
    lo.ListColumns(statusName).Name = " status ": wb.Save
    normalized = modInventoryDomainBridge.InventoryEventAppliedBridge(mWarehouse, mSeedEvent)
    lo.ListColumns(" eventid ").Name = idName: lo.ListColumns(" status ").Name = statusName
    Set column = lo.ListColumns.Add: column.Name = " EventID ": wb.Save
    ambiguous = Not modInventoryDomainBridge.InventoryEventAppliedBridge(mWarehouse, mSeedEvent)
    column.Delete
    lo.Name = "TestSavedReceiptUnavailable": wb.Save
    missingTable = Not modInventoryDomainBridge.InventoryEventAppliedBridge(mWarehouse, mSeedEvent)
    missingTable = missingTable And wb.Saved And lo.Name = "TestSavedReceiptUnavailable"
    lo.Name = "tblAppliedEvents": wb.Save
    borrowed = modInventoryDomainBridge.InventoryEventAppliedBridge(mWarehouse, mSeedEvent)
    borrowed = borrowed And wb.Saved And wb.Worksheets.Count > 0
    wb.Close SaveChanges:=False: Set wb = Nothing
    count = Application.Workbooks.Count
    owned = modInventoryDomainBridge.InventoryEventAppliedBridge(mWarehouse, mSeedEvent)
    owned = owned And Application.Workbooks.Count = count
    ReceiptSchemaChecks = CStr(exact) & "|" & CStr(duplicate) & "|" & CStr(pending) & "|" & _
        CStr(absent) & "|" & CStr(normalized) & "|" & CStr(ambiguous) & "|" & _
        CStr(missingTable) & "|" & CStr(borrowed) & "|" & CStr(owned)
End Function

Public Function ReadExactReceipt() As Boolean
    ReadExactReceipt = modInventoryDomainBridge.InventoryEventAppliedBridge(mWarehouse, mSeedEvent)
End Function
'@)
}

function Test-AdminSeedWorkflow($Fixture) {
    function InvokeDemo([string]$Action,[string]$Path='',[bool]$Confirmed=$true){
        [string](Run 'invSys.Admin.xlam' 'TestSeedOutcome.InvokeWorkflow' @($Action,$Path,$Confirmed))
    }
    function Fingerprint {[string](Run 'invSys.Admin.xlam' 'TestSeedOutcome.WorkflowFingerprint')}
    function State {([string](Run 'invSys.Admin.xlam' 'TestSeedOutcome.State')).Split('|')}
    function Prepare([bool]$Earlier=$false){
        if(-not [bool](Run 'invSys.Admin.xlam' 'TestSeedOutcome.Prepare' @($Fixture.Warehouse,$Earlier))){throw 'Workflow fixture preparation failed; not behavioral RED.'}
    }
    function Release {[void](Run 'invSys.Admin.xlam' 'TestSeedOutcome.Release')}
    SelectTarget $Fixture
    Prepare
    try {
        $before=Fingerprint
        Check 'SeedWorkflow.DeleteEmpty.NoOpSuccess' ((InvokeDemo 'DELETE') -ceq 'OK')
        Check 'SeedWorkflow.DeleteEmpty.NoNewEventOrMutation' ((Fingerprint) -ceq $before)
        Check 'SeedWorkflow.Reseed.Success' ((InvokeDemo 'SEED') -ceq 'OK')
        $state=State
        Check 'SeedWorkflow.Reseed.ActiveAndUnique' ($state[7] -ceq '24' -and $state[6] -ceq 'True')
        $before=Fingerprint
        Check 'SeedWorkflow.RepeatSeed.NoOpSuccess' ((InvokeDemo 'SEED') -ceq 'OK')
        Check 'SeedWorkflow.RepeatSeed.NoNewEventOrMutation' ((Fingerprint) -ceq $before)
        Check 'SeedWorkflow.DeleteCancellation.Result' ((InvokeDemo 'DELETE' '' $false) -ceq 'CANCEL')
        Check 'SeedWorkflow.DeleteCancellation.NoMutation' ((Fingerprint) -ceq $before)
        Check 'SeedWorkflow.UploadCancellation.Result' ((InvokeDemo 'UPLOAD_DATASET') -ceq 'CANCEL')
        Check 'SeedWorkflow.UploadCancellation.NoMutation' ((Fingerprint) -ceq $before)

        $source=Join-Path $Fixture.Root 'workflow-demo.csv'
        $invalid=Join-Path $Fixture.Root 'workflow-invalid.csv'
        $csv="ITEM_CODE,ITEM,QTY,UOM,LOCATION`r`nDEMO-WORKFLOW-UPLOAD,Workflow upload,9,each,WORKFLOW`r`n"
        [IO.File]::WriteAllText($source,$csv,[Text.UTF8Encoding]::new($false))
        [IO.File]::WriteAllText($invalid,$csv+"REAL-WORKFLOW,Invalid second row,1,each,WORKFLOW`r`n",[Text.UTF8Encoding]::new($false))
        Check 'SeedWorkflow.InvalidCsv.RefusedBeforeQueue' ((InvokeDemo 'SEED' $invalid) -ceq 'FAIL' -and (Fingerprint) -ceq $before)
        Check 'SeedWorkflow.DataSetImport.Success' ((InvokeDemo 'UPLOAD_DATASET' $source) -ceq 'OK')
        $stored=Join-Path $Fixture.Root 'admin/demo-inventory-data-sets/workflow-demo.csv'
        Check 'SeedWorkflow.DataSetImport.StoredExactDefinition' ((Test-Path -LiteralPath $stored) -and (Get-FileHash $stored).Hash -ceq (Get-FileHash $source).Hash)
        Check 'SeedWorkflow.DataSetImport.InventoryUnchanged' ((Fingerprint) -ceq $before)
    }finally{Release}

    Prepare $true
    try {
        $before=State
        $ack=InvokeDemo 'SEED' $stored
        $state=State
        Check 'SeedWorkflow.UploadEarlierEvent.NoFalseSuccess' ($ack -ceq 'FAIL')
        Check 'SeedWorkflow.UploadEarlierEvent.ExactPending' ($state[0] -ceq '1' -and $state[1] -ceq 'NEW' -and $state[2] -ceq 'False' -and $state[3] -ceq 'True' -and $state[7] -ceq $before[7])
        $applied=[int](Run 'invSys.Admin.xlam' 'TestSeedOutcome.ApplyPending')
        $state=State
        Check 'SeedWorkflow.UploadEarlierEvent.LaterExactApplication' ($applied -eq 1 -and $state[2] -ceq 'True' -and $state[7] -ceq '25' -and $state[6] -ceq 'True')
        $before=Fingerprint
        Check 'SeedWorkflow.RepeatUpload.NoOpSuccess' ((InvokeDemo 'SEED' $stored) -ceq 'OK')
        Check 'SeedWorkflow.RepeatUpload.NoNewEventOrMutation' ((Fingerprint) -ceq $before)
        Check 'SeedWorkflow.DataSetDeleteCancellation.Result' ((InvokeDemo 'DELETE_DATASET' $stored $false) -ceq 'CANCEL')
        Check 'SeedWorkflow.DataSetDeleteCancellation.DefinitionAndStockRemain' ((Test-Path -LiteralPath $stored) -and (Fingerprint) -ceq $before)
        Check 'SeedWorkflow.DataSetDelete.Success' ((InvokeDemo 'DELETE_DATASET' $stored) -ceq 'OK')
        Check 'SeedWorkflow.DataSetDelete.OnlyDefinitionRemoved' (-not (Test-Path -LiteralPath $stored) -and (Test-Path -LiteralPath $source) -and (Fingerprint) -ceq $before)
        Check 'SeedWorkflow.BuiltInDataSetDelete.Refused' ((InvokeDemo 'DELETE_DATASET') -ceq 'FAIL')
        Check 'SeedWorkflow.BuiltInDataSetDelete.InventoryUnchanged' ((Fingerprint) -ceq $before)
    }finally{Release}
    $inventory=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Data.Inventory.xlsb')
    $receipts=([string](Run 'invSys.Admin.xlam' 'TestSeedOutcome.ReceiptSchemaChecks')).Split('|')
    foreach($entry in @(@('ExactSavedFixture',0),@('DuplicateReceiptRefused',1),@('UnappliedStatusRefused',2),@('MissingHeaderNoRepair',3),@('NormalizedHeadersAccepted',4),@('AmbiguousHeadersRefused',5),@('MissingTableNoRepair',6),@('BorrowedWorkbookRetained',7),@('OwnedReadClosesWorkbook',8))){
        Check ('SeedWorkflow.Receipt.'+$entry[0]) ($receipts[$entry[1]] -ceq 'True')
    }
    $beforeHash=(Get-FileHash -LiteralPath $inventory).Hash
    $beforeCount=$excel.Workbooks.Count
    $again=[bool](Run 'invSys.Admin.xlam' 'TestSeedOutcome.ReadExactReceipt')
    Check 'SeedWorkflow.Receipt.OwnedReadPreservesFile' ($again -and $excel.Workbooks.Count -eq $beforeCount -and (Get-FileHash -LiteralPath $inventory).Hash -ceq $beforeHash)
}
