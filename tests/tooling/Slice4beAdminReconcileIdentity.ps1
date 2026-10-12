# D14 prerequisite. Only disposable, unsaved packaged test instrumentation.
# Prompt answers are supplied at the presentation seam; the callback, lookup,
# payload builder and queue writer remain unchanged. No native-input claim.
function Install-AdminReconcileIdentityProbe {
    $project=$packages['invSys.Admin.xlam'].VBProject
    $probe=$project.VBComponents.Add(1);$probe.Name='TestReconcileIdentity'
    $probe.CodeModule.AddFromString(@'
Option Explicit
Private mInventory As Workbook
Private mWarehouse As String
Private mKeyA As String
Private mKeyB As String
Private mShipId As String
Private mPrompts As Long
Private mQueuedNotice As Boolean
Private Const SKU As String = "DEMO-RECONCILE-IDENTITY"

Public Function Prepare(ByVal warehouseId As String, ByVal csvPath As String) As Boolean
    Dim report As String, lo As ListObject, r As Long, item As Object
    mWarehouse = warehouseId
    If Not modAdminInventorySeed.UploadDemoInventoryForWarehouse(warehouseId, "S1", "config-admin", csvPath, report) Then Exit Function
    Set mInventory = modInventoryDomainBridge.ResolveInventoryWorkbookBridge(warehouseId)
    Set lo = FindTable(mInventory, "tblInventoryEntities")
    For r = 1 To lo.ListRows.Count
        If CStr(Value(lo, r, "SKU")) = SKU Then
            If CStr(Value(lo, r, "Location")) = "IDENTITY-A" Then mKeyA = CStr(Value(lo, r, "System_Key"))
            If CStr(Value(lo, r, "Location")) = "IDENTITY-B" Then mKeyB = CStr(Value(lo, r, "System_Key"))
        End If
    Next r
    If mKeyA = "" Or mKeyB = "" Or StrComp(mKeyA, mKeyB, vbBinaryCompare) = 0 Then Exit Function
    Set item = modRoleEventWriter.CreateInventoryEntityPayloadItem(mKeyA, SKU, 5, "IDENTITY-A", "GOOD")
    item("IoType") = "SHIP"
    mShipId = modRoleEventWriter.CreateSystemKey()
    If Not ApplyFixtureEvent("SHIP", mShipId, item) Then Exit Function
    Set item = modRoleEventWriter.CreateInventoryEntityPayloadItem(mKeyA, SKU, 3, "IDENTITY-A", "GOOD")
    item("IoType") = "ADJUST"
    If Not ApplyFixtureEvent("ADMIN_INVENTORY_ADJUST", modRoleEventWriter.CreateSystemKey(), item) Then Exit Function
    Prepare = (EntityQty(mKeyA) = 18 And EntityQty(mKeyB) = 30)
End Function

Private Function ApplyFixtureEvent(ByVal kind As String, ByVal eventId As String, ByVal item As Object) As Boolean
    Dim evt As Object, status As String, code As String, message As String
    Set evt = CreateObject("Scripting.Dictionary")
    evt("EventID") = eventId: evt("EventType") = kind
    evt("WarehouseId") = mWarehouse: evt("StationId") = "S1": evt("UserId") = "config-admin"
    evt("CreatedAtUTC") = Now: evt("PayloadJson") = modRoleEventWriter.BuildPayloadJson(item)
    ApplyFixtureEvent = modInventoryDomainBridge.ApplyInventoryEventBridge(evt, mInventory, "", status, code, message)
    ApplyFixtureEvent = ApplyFixtureEvent And status = "APPLIED"
End Function

Public Function PromptAnswer(ByVal title As String) As String
    mPrompts = mPrompts + 1
    Select Case title
        Case "invSys Admin - Shipment Reconcile": PromptAnswer = "1"
        Case "invSys Admin - Confirm Shipment Reconcile": PromptAnswer = "Generated training identity correction"
        Case Else: Err.Raise 5, , "Unexpected reconcile prompt title"
    End Select
End Function

Public Function ObserveNotice(ByVal text As String) As VbMsgBoxResult
    If InStr(1, text, "Shipment reconcile event queued.", vbBinaryCompare) = 1 Then mQueuedNotice = True
    ObserveNotice = vbOK
End Function

Public Function InvokeCallback() As Boolean
    mPrompts = 0: mQueuedNotice = False
    modAdminShipmentReconcile.OpenShipmentReconcileTool
    InvokeCallback = (mPrompts = 2 And mQueuedNotice)
End Function

Public Function FixtureValue(ByVal field As String) As String
    Select Case field
        Case "KeyA": FixtureValue = mKeyA
        Case "KeyB": FixtureValue = mKeyB
        Case "ShipId": FixtureValue = mShipId
        Case "InventoryPath": FixtureValue = mInventory.FullName
        Case "InboxPath": FixtureValue = modRoleEventWriter.ResolveInboxWorkbookPath("ADMIN_SHIPMENT_RECONCILE", mWarehouse, "S1")
    End Select
End Function

Public Function QuantitiesPreserved() As Boolean
    QuantitiesPreserved = (EntityQty(mKeyA) = 18 And EntityQty(mKeyB) = 30)
End Function

Public Sub ReleaseFixture()
    Set mInventory = Nothing
End Sub

Public Function StageCanary() As Boolean
    Dim lo As ListObject, added As ListColumn
    Set lo = FindTable(mInventory, "tblInventoryEntities")
    lo.Parent.Unprotect
    Set added = lo.ListColumns.Add
    added.Name = "Operator_Identity_Canary"
    added.DataBodyRange.Value2 = "preserve-generated-canary"
    mInventory.Save
    StageCanary = mInventory.Saved And CanaryPreserved()
End Function

Public Function CanaryPreserved() As Boolean
    Dim lo As ListObject, column As ListColumn, cell As Range
    Set lo = FindTable(mInventory, "tblInventoryEntities")
    For Each column In lo.ListColumns
        If column.Name = "Operator_Identity_Canary" Then
            For Each cell In column.DataBodyRange.Cells
                If CStr(cell.Value2) <> "preserve-generated-canary" Then Exit Function
            Next cell
            CanaryPreserved = True
            Exit Function
        End If
    Next column
End Function

Private Function EntityQty(ByVal systemKey As String) As Double
    Dim lo As ListObject, r As Long
    Set lo = FindTable(mInventory, "tblInventoryEntities")
    EntityQty = -999
    For r = 1 To lo.ListRows.Count
        If StrComp(CStr(Value(lo, r, "System_Key")), systemKey, vbBinaryCompare) = 0 Then
            EntityQty = CDbl(Value(lo, r, "QtyOnHand"))
            Exit Function
        End If
    Next r
End Function

Private Function FindTable(ByVal wb As Workbook, ByVal name As String) As ListObject
    Dim ws As Worksheet, lo As ListObject
    For Each ws In wb.Worksheets
        For Each lo In ws.ListObjects
            If lo.Name = name Then Set FindTable = lo: Exit Function
        Next lo
    Next ws
    Err.Raise 5, , "Generated fixture table unavailable"
End Function

Private Function Value(ByVal lo As ListObject, ByVal r As Long, ByVal column As String) As Variant
    Value = lo.DataBodyRange.Cells(r, lo.ListColumns(column).Index).Value2
End Function
'@)
    $owner=$project.VBComponents.Item('modAdminShipmentReconcile').CodeModule
    $owner.AddFromString(@'
Private Function InputBox(ByVal Prompt As String, Optional ByVal Title As String = "") As String
    InputBox = TestReconcileIdentity.PromptAnswer(Title)
End Function
Private Function MsgBox(ByVal Prompt As String, Optional ByVal Buttons As VbMsgBoxStyle = vbOKOnly, Optional ByVal Title As String = "") As VbMsgBoxResult
    MsgBox = TestReconcileIdentity.ObserveNotice(Prompt)
End Function
'@)
}

function Test-AdminReconcileIdentity($Fixture,$OtherFixture) {
    function Hash([string]$Path){$stream=[IO.File]::Open($Path,'Open','Read','ReadWrite');try{(Get-FileHash -InputStream $stream).Hash}finally{$stream.Dispose()}}
    try {
    SelectTarget $Fixture
    $csv=Join-Path $runRoot 'reconcile-demo.csv'
    @('ITEM_CODE,ITEM,UOM,LOCATION,QTY,CONDITION',
      'DEMO-RECONCILE-IDENTITY,Identity fixture,EA,IDENTITY-A,20,GOOD',
      'DEMO-RECONCILE-IDENTITY,Identity fixture,EA,IDENTITY-B,30,GOOD') | Set-Content -LiteralPath $csv -Encoding ASCII
    $ready=[bool](Run 'invSys.Admin.xlam' 'TestReconcileIdentity.Prepare' @($Fixture.Warehouse,$csv))
    Check 'Reconcile.Fixture.TwoOwnerCreatedEntitiesAndExactHistory' $ready
    if(-not $ready){throw 'Reconcile fixture prerequisites failed; not behavioral RED.'}
    $state=@{}
    foreach($field in @('KeyA','KeyB','ShipId','InventoryPath','InboxPath')){
        $state[$field]=[string](Run 'invSys.Admin.xlam' 'TestReconcileIdentity.FixtureValue' @($field))
    }
    $inventory=$excel.Workbooks.Item((Split-Path $state.InventoryPath -Leaf))
    $openedHere=$false;$inbox=$null
    foreach($book in @($excel.Workbooks)){if($book.FullName -ceq $state.InboxPath){$inbox=$book;break}}
    if($null -eq $inbox){$inbox=$excel.Workbooks.Open($state.InboxPath,0,$false);$openedHere=$true}
    $events=Table $inbox 'tblInboxShip';$beforeRows=$events.ListRows.Count
    $staged=[bool](Run 'invSys.Admin.xlam' 'TestReconcileIdentity.StageCanary')
    Check 'Reconcile.Fixture.SavedInventoryAndCustomColumn' $staged
    if(-not $staged){throw 'Fixture custom-column staging failed before the callback; not behavioral RED.'}
    $before=Hash $state.InventoryPath
    $otherBefore=Hash $OtherFixture.Config
    $called=[bool](Run 'invSys.Admin.xlam' 'TestReconcileIdentity.InvokeCallback')
    Check 'Reconcile.Callback.ReachedQueueConfirmation' $called
    Check 'Reconcile.Callback.OneQueuedEvent' ($events.ListRows.Count -eq $beforeRows+1)
    if(-not $called -or $events.ListRows.Count -ne $beforeRows+1){throw 'Callback did not reach queue inspection; classify separately from exact-key RED.'}
    $row=$events.ListRows.Item($events.ListRows.Count)
    $kind=[string]$row.Range.Cells.Item(1,$events.ListColumns.Item('EventType').Index).Value2
    $payload=[string]$row.Range.Cells.Item(1,$events.ListColumns.Item('PayloadJson').Index).Value2
    # Windows PowerShell 5 emits the parsed JSON array as one pipeline object.
    # Assign directly so property inspection targets the correction, not an array.
    $items=ConvertFrom-Json -InputObject $payload
    if($items -isnot [array]){throw 'Correction payload must be a JSON array.'}
    Check 'Reconcile.Payload.OneCorrectionLine' ($kind -ceq 'ADMIN_SHIPMENT_RECONCILE' -and $items.Count -eq 1)
    if($items.Count -ne 1){throw 'Unexpected correction payload shape.'}
    $item=$items[0];$names=@($item.PSObject.Properties.Name)
    if($item -isnot [pscustomobject]){throw 'Correction array element must be an object.'}
    $hasKey=$names -ccontains 'System_Key'
    Check 'Reconcile.Payload.ExactSelectedSystemKey' ($hasKey -and [string]$item.System_Key -ceq $state.KeyA)
    Check 'Reconcile.Payload.NoLegacyRow' (@($names|Where-Object{$_ -ieq 'ROW'}).Count -eq 0)
    Check 'Reconcile.Payload.SingleEventDeltaAndEvidence' ($item.Qty -eq -3 -and $item.CorrectedShipEventId -ceq $state.ShipId -and $item.RepairNarrative -ceq 'Generated training identity correction')
    Check 'Reconcile.Preservation.BothEntitiesUnchangedBeforeProcessing' ([bool](Run 'invSys.Admin.xlam' 'TestReconcileIdentity.QuantitiesPreserved'))
    Check 'Reconcile.Preservation.InventoryBytes' ((Hash $state.InventoryPath) -ceq $before)
    Check 'Reconcile.Preservation.InventoryNotDirtied' ([bool]$inventory.Saved)
    Check 'Reconcile.Preservation.UnknownColumn' ([bool](Run 'invSys.Admin.xlam' 'TestReconcileIdentity.CanaryPreserved'))
    Check 'Reconcile.Preservation.OtherWarehouseConfig' ((Hash $OtherFixture.Config) -ceq $otherBefore)
    if($openedHere){$inbox.Close($false)}
    } finally {
        [void](Run 'invSys.Admin.xlam' 'TestReconcileIdentity.ReleaseFixture')
    }
}
