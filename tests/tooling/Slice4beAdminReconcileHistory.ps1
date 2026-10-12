# D14-RECON-01: real generated history, actual callback, and Domain guard cases.
# Only counts, booleans and check identities enter the receipts.
function Install-AdminReconcileHistoryProbe($Module) {
    $Module.InsertLines($Module.CountOfDeclarationLines+1,"Private mHistoryShipId As String`r`nPrivate mNonShipId As String")
    $Module.AddFromString(@'
Public Function PrepareHistoryScenario(ByVal scenario As String) As Boolean
    Dim item As Object, items As Collection, laterShip As String
    Select Case scenario
        Case "OlderShipment"
            laterShip = modRoleEventWriter.CreateSystemKey()
            Set item = HistoryItem(mKeyA, 4, "IDENTITY-A", "GOOD", "SHIP")
            If Not ApplyFixtureEvent("SHIP", laterShip, item) Then Exit Function
            Set item = HistoryItem(mKeyA, 7, "IDENTITY-A", "GOOD")
            If Not ApplyFixtureEvent("ADMIN_INVENTORY_ADJUST", modRoleEventWriter.CreateSystemKey(), item) Then Exit Function
            mNonShipId = modRoleEventWriter.CreateSystemKey()
            Set item = HistoryItem(mKeyB, 9, "IDENTITY-B", "GOOD")
            If Not ApplyFixtureEvent("ADMIN_INVENTORY_ADJUST", mNonShipId, item) Then Exit Function
            mHistoryShipId = mShipId: mSelection = 2
            PrepareHistoryScenario = (EntityQty(mKeyA) = 21 And EntityQty(mKeyB) = 39)
        Case "MultiEntityAndLines"
            Set items = New Collection
            items.Add HistoryItem(mKeyA, 2, "IDENTITY-A", "GOOD", "SHIP")
            items.Add HistoryItem(mKeyB, 4, "IDENTITY-B", "GOOD", "SHIP")
            items.Add HistoryItem(mKeyA, 1, "IDENTITY-A", "GOOD", "SHIP")
            mHistoryShipId = modRoleEventWriter.CreateSystemKey()
            If Not ApplyHistoryPayload("SHIP", mHistoryShipId, modRoleEventWriter.BuildPayloadJsonFromCollection(items)) Then Exit Function
            Set item = HistoryItem(mKeyA, 4, "CURRENT-A", "DAMAGED")
            If Not ApplyFixtureEvent("ADMIN_INVENTORY_ADJUST", modRoleEventWriter.CreateSystemKey(), item) Then Exit Function
            Set item = HistoryItem(mKeyB, 20, "CURRENT-B", "GOOD")
            If Not ApplyFixtureEvent("ADMIN_INVENTORY_ADJUST", modRoleEventWriter.CreateSystemKey(), item) Then Exit Function
            mSelection = 1
            PrepareHistoryScenario = (EntityQty(mKeyA) = 22 And EntityQty(mKeyB) = 55)
    End Select
    If PrepareHistoryScenario Then mInventory.Save
End Function

Private Function HistoryItem(ByVal systemKey As String, ByVal qty As Double, ByVal location As String, ByVal condition As String, Optional ByVal ioType As String = "ADJUST") As Object
    Dim item As Object
    Set item = modRoleEventWriter.CreateInventoryEntityPayloadItem(systemKey, SKU, qty, location, condition)
    item("IoType") = ioType
    Set HistoryItem = item
End Function

Private Function ApplyHistoryPayload(ByVal kind As String, ByVal eventId As String, ByVal payload As String) As Boolean
    Dim evt As Object, status As String, code As String, message As String
    Set evt = CreateObject("Scripting.Dictionary")
    evt("EventID") = eventId: evt("EventType") = kind
    evt("WarehouseId") = mWarehouse: evt("StationId") = "S1": evt("UserId") = "config-admin"
    evt("CreatedAtUTC") = Now: evt("PayloadJson") = payload
    ApplyHistoryPayload = modInventoryDomainBridge.ApplyInventoryEventBridge(evt, mInventory, "", status, code, message)
    ApplyHistoryPayload = ApplyHistoryPayload And status = "APPLIED"
End Function

Public Function SelectedHistoryEventForTest() As String
    SelectedHistoryEventForTest = mHistoryShipId
End Function
Public Function HistoryQuantitiesEqual(ByVal expectedA As Double, ByVal expectedB As Double) As Boolean
    HistoryQuantitiesEqual = (EntityQty(mKeyA) = expectedA And EntityQty(mKeyB) = expectedB)
End Function
Public Function HistoryPromptsShowExactKey() As Boolean
    HistoryPromptsShowExactKey = (InStr(1, mListPrompt, mKeyA, vbBinaryCompare) > 0 And InStr(1, mConfirmPrompt, mKeyA, vbBinaryCompare) > 0)
End Function

Public Function DomainGuardForTest(ByVal scenario As String) As String
    Dim item As Object, beforeA As Double, beforeB As Double, beforeLog As Long, beforeApplied As Long
    Dim applied As Boolean, sameQty As Boolean, sameHistory As Boolean
    beforeA = EntityQty(mKeyA): beforeB = EntityQty(mKeyB)
    beforeLog = FindTable(mInventory, "tblInventoryLog").ListRows.Count
    beforeApplied = FindTable(mInventory, "tblAppliedEvents").ListRows.Count
    Set item = HistoryItem(mKeyA, -1, "CURRENT-A", "DAMAGED")
    item("CorrectedShipEventId") = mShipId
    item("RepairNarrative") = "Generated invalid-association probe"
    Select Case scenario
        Case "WrongEntity": item("System_Key") = mKeyB
        Case "NonShipmentEvent": item("CorrectedShipEventId") = mNonShipId
        Case "PositiveCorrection": item("Qty") = 1
        Case "MissingNarrative": item("RepairNarrative") = ""
        Case Else: Err.Raise 5, , "Unknown Domain guard fixture"
    End Select
    applied = ApplyFixtureEvent("ADMIN_SHIPMENT_RECONCILE", modRoleEventWriter.CreateSystemKey(), item)
    sameQty = (EntityQty(mKeyA) = beforeA And EntityQty(mKeyB) = beforeB)
    sameHistory = (FindTable(mInventory, "tblInventoryLog").ListRows.Count = beforeLog And FindTable(mInventory, "tblAppliedEvents").ListRows.Count = beforeApplied)
    DomainGuardForTest = CStr(Not applied) & "|" & CStr(sameQty) & "|" & CStr(sameHistory)
End Function

Public Function ApplyQueuedHistoryCorrection(ByVal inboxName As String) As String
    Dim lo As ListObject, evt As Object, field As Variant, applied As Boolean
    Dim status As String, code As String, message As String, beforeLog As Long
    Dim quantities As Boolean, attributes As Boolean, custom As Boolean, duplicate As Boolean
    Set lo = FindTable(Application.Workbooks(inboxName), "tblInboxShip")
    Set evt = CreateObject("Scripting.Dictionary")
    For Each field In Array("EventID", "EventType", "WarehouseId", "StationId", "UserId", "CreatedAtUTC", "PayloadJson")
        evt(CStr(field)) = lo.DataBodyRange.Cells(lo.ListRows.Count, lo.ListColumns(CStr(field)).Index).Value
    Next field
    applied = modInventoryDomainBridge.ApplyInventoryEventBridge(evt, mInventory, "", status, code, message)
    applied = applied And status = "APPLIED"
    quantities = (EntityQty(mKeyA) = 18 And EntityQty(mKeyB) = 55)
    attributes = CurrentHistoryAttributesPreserved()
    custom = CanaryPreserved()
    beforeLog = FindTable(mInventory, "tblInventoryLog").ListRows.Count
    status = "": code = "": message = ""
    duplicate = modInventoryDomainBridge.ApplyInventoryEventBridge(evt, mInventory, "", status, code, message)
    duplicate = duplicate And status = "SKIP_DUP" And beforeLog = FindTable(mInventory, "tblInventoryLog").ListRows.Count
    ApplyQueuedHistoryCorrection = CStr(applied) & "|" & CStr(quantities) & "|" & CStr(attributes) & "|" & CStr(custom) & "|" & CStr(duplicate)
End Function

Private Function CurrentHistoryAttributesPreserved() As Boolean
    Dim lo As ListObject, r As Long, goodA As Boolean, goodB As Boolean
    Set lo = FindTable(mInventory, "tblInventoryEntities")
    For r = 1 To lo.ListRows.Count
        If StrComp(CStr(Value(lo, r, "System_Key")), mKeyA, vbBinaryCompare) = 0 Then
            goodA = (CStr(Value(lo, r, "Location")) = "CURRENT-A" And CStr(Value(lo, r, "Condition")) = "DAMAGED")
        ElseIf StrComp(CStr(Value(lo, r, "System_Key")), mKeyB, vbBinaryCompare) = 0 Then
            goodB = (CStr(Value(lo, r, "Location")) = "CURRENT-B" And CStr(Value(lo, r, "Condition")) = "GOOD")
        End If
    Next r
    CurrentHistoryAttributesPreserved = goodA And goodB
End Function
'@)
}

function Test-AdminReconcileHistory($Events,$Inventory,$State) {
    foreach($case in @(
        @{Name='OlderShipment';A=21;B=39;Delta=-6;Location='IDENTITY-A';Condition='GOOD'},
        @{Name='MultiEntityAndLines';A=22;B=55;Delta=-4;Location='CURRENT-A';Condition='DAMAGED'}
    )){
        $ready=[bool](Run 'invSys.Admin.xlam' 'TestReconcileIdentity.PrepareHistoryScenario' @($case.Name))
        Check ('Reconcile.History.'+$case.Name+'.OwnerGeneratedFixture') $ready
        if(-not $ready){throw 'History fixture did not reach independent expected quantities; not behavioral RED.'}
        $priorCanary=[bool](Run 'invSys.Admin.xlam' 'TestReconcileIdentity.CanaryPreserved')
        [pscustomobject]@{Scenario=$case.Name;PriorProjectionCanaryRetainedDuringFixtureEvents=$priorCanary}|ConvertTo-Json -Compress|Add-Content (Join-Path $reportRoot 'history-fixture-observations.jsonl')
        $staged=[bool](Run 'invSys.Admin.xlam' 'TestReconcileIdentity.StageCanary')
        Check ('Reconcile.History.'+$case.Name+'.SavedCanaryBeforeInspection') $staged
        if(-not $staged){throw 'History inspection canary baseline failed.'}
        $eventId=[string](Run 'invSys.Admin.xlam' 'TestReconcileIdentity.SelectedHistoryEventForTest')
        $beforeRows=$Events.ListRows.Count
        $before=Hash $State.InventoryPath
        $called=[bool](Run 'invSys.Admin.xlam' 'TestReconcileIdentity.InvokeCallback')
        Check ('Reconcile.History.'+$case.Name+'.OneQueuedCorrection') ($called -and $Events.ListRows.Count -eq $beforeRows+1)
        if(-not $called -or $Events.ListRows.Count -ne $beforeRows+1){throw 'History callback did not reach queue inspection.'}
        $row=$Events.ListRows.Item($Events.ListRows.Count)
        $items=ConvertFrom-Json -InputObject ([string]$row.Range.Cells.Item(1,$Events.ListColumns.Item('PayloadJson').Index).Value2)
        if($items -isnot [array] -or $items.Count -ne 1 -or $items[0] -isnot [pscustomobject]){throw 'History correction payload shape is invalid.'}
        $item=$items[0];$names=@($item.PSObject.Properties.Name)
        Check ('Reconcile.History.'+$case.Name+'.ExactEntityAndEvent') ($names -ccontains 'System_Key' -and $item.System_Key -ceq $State.KeyA -and $item.CorrectedShipEventId -ceq $eventId)
        Check ('Reconcile.History.'+$case.Name+'.EntityOnlyAfterEventDelta') ($item.Qty -eq $case.Delta)
        Check ('Reconcile.History.'+$case.Name+'.ExactKeyVisibleInBothPrompts') ([bool](Run 'invSys.Admin.xlam' 'TestReconcileIdentity.HistoryPromptsShowExactKey'))
        Check ('Reconcile.History.'+$case.Name+'.CurrentAttributesInPayload') ($names -ccontains 'Condition' -and $item.Condition -ceq $case.Condition -and $item.Location -ceq $case.Location)
        Check ('Reconcile.History.'+$case.Name+'.QuantitiesAndCustomColumnPreserved') ([bool](Run 'invSys.Admin.xlam' 'TestReconcileIdentity.HistoryQuantitiesEqual' @($case.A,$case.B)) -and [bool](Run 'invSys.Admin.xlam' 'TestReconcileIdentity.CanaryPreserved'))
        Check ('Reconcile.History.'+$case.Name+'.InventoryBytesPreserved') ((Hash $State.InventoryPath) -ceq $before)
        Check ('Reconcile.History.'+$case.Name+'.InventoryNotDirtied') ([bool]$Inventory.Saved)
    }
    $values=([string](Run 'invSys.Admin.xlam' 'TestReconcileIdentity.ApplyQueuedHistoryCorrection' @($Events.Parent.Parent.Name))) -split '\|'
    if($values.Count -ne 5){throw 'Invalid queued correction application envelope.'}
    $checks=@('AppliedThroughDomain','ExactEntityOnlyQuantity','CurrentAttributesPreserved','CustomColumnValuesPreserved','DuplicateDoesNotApplyTwice')
    for($i=0;$i -lt $checks.Count;$i++){
        Check ('Reconcile.QueuedApplication.'+$checks[$i]) ([bool]::Parse($values[$i]))
    }
    foreach($case in @('WrongEntity','NonShipmentEvent','PositiveCorrection','MissingNarrative')){
        $values=([string](Run 'invSys.Admin.xlam' 'TestReconcileIdentity.DomainGuardForTest' @($case))) -split '\|'
        if($values.Count -ne 3){throw 'Invalid Domain guard observation envelope.'}
        Check ('Reconcile.Domain.'+$case+'.Rejected') ([bool]::Parse($values[0]))
        Check ('Reconcile.Domain.'+$case+'.QuantitiesPreserved') ([bool]::Parse($values[1]))
        Check ('Reconcile.Domain.'+$case+'.HistoryPreserved') ([bool]::Parse($values[2]))
    }
}
