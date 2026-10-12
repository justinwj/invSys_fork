Attribute VB_Name = "modAdminShipmentReconcile"
Option Explicit

Private Const ADMIN_RECON_EVENT_TYPE As String = "ADMIN_SHIPMENT_RECONCILE"
Private Const SERVER_SHIP_EVENT_TYPE As String = "SHIP"
Private Const SERVER_SHIP_RESERVE_EVENT_TYPE As String = "SHIP_RESERVE"
Private Const SERVER_SHIP_RELEASE_EVENT_TYPE As String = "SHIP_RELEASE"
Private Const FLAG_PROJECTED_BELOW_NAS_NO_ACTIVE_ROW As String = "PROJECTED_BELOW_NAS_NO_ACTIVE_ROW"
Private Const FLAG_LOCKED_WITH_NO_ACTIVE_ROW As String = "LOCKED_WITH_NO_ACTIVE_ROW"
Private Const FLAG_COMPLETED_RESERVATION_STILL_ACTIVE As String = "COMPLETED_RESERVATION_STILL_ACTIVE"
Private Const FLAG_NAS_INCREASED_AFTER_SHIP As String = "NAS_INCREASED_AFTER_SHIP"

Public Sub OpenShipmentReconcileTool()
    On Error GoTo Failed
    Dim target As WarehouseTarget, warehouseId As String, stationId As String, runtimeRoot As String
    Dim userId As String, recentRows As Collection, selectedRow As Object, state As Variant
    Dim selectionText As String, selectedIndex As Long, systemKey As String, correctedShipEventId As String
    Dim currentQty As Double, afterQty As Double, delta As Double, narrative As String, eventIdOut As String, report As String
    Set target = modNasConnection.GetCurrentTarget()
    If target Is Nothing Then
        MsgBox "Select a warehouse before opening Shipment Reconcile.", vbExclamation, "invSys Admin"
        Exit Sub
    End If
    warehouseId = target.WarehouseId: stationId = target.StationId: runtimeRoot = target.RuntimeRoot
    userId = modRoleEventWriter.ResolveCurrentUserId()
    Set recentRows = BuildRecentShipmentSentRows(Nothing, 20, warehouseId, runtimeRoot)
    If recentRows.Count = 0 Then
        MsgBox "No valid server shipment deduction entries are available in the selected warehouse.", vbInformation, "invSys Admin - Shipment Reconcile"
        Exit Sub
    End If
    selectionText = Trim$(InputBox("Enter the number of the server shipment deduction log line to inspect." & vbCrLf & vbCrLf & _
        FormatRecentShipmentRows(recentRows), "invSys Admin - Shipment Reconcile"))
    If selectionText = "" Then Exit Sub
    If Not IsNumeric(selectionText) Then GoTo InvalidSelection
    If CDbl(selectionText) < 1 Or CDbl(selectionText) > recentRows.Count Then GoTo InvalidSelection
    If CDbl(selectionText) <> Fix(CDbl(selectionText)) Then GoTo InvalidSelection
    selectedIndex = CLng(selectionText)
    Set selectedRow = recentRows(selectedIndex)
    systemKey = CStr(selectedRow("System_Key")): correctedShipEventId = CStr(selectedRow("EventID"))
    state = modInventoryDomainBridge.ReadShipmentReconcileBridge("STATE" & vbTab & correctedShipEventId & vbTab & systemKey, _
        Nothing, warehouseId, runtimeRoot)
    If IsEmpty(state) Then
        MsgBox "The selected shipment and exact inventory entity could not be validated.", vbExclamation, "invSys Admin - Shipment Reconcile"
        Exit Sub
    End If
    currentQty = CDbl(state(0)): afterQty = CDbl(state(1))
    If currentQty <= afterQty Then
        MsgBox "Current quantity is not higher than the selected entity's balance after that shipment. No correction was queued.", vbInformation, "invSys Admin - Shipment Reconcile"
        Exit Sub
    End If
    delta = afterQty - currentQty
    narrative = Trim$(InputBox("Type one line explaining why this admin correction is being made." & vbCrLf & vbCrLf & _
        "Corrects server shipment deduction event: " & correctedShipEventId & vbCrLf & _
        "System_Key: " & systemKey & vbCrLf & _
        "Current Qty: " & Format$(currentQty, "0.########") & vbCrLf & _
        "Qty after that event: " & Format$(afterQty, "0.########") & vbCrLf & _
        "Queued correction delta: " & Format$(delta, "0.########"), "invSys Admin - Confirm Shipment Reconcile"))
    If narrative = "" Then
        MsgBox "Shipment reconcile cancelled. A human repair narrative is required.", vbExclamation, "invSys Admin"
        Exit Sub
    End If
    If QueueAdminShipmentReconcileEvent(warehouseId, stationId, userId, systemKey, delta, correctedShipEventId, narrative, _
        FLAG_NAS_INCREASED_AFTER_SHIP, Nothing, eventIdOut, report, runtimeRoot) Then
        MsgBox "Shipment reconcile event queued." & vbCrLf & "EventID: " & eventIdOut & vbCrLf & _
            "Corrects: " & correctedShipEventId & vbCrLf & "System_Key: " & systemKey & vbCrLf & _
            "Delta: " & Format$(delta, "0.########"), vbInformation, "invSys Admin"
    Else
        MsgBox report, vbExclamation, "invSys Admin"
    End If
    Exit Sub
InvalidSelection:
    MsgBox "Enter a whole number from 1 to " & CStr(recentRows.Count) & ".", vbExclamation, "invSys Admin - Shipment Reconcile"
    Exit Sub
Failed:
    MsgBox "Shipment reconcile could not finish: " & Err.Description, vbExclamation, "invSys Admin"
End Sub

Public Function QueueAdminShipmentReconcileEvent(ByVal warehouseId As String, ByVal stationId As String, ByVal userId As String, _
                                                 ByVal systemKey As String, ByVal signedQtyDelta As Double, _
                                                 ByVal correctedShipEventId As String, ByVal repairNarrative As String, _
                                                 Optional ByVal mismatchFlag As String = "", Optional ByVal targetInboxWb As Workbook = Nothing, _
                                                 Optional ByRef eventIdOut As String = "", Optional ByRef report As String = "", _
                                                 Optional ByVal runtimeRoot As String = "") As Boolean
    Dim state As Variant, item As Object, payloadJson As String
    state = modInventoryDomainBridge.ReadShipmentReconcileBridge("STATE" & vbTab & correctedShipEventId & vbTab & systemKey, _
        Nothing, warehouseId, runtimeRoot)
    If IsEmpty(state) Then
        report = "The selected shipment and exact inventory entity could not be validated."
        Exit Function
    End If
    If Not ValidateShipmentReconcileRequest(CStr(state(2)), signedQtyDelta, correctedShipEventId, repairNarrative, report) Then Exit Function
    If signedQtyDelta <> CDbl(state(1)) - CDbl(state(0)) Then
        report = "Inventory quantity changed. Inspect the shipment again before queuing a correction."
        Exit Function
    End If
    repairNarrative = Trim$(repairNarrative)
    Set item = modRoleEventWriter.CreateInventoryEntityPayloadItem(systemKey, CStr(state(2)), signedQtyDelta, _
        CStr(state(3)), CStr(state(4)), CStr(state(5)), repairNarrative)
    item("IoType") = "RECONCILE"
    item("CorrectedShipEventId") = correctedShipEventId
    item("RepairNarrative") = repairNarrative
    If mismatchFlag <> "" Then item("MismatchFlag") = mismatchFlag
    payloadJson = modRoleEventWriter.BuildPayloadJson(item)
    QueueAdminShipmentReconcileEvent = modRoleEventWriter.QueuePayloadEvent(ADMIN_RECON_EVENT_TYPE, warehouseId, stationId, userId, _
        payloadJson, repairNarrative, correctedShipEventId, "", Now, targetInboxWb, eventIdOut, report)
End Function

Public Function ValidateShipmentReconcileRequest(ByVal sku As String, _
                                                 ByVal signedQtyDelta As Double, _
                                                 ByVal correctedShipEventId As String, _
                                                 ByVal repairNarrative As String, _
                                                 Optional ByRef report As String = "") As Boolean
    If Trim$(sku) = "" Then
        report = "SKU is required."
        Exit Function
    End If
    If signedQtyDelta >= 0 Then
        report = "Shipment correction delta must be negative."
        Exit Function
    End If
    If Trim$(correctedShipEventId) = "" Then
        report = "CorrectedShipEventId is required."
        Exit Function
    End If
    If Trim$(repairNarrative) = "" Then
        report = "A one-line human repair narrative is required."
        Exit Function
    End If

    ValidateShipmentReconcileRequest = True
End Function

Public Function BuildShipmentMismatchFlags(ByVal projectedQty As Double, _
                                           ByVal nasQty As Double, _
                                           ByVal lockedQty As Double, _
                                           ByVal hasActiveShipmentRow As Boolean, _
                                           ByVal completedReservationStillActive As Boolean, _
                                           ByVal nasIncreasedAfterShip As Boolean) As String
    Dim flags As String

    If projectedQty < nasQty And Not hasActiveShipmentRow Then AddFlag flags, FLAG_PROJECTED_BELOW_NAS_NO_ACTIVE_ROW
    If lockedQty > 0 And Not hasActiveShipmentRow Then AddFlag flags, FLAG_LOCKED_WITH_NO_ACTIVE_ROW
    If completedReservationStillActive Then AddFlag flags, FLAG_COMPLETED_RESERVATION_STILL_ACTIVE
    If nasIncreasedAfterShip Then AddFlag flags, FLAG_NAS_INCREASED_AFTER_SHIP
    BuildShipmentMismatchFlags = flags
End Function

Public Function BuildRecentShipmentSentLogText(ByVal inventoryWb As Workbook, Optional ByVal limitCount As Long = 20) As String
    BuildRecentShipmentSentLogText = FormatRecentShipmentRows(BuildRecentShipmentSentRows(inventoryWb, limitCount))
End Function

Private Function FormatRecentShipmentRows(ByVal rows As Collection) As String
    Dim row As Object, i As Long, text As String
    For i = 1 To rows.Count
        Set row = rows(i)
        If text <> "" Then text = text & vbCrLf
        text = text & CStr(i) & ". " & CStr(row("EventID")) & " | System_Key: " & CStr(row("System_Key")) & _
            " | " & CStr(row("SKU")) & " | Qty " & Format$(Abs(CDbl(row("QtyDelta"))), "0.########") & _
            " | " & CStr(row("Location")) & " | " & CStr(row("OccurredAtUTC"))
    Next i
    If text = "" Then text = "No valid server shipment deduction entries are available."
    FormatRecentShipmentRows = text
End Function

Public Function DetectNasIncreaseAfterLastShip(ByVal inventoryWb As Workbook, ByVal systemKey As String, _
                                               ByRef correctedShipEventId As String, ByRef currentNasQty As Double, _
                                               ByRef qtyAfterShip As Double, Optional ByRef report As String = "") As Boolean
    Dim rows As Collection, row As Object
    Set rows = BuildRecentShipmentSentRows(inventoryWb, 2147483647)
    For Each row In rows
        If StrComp(CStr(row("System_Key")), systemKey, vbBinaryCompare) = 0 Then
            correctedShipEventId = CStr(row("EventID"))
            DetectNasIncreaseAfterLastShip = DetectNasIncreaseAfterShipEvent(inventoryWb, correctedShipEventId, systemKey, currentNasQty, qtyAfterShip, report)
            Exit Function
        End If
    Next row
    report = "No shipment history was found for the exact inventory entity."
End Function

Public Function DetectNasIncreaseAfterShipEvent(ByVal inventoryWb As Workbook, ByVal shipEventId As String, ByVal systemKey As String, _
                                                ByRef currentNasQty As Double, ByRef qtyAfterShip As Double, _
                                                Optional ByRef report As String = "") As Boolean
    Dim state As Variant
    state = modInventoryDomainBridge.ReadShipmentReconcileBridge("STATE" & vbTab & shipEventId & vbTab & systemKey, inventoryWb)
    If IsEmpty(state) Then
        report = "The selected shipment and exact inventory entity could not be validated."
        Exit Function
    End If
    currentNasQty = CDbl(state(0)): qtyAfterShip = CDbl(state(1))
    DetectNasIncreaseAfterShipEvent = (currentNasQty > qtyAfterShip)
    If Not DetectNasIncreaseAfterShipEvent Then report = "Current quantity is not higher than the selected entity's balance after that shipment."
End Function

Private Sub AddFlag(ByRef flags As String, ByVal flagText As String)
    If flags <> "" Then flags = flags & ";"
    flags = flags & flagText
End Sub



Private Function BuildRecentShipmentSentRows(ByVal inventoryWb As Workbook, ByVal limitCount As Long, _
                                             Optional ByVal warehouseId As String = "", Optional ByVal runtimeRoot As String = "") As Collection
    Dim rows As Collection, data As Variant, row As Object, r As Long
    Set rows = New Collection
    data = modInventoryDomainBridge.ReadShipmentReconcileBridge("LIST" & vbTab & CStr(limitCount), inventoryWb, warehouseId, runtimeRoot)
    If Not IsEmpty(data) Then
        For r = LBound(data, 1) To UBound(data, 1)
            Set row = CreateObject("Scripting.Dictionary")
            row("EventID") = data(r, 1): row("System_Key") = data(r, 2): row("SKU") = data(r, 3)
            row("QtyDelta") = data(r, 4): row("OccurredAtUTC") = FormatLogDateReconcile(data(r, 5)): row("Location") = data(r, 6)
            rows.Add row
        Next r
    End If
    Set BuildRecentShipmentSentRows = rows
End Function



Public Function ShipmentLogDiagnosticsText(ByVal inventoryWb As Workbook) As String
    Dim loLog As ListObject
    Dim rowIndex As Long
    Dim eventType As String
    Dim shipCount As Long
    Dim reserveCount As Long
    Dim deductionCount As Long
    Dim pathText As String
    Dim warehouseId As String
    Dim stationId As String
    Dim pendingCount As Long
    Dim matchingPending As Long
    Dim inboxReport As String
    Dim inboxError As String
    Dim stagedCount As Long
    Dim matchingStaged As Long
    Dim stagedReport As String
    Dim stagedError As String
    Dim target As WarehouseTarget

    If inventoryWb Is Nothing Then
        ShipmentLogDiagnosticsText = "Inventory workbook: not open."
        Exit Function
    End If

    pathText = inventoryWb.FullName
    If pathText = "" Then pathText = inventoryWb.Name
    Set loLog = FindListObjectByNameReconcile(inventoryWb, "tblInventoryLog")
    If loLog Is Nothing Then
        ShipmentLogDiagnosticsText = "Inventory workbook: " & pathText & vbCrLf & _
                                     "tblInventoryLog: missing."
        Exit Function
    End If

    If Not loLog.DataBodyRange Is Nothing Then
        For rowIndex = 1 To loLog.ListRows.Count
            eventType = UCase$(SafeTrimReconcile(GetCellByColumnReconcile(loLog, rowIndex, "EventType")))
            If eventType = SERVER_SHIP_EVENT_TYPE Then shipCount = shipCount + 1
            If eventType = SERVER_SHIP_RESERVE_EVENT_TYPE Then reserveCount = reserveCount + 1
            If IsShipmentDeductionEventType(eventType) And NzDblReconcile(GetCellByColumnReconcile(loLog, rowIndex, "QtyDelta")) < 0 Then
                deductionCount = deductionCount + 1
            End If
        Next rowIndex
    End If

    Set target = modNasConnection.GetCurrentTarget()
    If Not target Is Nothing Then
        warehouseId = Trim$(target.WarehouseId)
        stationId = Trim$(target.StationId)
    End If
    If warehouseId = "" Then warehouseId = Trim$(modConfig.GetWarehouseId())
    If stationId = "" Then stationId = Trim$(modConfig.GetStationId())
    If warehouseId = "" Then warehouseId = Trim$(modConfig.GetString("WarehouseId", ""))
    If stationId = "" Then stationId = Trim$(modConfig.GetString("StationId", ""))
    inboxReport = modRoleEventWriter.DescribeInboxPendingRows(SERVER_SHIP_EVENT_TYPE, warehouseId, stationId, "", pendingCount, matchingPending, inboxError)
    stagedReport = modRoleEventWriter.DescribeLocalStagedInboxRows(SERVER_SHIP_EVENT_TYPE & "," & SERVER_SHIP_RESERVE_EVENT_TYPE & "," & SERVER_SHIP_RELEASE_EVENT_TYPE, _
                                                                    warehouseId, _
                                                                    stationId, _
                                                                    stagedCount, _
                                                                    matchingStaged, _
                                                                    stagedError)

    ShipmentLogDiagnosticsText = "Inventory workbook: " & pathText & vbCrLf & _
                                 "tblInventoryLog rows: " & CStr(loLog.ListRows.Count) & vbCrLf & _
                                 "SHIP rows: " & CStr(shipCount) & vbCrLf & _
                                 "SHIP_RESERVE rows: " & CStr(reserveCount) & vbCrLf & _
                                 "shipment deduction rows: " & CStr(deductionCount) & vbCrLf & _
                                 "NAS shipping inbox pending: " & CStr(pendingCount) & vbCrLf & _
                                 "NAS shipping inbox: " & IIf(inboxReport <> "", inboxReport, inboxError) & vbCrLf & _
                                 "local staged shipment rows: " & CStr(matchingStaged) & vbCrLf & _
                                 "local staging: " & IIf(stagedReport <> "", stagedReport, stagedError)
End Function

Private Function IsShipmentDeductionEventType(ByVal eventType As String) As Boolean
    eventType = UCase$(Trim$(eventType))
    IsShipmentDeductionEventType = (eventType = SERVER_SHIP_EVENT_TYPE)
End Function

Private Function FormatLogDateReconcile(ByVal valueIn As Variant) As String
    If IsDate(valueIn) Then
        FormatLogDateReconcile = Format$(CDate(valueIn), "yyyy-mm-dd hh:nn:ss")
    Else
        FormatLogDateReconcile = SafeTrimReconcile(valueIn)
    End If
End Function







Private Function FindListObjectByNameReconcile(ByVal wb As Workbook, ByVal tableName As String) As ListObject
    Dim ws As Worksheet

    If wb Is Nothing Then Exit Function
    On Error Resume Next
    For Each ws In wb.Worksheets
        Set FindListObjectByNameReconcile = ws.ListObjects(tableName)
        If Not FindListObjectByNameReconcile Is Nothing Then Exit Function
    Next ws
    On Error GoTo 0
End Function

Private Function GetCellByColumnReconcile(ByVal lo As ListObject, ByVal rowIndex As Long, ByVal columnName As String) As Variant
    Dim idx As Long

    idx = GetColumnIndexReconcile(lo, columnName)
    If idx = 0 Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    GetCellByColumnReconcile = lo.DataBodyRange.Cells(rowIndex, idx).Value
End Function

Private Function GetColumnIndexReconcile(ByVal lo As ListObject, ByVal columnName As String) As Long
    Dim i As Long

    If lo Is Nothing Then Exit Function
    For i = 1 To lo.ListColumns.Count
        If StrComp(lo.ListColumns(i).Name, columnName, vbTextCompare) = 0 Then
            GetColumnIndexReconcile = i
            Exit Function
        End If
    Next i
End Function

Private Function SafeTrimReconcile(ByVal valueIn As Variant) As String
    On Error Resume Next
    SafeTrimReconcile = Trim$(CStr(valueIn))
End Function

Private Function NzDblReconcile(ByVal valueIn As Variant) As Double
    On Error Resume Next
    If IsNumeric(valueIn) Then NzDblReconcile = CDbl(valueIn)
End Function
