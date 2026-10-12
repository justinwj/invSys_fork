Attribute VB_Name = "modInventoryBridgeApi"
Option Explicit

Public Function InventoryEventAppliedBridgeResult(ByVal eventId As String, ByVal inventoryWb As Workbook) As Boolean
    InventoryEventAppliedBridgeResult = modInventoryEventReceipt.IsApplied(eventId, inventoryWb)
End Function

Public Function ReadShipmentReconcileBridgeResult(ByVal request As String, ByVal inventoryWb As Workbook) As Variant
    ReadShipmentReconcileBridgeResult = modShipmentReconcile.ReadShipmentReconcile(request, inventoryWb)
End Function

Public Function ResolveInventoryWorkbookBridgeResult(Optional ByVal warehouseId As String = "", _
                                                    Optional ByVal inventoryWb As Workbook = Nothing) As Workbook
    Set ResolveInventoryWorkbookBridgeResult = modInventoryApply.ResolveInventoryWorkbook(warehouseId, inventoryWb)
End Function

Public Function EnsureInventorySchemaBridgeResult(Optional ByVal targetWb As Workbook = Nothing) As Object
    Dim result As Object
    Dim report As String

    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare
    result("Success") = modInventorySchema.EnsureInventorySchema(targetWb, report)
    result("Report") = report
    Set EnsureInventorySchemaBridgeResult = result
End Function

Public Function EnsureInventorySchemaBridgeSuccess(Optional ByVal targetWb As Workbook = Nothing) As Boolean
    Dim report As String

    EnsureInventorySchemaBridgeSuccess = modInventorySchema.EnsureInventorySchema(targetWb, report)
End Function

Public Function EnsureInventorySchemaBridgeReport(Optional ByVal targetWb As Workbook = Nothing) As String
    Dim report As String

    Call modInventorySchema.EnsureInventorySchema(targetWb, report)
    EnsureInventorySchemaBridgeReport = report
End Function

Public Function ApplyEventBridgeResult(ByVal evt As Object, _
                                      Optional ByVal inventoryWb As Workbook = Nothing, _
                                      Optional ByVal runId As String = "") As Object
    Dim result As Object
    Dim statusOut As String
    Dim errorCode As String
    Dim errorMessage As String

    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare
    result("Success") = modInventoryApply.ApplyEvent(evt, inventoryWb, runId, statusOut, errorCode, errorMessage)
    result("StatusOut") = statusOut
    result("ErrorCode") = errorCode
    result("ErrorMessage") = errorMessage
    Set ApplyEventBridgeResult = result
End Function

Public Function ApplyEventBridgeEncoded(ByVal evt As Object, _
                                        Optional ByVal inventoryWb As Workbook = Nothing, _
                                        Optional ByVal runId As String = "", _
                                        Optional ByVal deferSave As Boolean = False) As String
    Dim statusOut As String
    Dim errorCode As String
    Dim errorMessage As String
    Dim success As Boolean

    success = modInventoryApply.ApplyEvent(evt, inventoryWb, runId, statusOut, errorCode, errorMessage, deferSave)
    ApplyEventBridgeEncoded = CStr(Abs(CLng(success))) & vbTab & statusOut & vbTab & errorCode & vbTab & errorMessage
End Function

Public Function ApplyEventBridgeEncodedDeferred(ByVal evt As Object, _
                                                Optional ByVal inventoryWb As Workbook = Nothing, _
                                                Optional ByVal runId As String = "") As String
    ApplyEventBridgeEncodedDeferred = ApplyEventBridgeEncoded(evt, inventoryWb, runId, True)
End Function

Public Function RemoveLastBulkLogEntriesBridgeResult(ByVal countToRemove As Long) As Collection
    ' Compatibility entry point only. Inventory history is append-only; undo
    ' must be represented by a compensating event, never by deleting log rows.
    Set RemoveLastBulkLogEntriesBridgeResult = New Collection
End Function

Public Sub ReAddBulkLogEntriesBridgeResult(ByVal logDataCollection As Collection)
    ' Compatibility entry point only. Replaying deleted rows would bypass
    ' ApplyEvent validation, idempotency, locking, and projection rebuilding.
End Sub

Public Sub ScheduleSourceWorkbookSyncBridgeResult()
    On Error Resume Next
    Application.Run "'" & ThisWorkbook.Name & "'!modInventoryInit.ScheduleSourceWorkbookSync"
    On Error GoTo 0
End Sub

Public Function PublishInventorySnapshotBridgeEncoded(Optional ByVal targetWb As Workbook = Nothing) As String
    Dim report As String
    Dim success As Boolean

    success = modInventoryPublisher.EnsureSnapshotPublicationForWorkbook(targetWb, report)
    PublishInventorySnapshotBridgeEncoded = CStr(Abs(CLng(success))) & vbTab & report
End Function

Public Function RebuildInventoryProjectionsBridgeEncoded(ByVal inventoryWb As Workbook) As String
    Dim report As String
    Dim success As Boolean

    success = modInventoryApply.RebuildInventoryProjectionsForWorkbook(inventoryWb, report)
    RebuildInventoryProjectionsBridgeEncoded = CStr(Abs(CLng(success))) & vbTab & report
End Function

Public Function GetOnHandQtyBridgeResult(ByVal sku As String, _
                                         Optional ByVal inventoryWb As Workbook = Nothing) As Double
    GetOnHandQtyBridgeResult = modInventoryQueries.GetOnHandQty(sku, inventoryWb)
End Function

Public Function GetLocationBalancesBridgeResult(ByVal sku As String, _
                                                Optional ByVal inventoryWb As Workbook = Nothing) As Variant
    GetLocationBalancesBridgeResult = modInventoryQueries.GetLocationBalances(sku, inventoryWb)
End Function

Public Function ListInventoryPickerItemsBridgeResult(Optional ByVal filterText As String = "", _
                                                     Optional ByVal inventoryWb As Workbook = Nothing) As Variant
    ListInventoryPickerItemsBridgeResult = modInventoryQueries.ListInventoryPickerItems(filterText, inventoryWb)
End Function

Public Function ListAvailableInventoryEntitiesBridgeResult(Optional ByVal filterText As String = "", _
                                                           Optional ByVal inventoryWb As Workbook = Nothing) As Variant
    ListAvailableInventoryEntitiesBridgeResult = _
        modInventoryQueries.ListAvailableInventoryEntities(filterText, inventoryWb)
End Function

Public Function DiagnoseInventoryDomainBridgeResult() As String
    DiagnoseInventoryDomainBridgeResult = modInventoryInit.DiagnoseInventoryDomain()
End Function
