Attribute VB_Name = "modShipmentReconcile"
Option Explicit

' Read-only result contract: LIST returns rows of event/key/SKU/delta/time/location;
' STATE returns current/after/SKU/location/condition/attributes/inventory-state.
' Empty means the supplied authority or requested identity/history is invalid.
Public Function ReadShipmentReconcile(ByVal request As String, ByVal wb As Workbook) As Variant
    On Error GoTo InvalidRead
    Dim parts As Variant
    parts = Split(request, vbTab)
    If wb Is Nothing Then Exit Function
    If UBound(parts) = 1 Then
        If parts(0) = "LIST" Then ReadShipmentReconcile = RecentShipments(wb, CLng(parts(1)))
    ElseIf UBound(parts) = 2 Then
        If parts(0) = "STATE" Then ReadShipmentReconcile = EntityShipmentState(wb, CStr(parts(1)), CStr(parts(2)))
    End If
    Exit Function
InvalidRead:
    ReadShipmentReconcile = Empty
End Function

Public Function ValidateCorrection(ByVal item As Object, ByVal wb As Workbook, _
                                   ByRef errorCode As String, ByRef report As String) As Boolean
    On Error GoTo InvalidCorrection
    Dim state As Variant
    If Len(Trim$(CStr(item("RepairNarrative")))) = 0 Then GoTo InvalidCorrection
    If Not IsNumeric(item("Qty")) Then GoTo InvalidCorrection
    If CDbl(item("Qty")) >= 0 Then GoTo InvalidCorrection
    state = EntityShipmentState(wb, CStr(item("CorrectedShipEventId")), CStr(item("System_Key")))
    If IsEmpty(state) Then GoTo InvalidCorrection
    If StrComp(CStr(item("SKU")), CStr(state(2)), vbTextCompare) <> 0 Then GoTo InvalidCorrection
    ' A quantity correction cannot restore stale location/condition/custom state.
    item("Location") = state(3)
    item("Condition") = state(4)
    item("AttributesJson") = state(5)
    item("InventoryState") = state(6)
    ValidateCorrection = True
    Exit Function
InvalidCorrection:
    errorCode = "INVALID_PAYLOAD"
    report = "Shipment correction requires a negative quantity, a human repair narrative, and an exact existing SHIP event/entity reference."
End Function

Private Function RecentShipments(ByVal wb As Workbook, ByVal limitCount As Long) As Variant
    Dim lo As ListObject, rows As Collection, row As Variant, result() As Variant
    Dim r As Long, i As Long, c As Long, eventId As String, systemKey As String
    If limitCount <= 0 Then Exit Function
    Set lo = RequiredTable(wb, "tblInventoryLog", "EventID|System_Key|SKU|QtyDelta|EventType|AppliedSeq|OccurredAtUTC|Location")
    Set rows = New Collection
    For r = lo.ListRows.Count To 1 Step -1
        If UCase$(Trim$(CStr(CellValue(lo, r, "EventType")))) = "SHIP" Then
            If Number(CellValue(lo, r, "QtyDelta")) < 0 Then
                eventId = CStr(CellValue(lo, r, "EventID"))
                systemKey = CStr(CellValue(lo, r, "System_Key"))
                If Len(eventId) = 0 Or Len(systemKey) = 0 Then Err.Raise 5
                If Sequence(CellValue(lo, r, "AppliedSeq")) <= 0 Then Err.Raise 5
                rows.Add Array(eventId, systemKey, CStr(CellValue(lo, r, "SKU")), _
                    Number(CellValue(lo, r, "QtyDelta")), CellValue(lo, r, "OccurredAtUTC"), CStr(CellValue(lo, r, "Location")))
                If rows.Count = limitCount Then Exit For
            End If
        End If
    Next r
    If rows.Count = 0 Then Exit Function
    ReDim result(1 To rows.Count, 1 To 6)
    For i = 1 To rows.Count
        row = rows(i)
        For c = 0 To 5
            result(i, c + 1) = row(c)
        Next c
    Next i
    RecentShipments = result
End Function

Private Function EntityShipmentState(ByVal wb As Workbook, ByVal eventId As String, ByVal systemKey As String) As Variant
    Dim entities As ListObject, history As ListObject, applied As ListObject
    Dim entityRow As Long, selectedSeq As Long, ledgerMatches As Long, r As Long, seq As Long
    Dim selectedKeyFound As Boolean, afterQty As Double, delta As Double
    If Len(eventId) = 0 Or Len(systemKey) = 0 Then Exit Function
    Set entities = RequiredTable(wb, "tblInventoryEntities", "System_Key|SKU|QtyOnHand|Location|Condition|AttributesJson|InventoryState")
    Set history = RequiredTable(wb, "tblInventoryLog", "EventID|System_Key|EventType|AppliedSeq|QtyDelta")
    Set applied = RequiredTable(wb, "tblAppliedEvents", "EventID|AppliedSeq")
    For r = 1 To entities.ListRows.Count
        If StrComp(CStr(CellValue(entities, r, "System_Key")), systemKey, vbBinaryCompare) = 0 Then
            If entityRow <> 0 Then Exit Function
            entityRow = r
        End If
    Next r
    If entityRow = 0 Then Exit Function
    For r = 1 To history.ListRows.Count
        If StrComp(CStr(CellValue(history, r, "EventID")), eventId, vbBinaryCompare) = 0 Then
            If UCase$(Trim$(CStr(CellValue(history, r, "EventType")))) <> "SHIP" Then Exit Function
            seq = Sequence(CellValue(history, r, "AppliedSeq"))
            If selectedSeq <> 0 And selectedSeq <> seq Then Exit Function
            selectedSeq = seq
            If StrComp(CStr(CellValue(history, r, "System_Key")), systemKey, vbBinaryCompare) = 0 Then
                If Number(CellValue(history, r, "QtyDelta")) >= 0 Then Exit Function
                selectedKeyFound = True
            End If
        End If
    Next r
    If Not selectedKeyFound Or selectedSeq = 0 Then Exit Function
    For r = 1 To applied.ListRows.Count
        If StrComp(CStr(CellValue(applied, r, "EventID")), eventId, vbBinaryCompare) = 0 Then
            If Sequence(CellValue(applied, r, "AppliedSeq")) <> selectedSeq Then Exit Function
            ledgerMatches = ledgerMatches + 1
        End If
    Next r
    If ledgerMatches <> 1 Then Exit Function
    For r = 1 To history.ListRows.Count
        If StrComp(CStr(CellValue(history, r, "System_Key")), systemKey, vbBinaryCompare) = 0 Then
            seq = Sequence(CellValue(history, r, "AppliedSeq"))
            delta = Number(CellValue(history, r, "QtyDelta"))
            If seq = selectedSeq And StrComp(CStr(CellValue(history, r, "EventID")), eventId, vbBinaryCompare) <> 0 Then Exit Function
            If seq <= selectedSeq Then afterQty = afterQty + delta
        End If
    Next r
    EntityShipmentState = Array(Number(CellValue(entities, entityRow, "QtyOnHand")), afterQty, _
        CStr(CellValue(entities, entityRow, "SKU")), CStr(CellValue(entities, entityRow, "Location")), _
        CStr(CellValue(entities, entityRow, "Condition")), CStr(CellValue(entities, entityRow, "AttributesJson")), _
        CStr(CellValue(entities, entityRow, "InventoryState")))
End Function

Private Function RequiredTable(ByVal wb As Workbook, ByVal name As String, ByVal headers As String) As ListObject
    Dim ws As Worksheet, lo As ListObject, found As ListObject, header As Variant, index As Long
    For Each ws In wb.Worksheets
        For Each lo In ws.ListObjects
            If StrComp(lo.Name, name, vbTextCompare) = 0 Then
                If Not found Is Nothing Then Err.Raise 5
                Set found = lo
            End If
        Next lo
    Next ws
    If found Is Nothing Then Err.Raise 5
    For Each header In Split(headers, "|")
        index = ColumnIndex(found, CStr(header))
    Next header
    Set RequiredTable = found
End Function

Private Function ColumnIndex(ByVal lo As ListObject, ByVal name As String) As Long
    Dim column As ListColumn
    For Each column In lo.ListColumns
        If StrComp(Trim$(column.Name), name, vbTextCompare) = 0 Then
            If ColumnIndex <> 0 Then Err.Raise 5
            ColumnIndex = column.Index
        End If
    Next column
    If ColumnIndex = 0 Then Err.Raise 5
End Function

Private Function CellValue(ByVal lo As ListObject, ByVal row As Long, ByVal header As String) As Variant
    CellValue = lo.DataBodyRange.Cells(row, ColumnIndex(lo, header)).Value2
End Function

Private Function Number(ByVal value As Variant) As Double
    If IsError(value) Or IsNull(value) Or IsEmpty(value) Then Err.Raise 5
    If Not IsNumeric(value) Then Err.Raise 5
    Number = CDbl(value)
End Function

Private Function Sequence(ByVal value As Variant) As Long
    Dim numeric As Double
    numeric = Number(value)
    If numeric <= 0 Or numeric <> Fix(numeric) Then Err.Raise 5
    Sequence = CLng(numeric)
End Function
