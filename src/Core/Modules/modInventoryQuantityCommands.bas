Attribute VB_Name = "modInventoryQuantityCommands"
Option Explicit

' Primitive command boundary; an empty context selects the explicit provisioned
' headless path, whose user is independently authorized by the queue service.
Public Function SetSingleEntityQuantity(ByVal warehouse As String, ByVal station As String, _
        ByVal user As String, ByVal sku As String, ByVal target As Double, _
        ByVal reason As String, ByRef report As String, Optional ByVal context As String = "") As Boolean
    Dim entity As Variant, current As Variant, values As Variant, delta As Double
    Dim items As Collection, item As Object, eventId As String, queueError As String
    Dim batchReport As String, applied As Boolean, processed As Long
    On Error GoTo Failed
    If target < 0 Or Trim$(reason) = "" Then
        report = "A nonnegative target quantity and edit reason are required.": Exit Function
    End If
    If context <> "" Then
        values = modInventoryCommands.SourceValues(context)
        If Not IsArray(values) Then GoTo Refused
        If values(0) <> warehouse Or values(1) <> station Or values(2) <> user Then GoTo Refused
    End If
    entity = ResolveSingleEntity(warehouse, sku, report)
    If Not IsArray(entity) Then Exit Function
    delta = target - CDbl(entity(1))
    If Abs(delta) < 0.0000001 Then
        report = "Inventory quantity already matched target.": SetSingleEntityQuantity = True: Exit Function
    End If
    If context <> "" Then
        If Not modInventoryCommands.CheckForm(context, report) Then Exit Function
    End If
    current = ResolveSingleEntity(warehouse, sku, report)
    If Not IsArray(current) Then Exit Function
    If Join(current, vbNullChar) <> Join(entity, vbNullChar) Then
        report = "The selected inventory entity changed. Reopen Add/Edit Inventory.": Exit Function
    End If
    Set items = New Collection
    Set item = modRoleEventWriter.CreateInventoryEntityPayloadItem(CStr(entity(0)), sku, _
        delta, CStr(entity(2)), CStr(entity(3)), CStr(entity(4)), "Admin set inventory quantity. Reason: " & reason)
    item("IoType") = "ADJUST": item("PreviousQty") = CDbl(entity(1))
    item("TargetQty") = target: item("Reason") = reason
    items.Add item
    If Not modRoleEventWriter.QueueAdminInventoryAdjustEvent(warehouse, station, user, _
            modRoleEventWriter.BuildPayloadJsonFromCollection(items), _
            "Admin set inventory quantity. Reason: " & reason, 0, eventId, queueError, "") Then
        report = "Inventory adjustment could not be queued: " & queueError: Exit Function
    End If
    If context <> "" Then
        If Not modInventoryCommands.CheckForm(context, report) Then
            report = report & vbCrLf & "Queued event: " & eventId & ". Application is unconfirmed."
            Exit Function
        End If
    End If
    processed = modProcessor.RunBatch(warehouse, 0, batchReport, eventId, applied)
    If Not applied Then
        report = "Inventory adjustment queued; exact saved application is unconfirmed." & _
                 vbCrLf & "Event: " & eventId & vbCrLf & batchReport
        Exit Function
    End If
    report = "Inventory quantity set. Exact event application was saved." & vbCrLf & _
             "Event: " & eventId & vbCrLf & "System_Key: " & CStr(entity(0)) & _
             vbCrLf & "Target quantity: " & CStr(target)
    SetSingleEntityQuantity = True
    Exit Function
Refused:
    report = "The invSys session or inventory source changed. Reopen Add/Edit Inventory."
    Exit Function
Failed:
    report = "Inventory quantity command failed."
    If eventId <> "" Then report = report & vbCrLf & "Event: " & eventId & ". Application is unconfirmed."
End Function

' Read only. A SKU is a catalog choice; a quantity command needs one exact active
' entity. Refuse ambiguity rather than allocate across physical entities.
Private Function ResolveSingleEntity(ByVal warehouse As String, ByVal sku As String, ByRef report As String) As Variant
    Dim source As cSnapshotInventorySource, book As Workbook, table As ListObject
    Dim fields As Object, row As Long, matches As Long, key As String, state As String
    Dim keys As Object, result As Variant, quantity As Variant
    On Error GoTo Invalid
    Set source = New cSnapshotInventorySource
    Set book = source.Read(warehouse, Nothing)
    If book Is Nothing Then GoTo Invalid
    Set table = modExecutionTarget.Table(book, "tblInventoryEntities")
    Set fields = modAggregationSources.RequiredColumns(table, _
        Array("System_Key", "SKU", "QtyOnHand", "Location", "Condition", "AttributesJson", "InventoryState"))
    If fields Is Nothing Then GoTo Invalid
    Set keys = CreateObject("Scripting.Dictionary"): keys.CompareMode = vbBinaryCompare
    For row = 1 To table.ListRows.Count
        key = CStr(table.DataBodyRange.Cells(row, fields("System_Key")).Value2)
        If key = "" Or keys.Exists(key) Then GoTo Invalid
        keys.Add key, True
        state = UCase$(Trim$(CStr(table.DataBodyRange.Cells(row, fields("InventoryState")).Value2)))
        If StrComp(CStr(table.DataBodyRange.Cells(row, fields("SKU")).Value2), sku, vbTextCompare) = 0 _
                And (state = "ACTIVE" Or state = "") Then
            matches = matches + 1
            If matches > 1 Then
                report = "Quantity target is ambiguous across active System_Key entities. No adjustment was queued."
                Exit Function
            End If
            quantity = table.DataBodyRange.Cells(row, fields("QtyOnHand")).Value2
            If Not IsNumeric(quantity) Then GoTo Invalid
            result = Array(key, CStr(CDbl(quantity)), _
                CStr(table.DataBodyRange.Cells(row, fields("Location")).Value2), _
                CStr(table.DataBodyRange.Cells(row, fields("Condition")).Value2), _
                CStr(table.DataBodyRange.Cells(row, fields("AttributesJson")).Value2))
        End If
    Next row
    If matches <> 1 Then GoTo Invalid
    ResolveSingleEntity = result
    Exit Function
Invalid:
    report = "A valid exact active inventory entity could not be resolved. No adjustment was queued."
End Function
