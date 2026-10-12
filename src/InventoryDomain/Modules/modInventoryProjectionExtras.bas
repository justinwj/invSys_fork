Attribute VB_Name = "modInventoryProjectionExtras"
Option Explicit

' Unknown projection columns stay attached to the immutable entity, not its row.
Public Function CaptureEntityExtras(ByVal lo As ListObject) As Object
    Dim saved As Object, values As Object, column As ListColumn, r As Long
    Dim keyColumn As Long, systemKey As String, header As String
    Set saved = CreateObject("Scripting.Dictionary")
    saved.CompareMode = vbBinaryCompare
    For Each column In lo.ListColumns
        If StrComp(Trim$(column.Name), "System_Key", vbTextCompare) = 0 Then keyColumn = column.Index
    Next column
    If keyColumn = 0 Then Err.Raise 5, , "Entity projection requires System_Key."
    For r = 1 To lo.ListRows.Count
        systemKey = CStr(lo.DataBodyRange.Cells(r, keyColumn).Value2)
        If Len(systemKey) > 0 Then
            If saved.Exists(systemKey) Then Err.Raise 5, , "Entity projection contains duplicate System_Key."
            Set values = CreateObject("Scripting.Dictionary")
            values.CompareMode = vbBinaryCompare
            For Each column In lo.ListColumns
                header = UCase$(Trim$(column.Name))
                Select Case header
                    Case "SYSTEM_KEY", "SKU", "QTYONHAND", "LOCATION", "CONDITION", "INVENTORYSTATE", "ATTRIBUTESJSON", "LASTAPPLIEDUTC"
                    Case Else
                        values.Add column.Name, lo.DataBodyRange.Cells(r, column.Index).Formula
                End Select
            Next column
            saved.Add systemKey, values
        End If
    Next r
    Set CaptureEntityExtras = saved
End Function

Public Sub RestoreEntityExtras(ByVal lo As ListObject, ByVal saved As Object)
    Dim column As ListColumn, r As Long, keyColumn As Long, systemKey As String, values As Object, header As String
    For Each column In lo.ListColumns
        If StrComp(Trim$(column.Name), "System_Key", vbTextCompare) = 0 Then keyColumn = column.Index
    Next column
    If keyColumn = 0 Then Err.Raise 5, , "Entity projection requires System_Key."
    For r = 1 To lo.ListRows.Count
        systemKey = CStr(lo.DataBodyRange.Cells(r, keyColumn).Value2)
        If saved.Exists(systemKey) Then
            Set values = saved(systemKey)
            For Each column In lo.ListColumns
                header = column.Name
                If values.Exists(header) Then lo.DataBodyRange.Cells(r, column.Index).Formula = values(header)
            Next column
        End If
    Next r
End Sub
