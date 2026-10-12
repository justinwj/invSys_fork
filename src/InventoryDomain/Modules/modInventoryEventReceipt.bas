Attribute VB_Name = "modInventoryEventReceipt"
Option Explicit

' A saved Domain receipt is evidence for this event, not for a whole batch.
Public Function IsApplied(ByVal eventId As String, ByVal wb As Workbook) As Boolean
    On Error GoTo Unconfirmed
    Dim ws As Worksheet, lo As ListObject, receipts As ListObject
    Dim idColumn As Long, statusColumn As Long, row As Long, matches As Long
    Dim applied As Boolean
    If wb Is Nothing Or Len(Trim$(eventId)) = 0 Then Exit Function
    If Not wb.Saved Then Exit Function
    For Each ws In wb.Worksheets
        For Each lo In ws.ListObjects
            If StrComp(lo.Name, "tblAppliedEvents", vbTextCompare) = 0 Then Set receipts = lo
        Next lo
    Next ws
    If receipts Is Nothing Then Exit Function
    idColumn = HeaderIndex(receipts, "EVENTID"): statusColumn = HeaderIndex(receipts, "STATUS")
    If idColumn = 0 Or statusColumn = 0 Then Exit Function
    For row = 1 To receipts.ListRows.Count
        If StrComp(CStr(receipts.DataBodyRange.Cells(row, idColumn).Value2), eventId, vbBinaryCompare) = 0 Then
            matches = matches + 1
            applied = (CStr(receipts.DataBodyRange.Cells(row, statusColumn).Value2) = "APPLIED")
        End If
    Next row
    IsApplied = (matches = 1 And applied)
Unconfirmed:
End Function

Private Function HeaderIndex(ByVal lo As ListObject, ByVal normalizedName As String) As Long
    Dim column As ListColumn
    For Each column In lo.ListColumns
        If UCase$(Trim$(column.Name)) = normalizedName Then
            If HeaderIndex <> 0 Then HeaderIndex = 0: Exit Function
            HeaderIndex = column.Index
        End If
    Next column
End Function
