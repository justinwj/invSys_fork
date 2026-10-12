VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cProductionStagingReload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' One captured worksheet reload. Retained tables keep workbook-local columns.
Private mTables As Object
Private mSheet As Worksheet

Public Sub Prepare(ByVal sheet As Worksheet, ByVal chooserHeaders As Variant, ByVal paletteHeaders As Variant)
    Dim table As ListObject, headers As Variant, index As Long, name As String, address As String
    Set mSheet = sheet
    Set mTables = CreateObject("Scripting.Dictionary")
    mTables.CompareMode = vbTextCompare
    For index = sheet.ListObjects.Count To 1 Step -1
        Set table = sheet.ListObjects(index): name = LCase$(table.Name)
        If name = "recipechooser_generated" Or name Like "proc_*_rchooser" Then
            headers = chooserHeaders
        ElseIf name = "inventorypalette_generated" Or name Like "proc_*_palette" Then
            headers = paletteHeaders
        Else
            GoTo NextTable
        End If
        If HasCustomColumns(table, headers) Then
            RequireFields table, headers
            mTables.Add table.Name, table
            ClearFields table, headers, False
        Else
            address = table.Range.Address
            table.Delete
            sheet.Range(address).Clear
        End If
NextTable:
    Next index
End Sub

Public Function ReuseRange(ByVal proposed As Range, ByVal name As String, ByVal rowCount As Long, ByRef reused As Boolean) As Range
    Dim table As ListObject, target As Range, growth As Range, other As ListObject, rows As Long
    Set ReuseRange = proposed
    reused = False
    If Not mTables.Exists(name) Then Exit Function
    Set table = mTables(name)
    rows = table.Range.Rows.Count
    If rowCount + 1 > rows Then rows = rowCount + 1
    If table.ShowTotals Then Err.Raise 5, , "Reload does not support a totals row."
    Set target = table.HeaderRowRange.Resize(rows, table.ListColumns.Count)
    If rows > table.Range.Rows.Count Then
        Set growth = target.Rows(table.Range.Rows.Count + 1).Resize(rows - table.Range.Rows.Count)
        For Each other In mSheet.ListObjects
            If Not (other Is table) Then
                If Not Application.Intersect(other.Range, growth) Is Nothing Then Err.Raise 5, , "Reload growth overlaps another table."
            End If
        Next other
        If Application.WorksheetFunction.CountA(growth) <> 0 Then Err.Raise 5, , "Reload growth overlaps worksheet values."
    End If
    Set ReuseRange = target
    reused = True
End Function

Public Function WriteRows(ByVal area As Range, ByVal headers As Variant, ByRef values As Variant) As ListObject
    Dim key As Variant, table As ListObject, header As Variant, source As Long, row As Long, column As Long
    If Not (area.Parent Is mSheet) Then Err.Raise 5, , "Reload worksheet changed."
    For Each key In mTables.Keys
        Set table = mTables(key)
        If table.HeaderRowRange.Cells(1, 1).Address = area.Cells(1, 1).Address Then
            RequireFields table, headers
            table.Resize area
            source = LBound(values, 2)
            For Each header In headers
                column = Field(table, CStr(header))
                For row = LBound(values, 1) To UBound(values, 1)
                    table.DataBodyRange.Cells(row, column).Value = values(row, source)
                Next row
                source = source + 1
            Next header
            Set WriteRows = table
            Exit Function
        End If
    Next key
    Err.Raise 5, , "Retained reload table is unavailable."
End Function

Public Sub ClearOutputFormulas(ByVal table As ListObject)
    ClearFields table, Array("PROCESS", "OUTPUT", "UOM", "REAL OUTPUT", "BATCH", "RECALL CODE", "System_Key", "ITEM_CODE"), True
End Sub

Private Sub ClearFields(ByVal table As ListObject, ByVal headers As Variant, ByVal formulasOnly As Boolean)
    Dim header As Variant, column As Long, cell As Range
    If table.DataBodyRange Is Nothing Then Exit Sub
    For Each header In headers
        column = Field(table, CStr(header))
        If column > 0 Then
            If formulasOnly Then
                For Each cell In table.ListColumns(column).DataBodyRange.Cells
                    If cell.HasFormula Then cell.ClearContents
                Next cell
            Else
                table.ListColumns(column).DataBodyRange.ClearContents
            End If
        End If
    Next header
End Sub

Private Function HasCustomColumns(ByVal table As ListObject, ByVal headers As Variant) As Boolean
    Dim column As ListColumn, header As Variant, managed As Boolean
    For Each column In table.ListColumns
        managed = False
        For Each header In headers
            If StrComp(Trim$(column.Name), CStr(header), vbTextCompare) = 0 Then managed = True: Exit For
        Next header
        If Not managed Then HasCustomColumns = True: Exit Function
    Next column
End Function

Private Sub RequireFields(ByVal table As ListObject, ByVal headers As Variant)
    Dim header As Variant
    For Each header In headers
        If Field(table, CStr(header)) = 0 Then Err.Raise 5, , "Required reload column is unavailable."
    Next header
End Sub

Private Function Field(ByVal table As ListObject, ByVal header As String) As Long
    Dim column As ListColumn
    For Each column In table.ListColumns
        If StrComp(Trim$(column.Name), Trim$(header), vbTextCompare) = 0 Then
            If Field > 0 Then Err.Raise 5, , "Ambiguous reload column."
            Field = column.Index
        End If
    Next column
End Function

Public Sub ReadSummary(ByVal sheet As Worksheet, ByVal recipeId As String, _
                       ByRef name As String, ByRef description As String, ByRef department As String)
    Dim table As ListObject, row As Long
    name = "": description = "": department = ""
    If sheet Is Nothing Then Exit Sub
    On Error Resume Next
    Set table = sheet.ListObjects("Recipes")
    On Error GoTo 0
    If table Is Nothing Then Exit Sub
    If Field(table, "RECIPE_ID") = 0 Or Field(table, "RECIPE") = 0 Then Exit Sub
    For row = 1 To table.ListRows.Count
        If CellText(table, row, "RECIPE_ID") = recipeId Then
            name = CellText(table, row, "RECIPE")
            description = CellText(table, row, "DESCRIPTION")
            department = CellText(table, row, "DEPARTMENT")
            Exit Sub
        End If
    Next row
End Sub

Private Function CellText(ByVal table As ListObject, ByVal row As Long, ByVal header As String) As String
    Dim column As Long, value As Variant
    column = Field(table, header)
    If column = 0 Then Exit Function
    value = table.DataBodyRange.Cells(row, column).Value
    If IsError(value) Or IsNull(value) Or IsEmpty(value) Then Exit Function
    CellText = CStr(value)
End Function
