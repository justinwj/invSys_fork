Attribute VB_Name = "modWarehouseDirectoryRows"
Option Explicit
Option Private Module

' The directory is a local projection. Unknown cells remain at their positions.
Public Sub ClearManaged(ByVal table As ListObject)
    Dim column As ListColumn
    For Each column In table.ListColumns
        Select Case LCase$(Trim$(column.Name))
            Case "warehouseid", "stationid", "warehousename", "runtimeroot", _
                 "configworkbook", "authworkbook", "inventoryworkbook", _
                 "receiveinbox", "shipinbox", "prodinbox", _
                 "processorserviceuserid", "pathsharepointroot", "roledefault", "status"
                If Not column.DataBodyRange Is Nothing Then column.DataBodyRange.ClearContents
        End Select
    Next column
End Sub

Public Function AtIndex(ByVal table As ListObject, ByVal index As Long) As ListRow
    Do While table.ListRows.Count < index
        table.ListRows.Add
    Loop
    Set AtIndex = table.ListRows(index)
End Function
