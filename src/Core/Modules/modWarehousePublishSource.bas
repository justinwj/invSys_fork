Attribute VB_Name = "modWarehousePublishSource"
Option Explicit
Option Private Module

' Read an existing local runtime without changing target, repairing or saving it.
Public Function ReadExisting(ByVal root As String, ByRef spec As WarehouseSpec) As Boolean
    Dim path As String, fileName As String, wb As Workbook, candidate As Workbook
    Dim owned As Boolean, wh As ListObject, stations As ListObject, row As Long
    Dim station As String, fso As Object
    On Error GoTo Finished
    root = modDeploymentPaths.NormalizeManagedFolderPath(root, False)
    Set fso = CreateObject("Scripting.FileSystemObject")
    If root = "" Or Not fso.FolderExists(root) Then Exit Function
    root = fso.GetAbsolutePathName(root)
    fileName = Dir$(root & "\*.invSys.Config.xlsb", vbNormal Or vbReadOnly Or vbHidden Or vbSystem)
    If fileName = "" Then Exit Function
    path = root & "\" & fileName
    If Dir$() <> "" Then Exit Function
    For Each candidate In Application.Workbooks
        If StrComp(candidate.FullName, path, vbTextCompare) = 0 Then Set wb = candidate: Exit For
    Next candidate
    If wb Is Nothing Then
        Set wb = Application.Workbooks.Open(Filename:=path, UpdateLinks:=0, ReadOnly:=True, _
            IgnoreReadOnlyRecommended:=True, Notify:=False, AddToMru:=False)
        owned = True
    End If
    If Not wb.Saved Then GoTo Finished
    Set wh = modExecutionTarget.Table(wb, "tblWarehouseConfig")
    Set stations = modExecutionTarget.Table(wb, "tblStationConfig")
    If wh Is Nothing Or stations Is Nothing Then GoTo Finished
    If wh.ListRows.Count <> 1 Then GoTo Finished
    spec.WarehouseId = FieldText(wh, 1, "WarehouseId")
    spec.WarehouseName = FieldText(wh, 1, "WarehouseName")
    If spec.WarehouseId = "" Then GoTo Finished
    If StrComp(fileName, spec.WarehouseId & ".invSys.Config.xlsb", vbBinaryCompare) <> 0 Then GoTo Finished
    If StrComp(modDeploymentPaths.NormalizeManagedFolderPath(FieldText(wh, 1, "PathDataRoot"), False), root, vbTextCompare) <> 0 Then GoTo Finished
    For row = 1 To stations.ListRows.Count
        If FieldText(stations, row, "WarehouseId") = spec.WarehouseId Then
            station = FieldText(stations, row, "StationId")
            If station <> "" Then spec.StationId = station: Exit For
        End If
    Next row
    If spec.StationId = "" Then GoTo Finished
    If Not fso.FileExists(root & "\" & spec.WarehouseId & ".invSys.Auth.xlsb") Then GoTo Finished
    If Not fso.FileExists(root & "\" & spec.WarehouseId & ".invSys.Data.Inventory.xlsb") Then GoTo Finished
    If Not fso.FolderExists(root & "\inbox") Then GoTo Finished
    spec.PathLocal = root
    ReadExisting = True
Finished:
    If owned And Not wb Is Nothing Then wb.Close SaveChanges:=False
End Function

Private Function FieldText(ByVal table As ListObject, ByVal row As Long, ByVal header As String) As String
    FieldText = Trim$(CStr(modTrackingPolicyModel.TableValue(table, row, header)))
End Function
