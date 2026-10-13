Attribute VB_Name = "modAdminDirectoryActivity"
Option Explicit

' The existing callback performs preparation and its optional root prompt first.
' Freeze discovery before optional recording can open configuration workbooks.
Public Sub OpenFromRibbon(ByVal target As Workbook)
    Dim scope As Object, books As Collection, wb As Workbook, surface As Object
    Dim action As String, notice As String, report As String, opened As Boolean
    Dim number As Long, source As String, description As String
    On Error GoTo Failed
    Set surface = modAdminActionObservation.CurrentSurface()
    Set scope = CreateObject("Scripting.Dictionary"): Set books = New Collection
    For Each wb In Application.Workbooks: books.Add wb: Next wb
    scope.Add "Books", books
    scope.Add "Roots", modAdminConsole.GetWarehouseDirectoryScanRoots()
    action = modAdminActionObservation.BeginObservation("ADMIN_DIRECTORY_OPEN", notice, _
        "Tracking unavailable: warehouse directory opening could not be recorded.")
    modAdminActionObservation.RestoreSurface surface
    opened = modAdminConsole.OpenWarehouseDirectory(target, report, scope)
    Set surface = modAdminActionObservation.CurrentSurface()
    If opened Then
        modAdminActionObservation.FinishObservation action, "OPENED", notice, _
            "Tracking unavailable: the warehouse directory opening result could not be recorded."
    Else
        modAdminActionObservation.FinishObservation action, "FAILED", notice, _
            "Tracking unavailable: the warehouse directory opening result could not be recorded."
    End If
    modAdminActionObservation.RestoreSurface surface
    If opened Then
        MsgBox "Warehouse directory refreshed.", vbInformation, "invSys Admin"
    Else
        If Len(Trim$(report)) = 0 Then report = "Warehouse directory could not be opened."
        MsgBox report, vbExclamation, "invSys Admin"
    End If
    If notice <> "" Then MsgBox notice, vbExclamation, "invSys Admin"
    Exit Sub
Failed:
    number = Err.Number: source = Err.Source: description = Err.Description
    Set surface = modAdminActionObservation.CurrentSurface()
    modAdminActionObservation.FinishObservation action, "FAILED", notice, _
        "Tracking unavailable: the warehouse directory opening result could not be recorded."
    modAdminActionObservation.RestoreSurface surface
    If notice <> "" Then MsgBox notice, vbExclamation, "invSys Admin"
    Err.Raise number, source, description
End Sub

' Preserve the compatibility entry's original uncaptured discovery by default.
Public Function OpenView(ByVal adminWb As Workbook, ByRef report As String, _
                         ByVal readScope As Object) As Boolean
    Dim wb As Workbook
    Set wb = modAdminWorkbookTarget.ResolveAdminTargetWorkbook(adminWb, ThisWorkbook, True)
    If wb Is Nothing Then
        report = "Admin workbook not resolved."
        Exit Function
    End If
    If Not modAdminConsole.EnsureAdminSchema(wb, report) Then Exit Function
    If Not modAdminConsole.RefreshWarehouseDirectory(wb, report, readScope) Then Exit Function
    wb.Worksheets("WarehouseDirectory").Activate
    OpenView = True
End Function
