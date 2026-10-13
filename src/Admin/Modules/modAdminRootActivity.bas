Attribute VB_Name = "modAdminRootActivity"
Option Explicit
Option Private Module

Public Sub AddFromRibbon()
    Dim target As Workbook, wb As Workbook, surface As Object, scope As Object, inputs As Object
    Dim books As New Collection, action As String, notice As String, rootPath As String
    Dim report As String, remembered As Boolean, opened As Boolean, outcome As String
    Dim number As Long, source As String, description As String
    On Error GoTo Failed
    Set surface = modAdminActionObservation.CurrentSurface()
    Set target = modAdmin.ResolveInteractiveAdminWorkbook()
    For Each wb In Application.Workbooks: books.Add wb: Next wb
    Set inputs = CaptureInputs(notice)
    If Not inputs Is Nothing Then
        action = modAdminActionObservation.BeginObservation("ADMIN_WAREHOUSE_ROOT_ADD", notice, _
            "Tracking unavailable: Add Warehouse Root could not be recorded.")
    End If
    modAdminActionObservation.RestoreSurface surface
    rootPath = Trim$(InputBox("Enter a NAS/server warehouse hub folder or a specific warehouse runtime folder to include in Admin warehouse scans.", _
                             "invSys Admin - Warehouse Root", "\\100.84.136.19\invSysWH1"))
    outcome = "CANCELLED"
    If rootPath <> "" Then
        remembered = modAdminConsole.RememberWarehouseScanRoot(rootPath)
        Call modRoleWorkbookSurfaces.EnsureAdminLegacyWorkbookSurface(target, report)
        Set scope = CreateObject("Scripting.Dictionary")
        scope.Add "Books", books
        scope.Add "Inputs", inputs
        opened = modAdminConsole.OpenWarehouseDirectory(target, report, scope)
        If remembered Then outcome = "REMEMBERED" Else outcome = "UNCONFIRMED"
        If opened Then outcome = outcome & "_REFRESHED" Else outcome = outcome & "_REFRESH_FAILED"
    End If
    Set surface = modAdminActionObservation.CurrentSurface()
    modAdminActionObservation.FinishObservation action, outcome, notice, _
        "Tracking unavailable: the Add Warehouse Root result could not be recorded."
    modAdminActionObservation.RestoreSurface surface
    If rootPath <> "" Then
        If opened Then
            report = "Warehouse root remembered and directory refreshed."
        Else
            If Len(Trim$(report)) = 0 Then report = "Warehouse directory could not be opened."
        End If
        If Not remembered Then
            If opened Then report = "Warehouse directory refreshed."
            report = "Warehouse root preference could not be confirmed." & vbCrLf & report
        End If
        MsgBox report, IIf(remembered And opened, vbInformation, vbExclamation), "invSys Admin"
    End If
    If notice <> "" Then MsgBox notice, vbExclamation, "invSys Admin"
    Exit Sub
Failed:
    number = Err.Number: source = Err.Source: description = Err.Description
    Set surface = modAdminActionObservation.CurrentSurface()
    modAdminActionObservation.FinishObservation action, "FAILED", notice, _
        "Tracking unavailable: the Add Warehouse Root result could not be recorded."
    modAdminActionObservation.RestoreSurface surface
    If notice <> "" Then MsgBox notice, vbExclamation, "invSys Admin"
    Err.Raise number, source, description
End Sub

Private Function CaptureInputs(ByRef notice As String) As Object
    Dim context As String
    On Error GoTo Unavailable
    context = modActivity.CaptureContext()
    Set CaptureInputs = modAdminConsole.CaptureWarehouseDirectoryInputs()
    Exit Function
Unavailable:
    notice = "Tracking unavailable: Add Warehouse Root discovery context could not be captured."
    On Error Resume Next
    modActionRecording.ContextClosed context, True
End Function
