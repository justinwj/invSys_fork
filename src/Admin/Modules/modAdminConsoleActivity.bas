Attribute VB_Name = "modAdminConsoleActivity"
Option Explicit

Public Sub OpenFromRibbon()
    Dim target As Workbook, configPath As String, report As String, action As String
    Dim notice As String, outcome As String, number As Long, source As String, description As String
    Dim surface As Object
    On Error GoTo Failed
    Set target = modAdmin.ResolveInteractiveAdminWorkbook()
    configPath = modAdminConsole.CaptureConsoleConfigPath()
    Set surface = modAdminActionObservation.CurrentSurface()
    action = modAdminActionObservation.BeginObservation("ADMIN_CONSOLE_OPEN", notice, _
        "Tracking unavailable: Admin Console opening could not be recorded.")
    modAdminActionObservation.RestoreSurface surface
    Call modRoleWorkbookSurfaces.EnsureAdminLegacyWorkbookSurface(target, report)
    outcome = "FAILED"
    If modAdminConsole.OpenAdminConsole(target, report, configPath) Then outcome = "OPENED"
    Set surface = modAdminActionObservation.CurrentSurface()
    modAdminActionObservation.FinishObservation action, outcome, notice, _
        "Tracking unavailable: the Admin Console opening result could not be recorded."
    modAdminActionObservation.RestoreSurface surface
    If notice <> "" Then MsgBox notice, vbExclamation, "invSys Admin"
    Exit Sub
Failed:
    number = Err.Number: source = Err.Source: description = Err.Description
    Set surface = modAdminActionObservation.CurrentSurface()
    modAdminActionObservation.FinishObservation action, "FAILED", notice, _
        "Tracking unavailable: the Admin Console opening result could not be recorded."
    modAdminActionObservation.RestoreSurface surface
    If notice <> "" Then MsgBox notice, vbExclamation, "invSys Admin"
    Err.Raise number, source, description
End Sub



' Existing navigation owner, delegated here to keep the console module bounded.
' vbNullChar means uncaptured legacy discovery; an empty captured path stays empty.
Public Function OpenConsoleView(ByVal adminWb As Workbook, ByRef report As String, _
                                ByVal capturedConfigPath As String) As Boolean
    Dim wb As Workbook
    Set wb = modAdminWorkbookTarget.ResolveAdminTargetWorkbook(adminWb, ThisWorkbook, True)
    If wb Is Nothing Then
        report = "Admin workbook not resolved."
        Exit Function
    End If
    If Not modAdminConsole.EnsureAdminSchema(wb, report) Then Exit Function
    If Not modAdminConsole.RefreshAdminConsole(wb, report, capturedConfigPath) Then Exit Function
    wb.Worksheets("AdminConsole").Activate
    OpenConsoleView = True
End Function
