Attribute VB_Name = "modTesterPresentation"
Option Explicit

Public Sub ShowTesterSetup(ByVal targetWb As Workbook)
    Dim report As String, context As String
    Dim observer As cProvisioningActivity, number As Long, source As String, description As String
    On Error GoTo Failed
    context = modTesterCommands.CaptureForm(targetWb.Name)
    Set observer = New cProvisioningActivity
    observer.Initialize targetWb, "ADMIN_TESTER_", "Test Environment Setup"
    Call modRoleWorkbookSurfaces.EnsureAdminLegacyWorkbookSurface(targetWb, report)
    If Not modLocalAddinsRegistration.EnsureLocalInvSysAddinsRegistered("", report) Then
        observer.OpeningFailed
        MsgBox "Current invSys add-ins are not registered cleanly for this Excel session." & vbCrLf & vbCrLf & _
               report, vbExclamation, "invSys Admin"
        modTesterCommands.ReleaseForm context
        Exit Sub
    End If
    frmSetupTesterStation.BindSource context
    frmSetupTesterStation.ObserveEntry observer
    frmSetupTesterStation.Show
    Exit Sub
Failed:
    number = Err.Number: source = Err.Source: description = Err.Description
    If Not observer Is Nothing Then observer.OpeningFailed
    modTesterCommands.ReleaseForm context
    Err.Raise number, source, description
End Sub
