Attribute VB_Name = "modTesterPresentation"
Option Explicit

Public Sub ShowTesterSetup(ByVal targetWb As Workbook)
    Dim report As String, context As String
    context = modTesterCommands.CaptureForm(targetWb.Name)
    Call modRoleWorkbookSurfaces.EnsureAdminLegacyWorkbookSurface(targetWb, report)
    If Not modLocalAddinsRegistration.EnsureLocalInvSysAddinsRegistered("", report) Then
        MsgBox "Current invSys add-ins are not registered cleanly for this Excel session." & vbCrLf & vbCrLf & _
               report, vbExclamation, "invSys Admin"
        modTesterCommands.ReleaseForm context
        Exit Sub
    End If
    frmSetupTesterStation.BindSource context
    frmSetupTesterStation.Show
End Sub
