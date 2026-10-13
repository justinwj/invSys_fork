Attribute VB_Name = "modAdminCreateEntry"
Option Explicit
Option Private Module

Public Sub OpenFromRibbon()
    Dim report As String, target As Workbook, observer As cCreateWarehouseActivity
    Dim writeContext As String, number As Long, source As String, description As String
    On Error GoTo Failed
    Set target = modAdmin.ResolveInteractiveAdminWorkbook()
    writeContext = modAuthUserCommands.CaptureContext()
    Set observer = New cCreateWarehouseActivity
    observer.Initialize target
    Call modRoleWorkbookSurfaces.EnsureAdminLegacyWorkbookSurface(target, report)
    If Not modLocalAddinsRegistration.EnsureLocalInvSysAddinsRegistered("", report) Then
        observer.OpeningFailed
        MsgBox "Current invSys add-ins are not registered cleanly for this Excel session." & vbCrLf & vbCrLf & _
               report, vbExclamation, "invSys Admin"
        Exit Sub
    End If
    frmCreateWarehouse.ObserveEntry observer, writeContext
    frmCreateWarehouse.Show
    Exit Sub
Failed:
    number = Err.Number: source = Err.Source: description = Err.Description
    If Not observer Is Nothing Then observer.OpeningFailed
    Err.Raise number, source, description
End Sub
