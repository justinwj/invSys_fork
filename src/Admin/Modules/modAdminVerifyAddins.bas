Attribute VB_Name = "modAdminVerifyAddins"
Option Explicit

' One invocation captures its operator workbook before optional tracking.
Public Sub VerifyAddinsPublished()
    Dim targetWb As Workbook, report As String, detail As String, action As String, addinsRoot As String
    Dim notice As String, outcome As String, errorNumber As Long, errorSource As String, errorDescription As String
    On Error GoTo Failed
    Set targetWb = modAdmin.ResolveInteractiveAdminWorkbook()
    addinsRoot = modAddinsPublish.CaptureVerificationRoot()
    action = BeginObservation(notice)
    Call modRoleWorkbookSurfaces.EnsureAdminLegacyWorkbookSurface(targetWb, report)
    If modAddinsPublish.VerifyAddinsPublished(addinsRoot, True) Then
        outcome = "VERIFIED"
        MsgBox "All required add-ins are published." & vbCrLf & modAddinsPublish.GetLastAddinsPublishReport(), vbInformation, "invSys Admin"
    Else
        outcome = "FAILED"
        detail = modAddinsPublish.GetLastAddinsPublishReport()
        If Len(detail) = 0 Then detail = "One or more required add-ins are missing or zero-byte."
        If InStr(1, detail, "PathSharePointRoot is not configured", vbTextCompare) > 0 Then
            detail = detail & vbCrLf & _
                     "Use Create New Warehouse or Test Environment Setup to choose the locally synced invSys SharePoint root first."
        End If
        MsgBox "Add-ins publish verification failed." & vbCrLf & detail, vbExclamation, "invSys Admin"
    End If
    FinishObservation action, outcome, notice
    If notice <> "" Then MsgBox notice, vbExclamation, "invSys Admin"
    Exit Sub
Failed:
    errorNumber = Err.Number: errorSource = Err.Source: errorDescription = Err.Description
    FinishObservation action, "FAILED", notice
    If notice <> "" Then MsgBox notice, vbExclamation, "invSys Admin"
    Err.Raise errorNumber, errorSource, errorDescription
End Sub

Private Function BeginObservation(ByRef notice As String) As String
    Dim context As String
    On Error GoTo Unavailable
    context = modActivity.CaptureContext()
    If context <> "" Then
        BeginObservation = modActivity.BeginAction("ADMIN_VERIFY_ADDINS", context, notice)
        If context <> modActivity.CaptureContext() Then GoTo Unavailable
    End If
    Exit Function
Unavailable:
    notice = "Tracking unavailable: add-in file verification could not be recorded."
End Function

Private Sub FinishObservation(ByVal action As String, ByVal outcome As String, ByRef notice As String)
    Dim resultNotice As String
    On Error GoTo Unavailable
    If action = "" Then Exit Sub
    modActivity.FinishAction action, outcome, resultNotice
    If resultNotice <> "" Then notice = resultNotice
    Exit Sub
Unavailable:
    notice = "Tracking unavailable: the add-in file verification result could not be recorded."
End Sub
