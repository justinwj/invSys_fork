Attribute VB_Name = "modAdminSettingsClose"
Option Explicit
Option Private Module

' Observation never grants or gates dismissal; internal Unload does not enter here.
Public Sub CloseForm(ByVal form As frmAdminSettings, ByVal capturedContext As String, _
                     Optional ByVal nativeClosing As Boolean = False)
    Dim activityId As String, notice As String, outcome As String
    On Error Resume Next
    If capturedContext <> "" And capturedContext = modActivity.CaptureContext() Then _
        activityId = modActivity.BeginAction("ADMIN_SETTINGS_CLOSE", capturedContext, notice)
    On Error GoTo DismissalFailed
    If nativeClosing Then form.Hide Else Unload form
    outcome = "CLOSED"
    GoTo Finished
DismissalFailed:
    outcome = "FAILED"
Finished:
    On Error Resume Next
    If activityId <> "" Then modActivity.FinishAction activityId, outcome, notice
    If outcome = "FAILED" Then notice = "Settings could not be dismissed. Inspect the form before retrying. " & notice
    If notice <> "" Then MsgBox Trim$(notice), vbExclamation, "invSys Settings"
End Sub
