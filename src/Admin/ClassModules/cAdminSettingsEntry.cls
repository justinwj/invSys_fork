VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cAdminSettingsEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' One public invocation owns this observer; it never owns a form or grants access.
Private mContext As String
Private mActivityId As String
Private mNotice As String
Private mFinished As Boolean

Public Sub OpenSettings()
    Dim report As String, notice As String, targetWb As Workbook
    BeginObservation
    On Error GoTo OpeningFailed
    If Not modRoleUiAccess.RequireCurrentUserCapabilityCached("ADMIN_MAINT", "", report) Then
        Finish "DENIED": GoTo Finished
    End If
    Set targetWb = modAdmin.ResolveInteractiveAdminWorkbook()
    Call modRoleWorkbookSurfaces.EnsureAdminLegacyWorkbookSurface(targetWb, report)
    If Not modConfig.IsLoaded() Then
        If Not modConfig.LoadConfig("", "") Then
            MsgBox "Canonical config could not be loaded." & vbCrLf & vbCrLf & _
                   modConfig.Validate(), vbExclamation, "invSys Settings"
            Finish "REJECTED": GoTo Finished
        End If
    End If
    frmAdminSettings.ObservePublicEntry Me
    frmAdminSettings.Show
Finished:
    notice = TakeNotice()
    If notice <> "" Then MsgBox notice, vbExclamation, "invSys Settings"
    Exit Sub
OpeningFailed:
    Finish "FAILED"
    MsgBox "Settings could not be opened. Inspect the current window before retrying.", vbExclamation, "invSys Settings"
    Resume Finished
End Sub

Private Sub BeginObservation()
    On Error GoTo Unavailable
    mContext = modActivity.CaptureContext()
    If mContext <> "" Then mActivityId = modActivity.BeginAction("ADMIN_SETTINGS_OPEN", mContext, mNotice)
    Exit Sub
Unavailable:
    mNotice = "Tracking unavailable: Settings opening could not be recorded."
End Sub

Public Sub Presented(ByVal capturedContext As String, ByVal visible As Boolean)
    If Not visible Or mFinished Then Exit Sub
    If capturedContext <> mContext Then
        mNotice = "Tracking unavailable: the form context changed. Reopen Settings."
        mFinished = True
    Else
        Finish "OPENED"
    End If
End Sub

Public Sub Finish(ByVal outcome As String)
    If mFinished Then Exit Sub
    mFinished = True
    On Error GoTo Unavailable
    If mActivityId <> "" Then modActivity.FinishAction mActivityId, outcome, mNotice
    Exit Sub
Unavailable:
    mNotice = "Tracking unavailable: Settings opening could not be recorded."
End Sub

Public Function TakeNotice() As String
    TakeNotice = mNotice
    mNotice = ""
End Function
