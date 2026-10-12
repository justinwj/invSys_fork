VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cAdminSettingsNavigation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' One form-local observer per MultiPage; no selected values or input are persisted.
Private WithEvents mPage As MSForms.MultiPage
Private mStatus As MSForms.Label
Private mContext As String
Private mIds As Variant
Private mArmed As Boolean
Private mOriginIndex As Long

Public Sub Attach(ByVal page As MSForms.MultiPage, ByVal status As MSForms.Label, _
                  ByVal capturedContext As String, ByVal nested As Boolean)
    Set mPage = page: Set mStatus = status: mContext = capturedContext
    If nested Then
        mIds = Array("ADMIN_SETTINGS_PAGE_TRACKING", "ADMIN_SETTINGS_PAGE_EVENT_DETAIL", "ADMIN_SETTINGS_PAGE_ACTION_PATHS")
    Else
        mIds = Array("ADMIN_SETTINGS_PAGE_GENERAL", "ADMIN_SETTINGS_PAGE_EVENT_TRACKING")
    End If
End Sub

Public Sub Detach()
    mArmed = False: Set mPage = Nothing: Set mStatus = Nothing
End Sub

Private Sub mPage_MouseDown(ByVal Index As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    mOriginIndex = mPage.Value
    mArmed = (Button = 1 And Index >= 0)
End Sub

Private Sub mPage_MouseUp(ByVal Index As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    mArmed = False
End Sub

Private Sub mPage_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    mOriginIndex = mPage.Value
    Select Case KeyCode
        Case vbKeyLeft, vbKeyRight, vbKeyUp, vbKeyDown, vbKeyHome, vbKeyEnd, vbKeyPageUp, vbKeyPageDown
            mArmed = True
        Case vbKeyTab: mArmed = ((Shift And 2) <> 0)
        Case Else: mArmed = False
    End Select
End Sub

Private Sub mPage_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    mArmed = False
End Sub

Private Sub mPage_Change()
    Dim activityId As String, notice As String, index As Long
    If Not mArmed Then Exit Sub
    mArmed = False
    On Error GoTo Failed
    If mContext = "" Or mContext <> modActivity.CaptureContext() Then
        mStatus.Caption = "Session or warehouse changed. Reopen Settings."
        Exit Sub
    End If
    index = mPage.Value
    If index = mOriginIndex Then Exit Sub
    If index < LBound(mIds) Or index > UBound(mIds) Or Not mPage.Visible Then Exit Sub
    activityId = modActivity.BeginAction(CStr(mIds(index)), mContext, notice)
    If activityId <> "" Then modActivity.FinishAction activityId, "SELECTED", notice
    If notice <> "" Then mStatus.Caption = notice
    Exit Sub
Failed:
    On Error Resume Next
    If activityId <> "" Then modActivity.FinishAction activityId, "FAILED", notice
    mStatus.Caption = "Tracking unavailable: Settings page selection could not be recorded."
End Sub
