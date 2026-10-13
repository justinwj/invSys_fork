VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cCreateWarehouseActivity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mContext As String, mNotice As String, mOpening As String
Private mTarget As Workbook, mPresented As Boolean

Public Sub Initialize(ByVal target As Workbook)
    On Error GoTo Unavailable
    Set mTarget = target
    mContext = modActivity.CaptureContext()
    mOpening = BeginAction("OPEN")
    Exit Sub
Unavailable:
    UnavailableNotice
End Sub

Public Function BeginAction(ByVal suffix As String) As String
    Dim surface As Object
    On Error GoTo Unavailable
    mNotice = ""
    Set surface = modAdminActionObservation.CurrentSurface()
    If Not TargetLive() Then GoTo Unavailable
    BeginAction = modActivity.BeginAction("ADMIN_CREATE_" & suffix, mContext, mNotice)
    If mContext <> modActivity.CaptureContext() Then UnavailableNotice
Finished:
    modAdminActionObservation.RestoreSurface surface
    Exit Function
Unavailable:
    UnavailableNotice
    Resume Finished
End Function

Public Sub Finish(ByVal action As String, ByVal outcome As String)
    Dim surface As Object, notice As String
    On Error GoTo Unavailable
    Set surface = modAdminActionObservation.CurrentSurface()
    If action <> "" Then
        If Not TargetLive() Then GoTo Unavailable
        If Not modActivity.FinishAction(action, outcome, notice) And notice = "" Then _
            notice = "Tracking unavailable: the Create Warehouse result could not be recorded."
        If notice <> "" Then mNotice = notice
    End If
Finished:
    modAdminActionObservation.RestoreSurface surface
    ShowNotice
    Exit Sub
Unavailable:
    UnavailableNotice
    Resume Finished
End Sub

Public Sub Presented(ByVal visible As Boolean)
    If mPresented Or Not visible Then Exit Sub
    mPresented = True
    Finish mOpening, "OPENED"
End Sub

Public Sub OpeningFailed()
    If mPresented Then Exit Sub
    mPresented = True
    Finish mOpening, "FAILED"
End Sub

Private Function TargetLive() As Boolean
    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If wb Is mTarget Then TargetLive = True: Exit Function
    Next wb
End Function

Private Sub UnavailableNotice()
    mNotice = "Tracking unavailable: Create Warehouse could not be recorded."
    On Error Resume Next
    modActionRecording.ContextClosed mContext, True
End Sub

Private Sub ShowNotice()
    Dim notice As String
    notice = mNotice: mNotice = ""
    If notice <> "" Then MsgBox notice, vbExclamation, "invSys Admin"
End Sub
