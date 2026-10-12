VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cAdminReconcileActivity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Invocation-owned observations only; authorization remains with the existing owner.
Private mContext As String, mAction As String, mNotice As String, mPresented As Boolean

Public Sub Capture()
    On Error GoTo Unavailable
    mContext = modActivity.CaptureContext()
    Start "ADMIN_RECONCILE_OPEN"
    Exit Sub
Unavailable:
    mNotice = "Tracking unavailable: Shipment Reconcile could not be recorded."
End Sub

Private Sub Start(ByVal controlId As String)
    Dim notice As String
    On Error GoTo Unavailable
    mAction = ""
    If mContext <> "" Then
        mAction = modActivity.BeginAction(controlId, mContext, notice)
        If mAction = "" And mContext <> modActivity.CaptureContext() Then _
            notice = "Tracking unavailable: the Reconcile session or warehouse changed."
    End If
    If notice <> "" Then mNotice = notice
    Exit Sub
Unavailable:
    mNotice = "Tracking unavailable: Shipment Reconcile could not be recorded."
End Sub

Public Sub Presented(ByVal visible As Boolean)
    If Not visible Or mPresented Then Exit Sub
    mPresented = True
    Finish "OPENED"
End Sub

Public Sub SelectAction(ByVal cancel As Boolean)
    If cancel Then Start "ADMIN_RECONCILE_CANCEL" Else Start "ADMIN_RECONCILE_INSPECT"
End Sub

Public Sub Finish(ByVal outcome As String, Optional ByVal eventId As String = "")
    Dim references As String, submission As String, action As String, notice As String
    On Error GoTo Unavailable
    action = mAction: mAction = ""
    If action = "" Then Exit Sub
    references = "[]"
    If eventId <> "" And (outcome = "CONFIRMED" Or outcome = "FAILED") Then
        submission = "Unknown"
        If outcome = "CONFIRMED" Then submission = "Submitted"
        references = modActivity.InventorySourceReferences(action, eventId, submission)
    End If
    modActivity.FinishAction action, outcome, notice, references
    If notice <> "" Then mNotice = notice
    Exit Sub
Unavailable:
    mNotice = "Tracking unavailable: Shipment Reconcile result could not be recorded."
End Sub

Public Property Get Notice() As String
    Notice = mNotice
End Property
