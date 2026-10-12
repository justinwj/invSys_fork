VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmShipmentReconcileSelection
   Caption         =   "invSys Admin - Shipment Reconcile"
   ClientHeight    =   7200
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10800
   StartUpPosition =   1
End
Attribute VB_Name = "frmShipmentReconcileSelection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@RuntimeStubUserFormCode
Option Explicit

Private mHistory As MSForms.TextBox, mSelection As MSForms.TextBox
Private WithEvents mInspect As MSForms.CommandButton
Private WithEvents mCancel As MSForms.CommandButton
Private mAnchors As Object, mResizeReady As Boolean, mSelectedText As String
Private mActivity As cAdminReconcileActivity

Private Sub UserForm_Initialize()
    Dim definition As Variant, control As Object
    Me.Caption = "invSys Admin - Shipment Reconcile"
    Me.Width = 720: Me.Height = 480
    Set mAnchors = modDynamicForms.CreateFormAnchorManager()
    mAnchors.Initialize Me, 540, 360
    For Each definition In Array( _
        Array("Label", "lblInstruction", "Enter the number of the server shipment deduction log line to inspect.", 12, 10, 684, 28, 7), _
        Array("TextBox", "txtHistory", "", 12, 44, 684, 330, 15), _
        Array("Label", "lblSelection", "Entry number", 12, 394, 86, 18, 9), _
        Array("TextBox", "txtSelection", "", 104, 390, 76, 24, 9), _
        Array("CommandButton", "btnInspect", "Inspect", 500, 390, 96, 28, 12), _
        Array("CommandButton", "btnCancel", "Cancel", 608, 390, 88, 28, 12))
        Set control = Me.Controls.Add("Forms." & definition(0) & ".1", CStr(definition(1)), True)
        control.Move definition(3), definition(4), definition(5), definition(6)
        control.Font.Size = 10
        If definition(0) = "Label" Or definition(0) = "CommandButton" Then control.Caption = definition(2)
        If definition(0) = "Label" Then control.WordWrap = True
        mAnchors.Add control, CLng(definition(7))
    Next definition
    Set mHistory = Me.Controls("txtHistory"): Set mSelection = Me.Controls("txtSelection")
    Set mInspect = Me.Controls("btnInspect"): Set mCancel = Me.Controls("btnCancel")
    mHistory.MultiLine = True: mHistory.WordWrap = True: mHistory.ScrollBars = fmScrollBarsVertical
    mHistory.Locked = True: mHistory.MaxLength = 0: mHistory.TabIndex = 0
    mSelection.TabIndex = 1: mInspect.TabIndex = 2: mCancel.TabIndex = 3
    mInspect.Default = True: mCancel.Cancel = True
End Sub

Public Sub Configure(ByVal history As String, Optional ByVal observer As cAdminReconcileActivity = Nothing)
    Set mActivity = observer
    mHistory.Value = history: mHistory.SelStart = 0
    mSelection.Value = "": mSelectedText = ""
End Sub

Public Property Get SelectedText() As String
    SelectedText = mSelectedText
End Property

Private Sub mInspect_Click()
    If Not mActivity Is Nothing Then mActivity.SelectAction False
    mSelectedText = Trim$(CStr(mSelection.Value))
    Me.Hide
End Sub

Private Sub mCancel_Click()
    If Not mActivity Is Nothing Then mActivity.SelectAction True
    mSelectedText = ""
    Me.Hide
    If Not mActivity Is Nothing Then mActivity.Finish "CANCELLED"
End Sub

Private Sub UserForm_Activate()
    If Not mResizeReady Then
        modUserFormResizeWin.EnableResizableUserForm Me, True, True
        mResizeReady = True
    End If
    mAnchors.ResizeControls
    mSelection.SetFocus
    If Not mActivity Is Nothing Then mActivity.Presented Me.Visible
End Sub

Private Sub UserForm_Layout()
    If Not mAnchors Is Nothing Then mAnchors.ResizeControls
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then
        Cancel = True: mCancel_Click
    Else
        Set mAnchors = Nothing
        Set mActivity = Nothing
    End If
End Sub
