Attribute VB_Name = "modAdminActionObservation"
Option Explicit
Option Private Module

' Optional observations never determine or replace an Admin business result.
Public Function BeginObservation(ByVal controlId As String, ByRef notice As String, _
                      ByVal unavailableMessage As String) As String
    Dim context As String
    On Error GoTo Unavailable
    context = modActivity.CaptureContext()
    If context <> "" Then
        BeginObservation = modActivity.BeginAction(controlId, context, notice)
        If context <> modActivity.CaptureContext() Then GoTo Unavailable
    End If
    Exit Function
Unavailable:
    notice = unavailableMessage
End Function

Public Sub FinishObservation(ByVal action As String, ByVal outcome As String, ByRef notice As String, _
                  ByVal unavailableMessage As String)
    Dim resultNotice As String
    On Error GoTo Unavailable
    If action = "" Then Exit Sub
    modActivity.FinishAction action, outcome, resultNotice
    If resultNotice <> "" Then notice = resultNotice
    Exit Sub
Unavailable:
    notice = unavailableMessage
End Sub

' Optional reads may activate workbooks. Restore only the exact still-live sheet.
' Never resolve a substitute workbook or reopen one that the operator closed.
Public Function CurrentSurface() As Object
    On Error Resume Next
    Set CurrentSurface = Application.ActiveSheet
End Function

Public Sub RestoreSurface(ByVal surface As Object)
    On Error Resume Next
    If surface Is Nothing Then Exit Sub
    surface.Parent.Activate
    surface.Activate
End Sub
