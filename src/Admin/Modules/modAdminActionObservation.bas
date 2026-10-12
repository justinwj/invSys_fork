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
