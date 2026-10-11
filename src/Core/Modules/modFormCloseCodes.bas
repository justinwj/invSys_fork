Attribute VB_Name = "modFormCloseCodes"
Option Explicit
Option Private Module

' Shared immutable dismissal vocabulary; no form lifecycle or authority work.
Public Function Control(ByVal id As String) As Object
    Dim record As Object, owner As String, role As String, surface As String, capability As String
    Select Case id
        Case "PRODUCTION_CLOSE"
            owner = "PRODUCTION_WORKFLOW": role = "Production"
            surface = "Operations > Production": capability = "PROD_POST"
        Case "ADMIN_SETTINGS_CLOSE"
            owner = "ADMIN_SETTINGS_UI": role = "Admin"
            surface = "Admin > Settings": capability = "ADMIN_MAINT"
        Case Else: Exit Function
    End Select
    Set record = CreateObject("Scripting.Dictionary")
    record.Add "ControlId", id: record.Add "OwnerId", owner
    record.Add "Class", "Command": record.Add "Role", role
    record.Add "Caption", "Close": record.Add "Surface", surface
    record.Add "Capability", capability: record.Add "CodePrefix", id & "_"
    Set Control = record
End Function

Public Function Outcome(ByVal id As String, ByVal code As String) As Object
    Dim definition As Object, record As Object, label As String, formName As String
    Dim severity As String, effect As String, message As String, nextStep As String
    Set definition = Control(id)
    If definition Is Nothing Then Exit Function
    If id = "PRODUCTION_CLOSE" Then
        label = "Production": formName = "the Production form"
    Else
        label = "Settings": formName = "Settings"
    End If
    severity = "Info": effect = "Unknown"
    Select Case code
        Case "REQUESTED": message = label & " form dismissal requested."
        Case "CLOSED"
            effect = "Unchanged"
            If id = "PRODUCTION_CLOSE" Then
                message = "Production form dismissed; no save, posting or Domain application is asserted."
            Else
                message = "Settings form dismissed; unsaved form edits discarded. No settings save or Domain application is asserted."
            End If
        Case "FAILED"
            severity = "Error": message = label & " form dismissal failed; inspect the form before retrying."
            nextStep = "Inspect whether " & formName & " is still open before retrying Close."
        Case Else: Exit Function
    End Select
    Set record = CreateObject("Scripting.Dictionary")
    record.Add "EventCode", id & "_" & code: record.Add "OutcomeCode", code
    record.Add "Severity", severity: record.Add "DataEffect", effect
    record.Add "UserMessage", message: record.Add "NextStep", nextStep
    Set Outcome = record
End Function
