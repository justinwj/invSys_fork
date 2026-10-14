Attribute VB_Name = "modTesterActivityCodes"
Option Explicit
Option Private Module

Public Function ControlIds() As Variant
    ControlIds = Array("ADMIN_TESTER_OPEN", "ADMIN_TESTER_SETUP", "ADMIN_TESTER_DELETE", _
        "ADMIN_TESTER_OPEN_WORKBOOK", "ADMIN_TESTER_BROWSE_HUB", "ADMIN_TESTER_FIND_SHAREPOINT", "ADMIN_TESTER_CLOSE")
End Function

Public Function Control(ByVal id As String) As Object
    Dim caption As String, kind As String, record As Object
    kind = "Command"
    Select Case id
        Case "ADMIN_TESTER_OPEN": caption = "Test Environment Setup": kind = "Navigation"
        Case "ADMIN_TESTER_SETUP": caption = "Setup"
        Case "ADMIN_TESTER_DELETE": caption = "Delete Generated"
        Case "ADMIN_TESTER_OPEN_WORKBOOK": caption = "Open Workbook": kind = "Navigation"
        Case "ADMIN_TESTER_BROWSE_HUB": caption = "Find warehouse folder"
        Case "ADMIN_TESTER_FIND_SHAREPOINT": caption = "Find publication folder"
        Case "ADMIN_TESTER_CLOSE": caption = "Close": kind = "Navigation"
        Case Else: Exit Function
    End Select
    Set record = CreateObject("Scripting.Dictionary")
    record.Add "ControlId", id: record.Add "OwnerId", "ADMIN_TESTER_SETUP"
    record.Add "Class", kind: record.Add "Role", "Admin": record.Add "Caption", caption
    record.Add "Surface", "Admin > Test Environment Setup": record.Add "Capability", "ADMIN_MAINT"
    record.Add "CodePrefix", id & "_"
    Set Control = record
End Function

Public Function PositiveOutcome(ByVal id As String, ByVal code As String) As Boolean
    Select Case id
        Case "ADMIN_TESTER_SETUP", "ADMIN_TESTER_DELETE": PositiveOutcome = (code = "COMPLETED")
        Case "ADMIN_TESTER_OPEN", "ADMIN_TESTER_OPEN_WORKBOOK": PositiveOutcome = (code = "OPENED")
        Case "ADMIN_TESTER_BROWSE_HUB", "ADMIN_TESTER_FIND_SHAREPOINT": PositiveOutcome = (code = "STAGED")
        Case "ADMIN_TESTER_CLOSE": PositiveOutcome = (code = "CLOSED")
    End Select
End Function

Public Function Outcome(ByVal id As String, ByVal code As String) As Object
    Dim definition As Object, record As Object, message As String, severity As String, effect As String
    Set definition = Control(id)
    If definition Is Nothing Then Exit Function
    severity = "Info": effect = "Unknown"
    Select Case code
        Case "REQUESTED": message = definition("Caption") & " requested."
        Case "REJECTED": severity = "Warning": message = "The request was refused. Read the form result."
        Case "DENIED": severity = "Warning": message = "Source authority or destination scope changed. No rollback is asserted."
        Case "FAILED": severity = "Error": message = "The action could not finish. No rollback is asserted."
        Case "CANCELLED"
            If id <> "ADMIN_TESTER_DELETE" And id <> "ADMIN_TESTER_BROWSE_HUB" Then Exit Function
            effect = "Unchanged": message = "The action was cancelled."
        Case "UNAVAILABLE"
            If id <> "ADMIN_TESTER_FIND_SHAREPOINT" Then Exit Function
            severity = "Warning": message = "No publication folder was detected or selected."
        Case Else
            If Not PositiveOutcome(id, code) Then Exit Function
            Select Case code
                Case "COMPLETED": message = "The owning Tester command confirmed completion. No inventory application or replay eligibility is asserted."
                Case "OPENED": message = definition("Caption") & " opened. No inventory application is asserted."
                Case "STAGED": message = "The folder choice was staged. No provisioning is asserted."
                Case "CLOSED": message = "Test Environment Setup closed. No save or provisioning is asserted."
            End Select
    End Select
    Set record = CreateObject("Scripting.Dictionary")
    record.Add "EventCode", id & "_" & code: record.Add "OutcomeCode", code
    record.Add "Severity", severity: record.Add "DataEffect", effect
    record.Add "UserMessage", message: record.Add "NextStep", ""
    Set Outcome = record
End Function
