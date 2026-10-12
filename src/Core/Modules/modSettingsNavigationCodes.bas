Attribute VB_Name = "modSettingsNavigationCodes"
Option Explicit
Option Private Module

Public Function ControlIds() As Variant
    ControlIds = Array("ADMIN_SETTINGS_OPEN", "ADMIN_SETTINGS_PAGE_GENERAL", _
        "ADMIN_SETTINGS_PAGE_EVENT_TRACKING", "ADMIN_SETTINGS_PAGE_TRACKING", _
        "ADMIN_SETTINGS_PAGE_EVENT_DETAIL", "ADMIN_SETTINGS_PAGE_ACTION_PATHS")
End Function

Public Function Control(ByVal id As String) As Object
    Dim caption As String, surface As String, record As Object
    surface = "Admin > Settings > Event Tracking"
    Select Case id
        Case "ADMIN_SETTINGS_OPEN": caption = "Settings": surface = "Admin"
        Case "ADMIN_SETTINGS_PAGE_GENERAL": caption = "General": surface = "Admin > Settings"
        Case "ADMIN_SETTINGS_PAGE_EVENT_TRACKING": caption = "Event Tracking": surface = "Admin > Settings"
        Case "ADMIN_SETTINGS_PAGE_TRACKING": caption = "Tracking"
        Case "ADMIN_SETTINGS_PAGE_EVENT_DETAIL": caption = "Event Detail"
        Case "ADMIN_SETTINGS_PAGE_ACTION_PATHS": caption = "Action Paths"
        Case Else: Exit Function
    End Select
    Set record = CreateObject("Scripting.Dictionary")
    record.Add "ControlId", id: record.Add "OwnerId", "ADMIN_SETTINGS_UI"
    record.Add "Class", "Navigation": record.Add "Role", "Admin"
    record.Add "Caption", caption: record.Add "Surface", surface
    record.Add "Capability", "ADMIN_MAINT": record.Add "CodePrefix", id & "_"
    Set Control = record
End Function

Public Function Outcome(ByVal id As String, ByVal code As String) As Object
    Dim definition As Object, record As Object, severity As String, effect As String, message As String
    Set definition = Control(id)
    If definition Is Nothing Then Exit Function
    severity = "Info": effect = "Unknown"
    Select Case code
        Case "REQUESTED": message = "Settings navigation requested."
        Case "OPENED"
            If id <> "ADMIN_SETTINGS_OPEN" Then Exit Function
            effect = "Unchanged"
            message = "Settings form opened. No settings save or Domain application is asserted."
        Case "SELECTED"
            If id = "ADMIN_SETTINGS_OPEN" Then Exit Function
            effect = "Unchanged"
            message = "Settings page selected. No settings save or Domain application is asserted."
        Case "DENIED", "REJECTED"
            If id <> "ADMIN_SETTINGS_OPEN" Then Exit Function
            effect = "Unchanged": severity = "Warning"
            message = "Settings could not be opened in the current context."
            If code = "DENIED" Then severity = "Blocked"
        Case "FAILED"
            severity = "Error": message = "Settings navigation failed; inspect the form before retrying."
        Case Else: Exit Function
    End Select
    Set record = CreateObject("Scripting.Dictionary")
    record.Add "EventCode", id & "_" & code: record.Add "OutcomeCode", code
    record.Add "Severity", severity: record.Add "DataEffect", effect
    record.Add "UserMessage", message: record.Add "NextStep", ""
    Set Outcome = record
End Function
