Attribute VB_Name = "modCreateWarehouseCodes"
Option Explicit
Option Private Module

Public Function ControlIds() As Variant
    ControlIds = Array("ADMIN_CREATE_OPEN", "ADMIN_CREATE_SUBMIT", "ADMIN_CREATE_RETRY_PUBLISH", _
        "ADMIN_CREATE_PUBLISH_EXISTING", "ADMIN_CREATE_OPEN_WORKBOOK", "ADMIN_CREATE_BROWSE_HUB", _
        "ADMIN_CREATE_FIND_SHAREPOINT", "ADMIN_CREATE_PURPOSE", "ADMIN_CREATE_PUBLISH_OPTION", "ADMIN_CREATE_CLOSE")
End Function

Public Function Control(ByVal id As String) As Object
    Dim caption As String, kind As String, record As Object
    kind = "Command"
    Select Case id
        Case "ADMIN_CREATE_OPEN": caption = "Create New Warehouse": kind = "Navigation"
        Case "ADMIN_CREATE_SUBMIT": caption = "Create"
        Case "ADMIN_CREATE_RETRY_PUBLISH": caption = "Retry Publish"
        Case "ADMIN_CREATE_PUBLISH_EXISTING": caption = "Publish Existing Warehouse"
        Case "ADMIN_CREATE_OPEN_WORKBOOK": caption = "Open Workbook": kind = "Navigation"
        Case "ADMIN_CREATE_BROWSE_HUB": caption = "Browse..."
        Case "ADMIN_CREATE_FIND_SHAREPOINT": caption = "Find..."
        Case "ADMIN_CREATE_PURPOSE": caption = "Warehouse purpose": kind = "Navigation"
        Case "ADMIN_CREATE_PUBLISH_OPTION": caption = "Publish initial artifacts to SharePoint"
        Case "ADMIN_CREATE_CLOSE": caption = "Close": kind = "Navigation"
        Case Else: Exit Function
    End Select
    Set record = CreateObject("Scripting.Dictionary")
    record.Add "ControlId", id: record.Add "OwnerId", "ADMIN_WAREHOUSE_CREATE"
    record.Add "Class", kind: record.Add "Role", "Admin": record.Add "Caption", caption
    record.Add "Surface", "Admin > Create Warehouse": record.Add "Capability", "ADMIN_MAINT"
    record.Add "CodePrefix", id & "_"
    Set Control = record
End Function

Public Function PositiveOutcome(ByVal id As String, ByVal code As String) As Boolean
    Select Case id
        Case "ADMIN_CREATE_SUBMIT": PositiveOutcome = (code = "CREATED_LOCAL" Or code = "CREATED_PUBLISHED")
        Case "ADMIN_CREATE_RETRY_PUBLISH", "ADMIN_CREATE_PUBLISH_EXISTING": PositiveOutcome = (code = "PUBLISHED")
        Case "ADMIN_CREATE_OPEN", "ADMIN_CREATE_OPEN_WORKBOOK": PositiveOutcome = (code = "OPENED")
        Case "ADMIN_CREATE_BROWSE_HUB", "ADMIN_CREATE_FIND_SHAREPOINT", "ADMIN_CREATE_PUBLISH_OPTION": PositiveOutcome = (code = "STAGED")
        Case "ADMIN_CREATE_PURPOSE": PositiveOutcome = (code = "SELECTED")
        Case "ADMIN_CREATE_CLOSE": PositiveOutcome = (code = "CLOSED")
    End Select
End Function

Public Function Outcome(ByVal id As String, ByVal code As String) As Object
    Dim definition As Object, record As Object, message As String, severity As String, effect As String
    Set definition = Control(id)
    If definition Is Nothing Then Exit Function
    severity = "Info": effect = "Unknown"
    Select Case code
        Case "REQUESTED": message = definition("Caption") & " requested."
        Case "REJECTED": severity = "Warning": message = "The request was refused by validation. Read the form result."
        Case "DENIED": severity = "Warning": message = "The request lacked the captured source Admin authority."
        Case "FAILED": severity = "Error": message = "The action could not finish. Inspect the current state; no rollback is asserted."
        Case "CREATED_PUBLISH_FAILED", "CREATED_PUBLISH_DENIED"
            If id <> "ADMIN_CREATE_SUBMIT" Then Exit Function
            severity = "Error": message = "Local creation succeeded; initial publication did not finish."
            If code = "CREATED_PUBLISH_DENIED" Then message = "Local creation succeeded; initial publication lacked current source authority."
        Case "CANCELLED"
            If id <> "ADMIN_CREATE_BROWSE_HUB" Then Exit Function
            effect = "Unchanged": message = "Folder selection was cancelled."
        Case "UNAVAILABLE"
            If id <> "ADMIN_CREATE_FIND_SHAREPOINT" Then Exit Function
            severity = "Warning": message = "No publication folder was detected or selected."
        Case Else
            If Not PositiveOutcome(id, code) Then Exit Function
            Select Case code
                Case "CREATED_LOCAL": message = "Local warehouse creation completed; initial publication was not requested."
                Case "CREATED_PUBLISHED": message = "Local warehouse creation and initial publication completed."
                Case "PUBLISHED": message = "Initial publication completed for the existing warehouse."
                Case "OPENED": message = definition("Caption") & " opened. No inventory application is asserted."
                Case "STAGED": message = "The form choice was staged. No warehouse creation or publication is asserted."
                Case "SELECTED": message = "The warehouse purpose choice changed. Its value is not recorded."
                Case "CLOSED": message = "Create Warehouse closed. No save or publication is asserted."
            End Select
    End Select
    Set record = CreateObject("Scripting.Dictionary")
    record.Add "EventCode", id & "_" & code: record.Add "OutcomeCode", code
    record.Add "Severity", severity: record.Add "DataEffect", effect
    record.Add "UserMessage", message: record.Add "NextStep", ""
    Set Outcome = record
End Function
