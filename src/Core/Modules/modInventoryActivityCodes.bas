Attribute VB_Name = "modInventoryActivityCodes"
Option Explicit
Option Private Module

Public Function Control(ByVal id As String) As Object
    Dim caption As String, record As Object
    Select Case id
        Case "ADMIN_INVENTORY_OPEN": caption = "Add/Edit Inventory"
        Case "ADMIN_INVENTORY_ADD_MODE": caption = "Add Item Mode"
        Case "ADMIN_INVENTORY_EDIT_MODE": caption = "Edit Item"
        Case "ADMIN_INVENTORY_CLOSE": caption = "Close"
        Case Else: Exit Function
    End Select
    Set record = CreateObject("Scripting.Dictionary")
    record.Add "ControlId", id: record.Add "OwnerId", "ADMIN_INVENTORY"
    record.Add "Class", "Navigation": record.Add "Role", "Admin": record.Add "Caption", caption
    record.Add "Surface", "Admin > Inventory": record.Add "Capability", "ADMIN_MAINT"
    record.Add "CodePrefix", id & "_"
    Set Control = record
End Function

Public Function PositiveOutcome(ByVal id As String, ByVal code As String) As Boolean
    Select Case id
        Case "ADMIN_INVENTORY_OPEN": PositiveOutcome = (code = "OPENED")
        Case "ADMIN_INVENTORY_ADD_MODE", "ADMIN_INVENTORY_EDIT_MODE": PositiveOutcome = (code = "SELECTED")
        Case "ADMIN_INVENTORY_CLOSE": PositiveOutcome = (code = "CLOSED")
    End Select
End Function

Public Function Outcome(ByVal id As String, ByVal code As String) As Object
    Dim definition As Object, record As Object, message As String, severity As String, effect As String
    Set definition = Control(id)
    If definition Is Nothing Then Exit Function
    severity = "Info": effect = "Unknown"
    Select Case code
        Case "REQUESTED": message = definition("Caption") & " requested."
        Case "FAILED": severity = "Error": message = "Inventory navigation did not finish. Inspect the owning form; no rollback is asserted."
        Case Else
            If Not PositiveOutcome(id, code) Then Exit Function
            effect = "Unchanged"
            Select Case code
                Case "OPENED": message = "Add/Edit Inventory opened. No inventory change is asserted."
                Case "SELECTED": message = definition("Caption") & " selected. No inventory change is asserted."
                Case "CLOSED": message = "Add/Edit Inventory closed. No inventory change or rollback is asserted."
            End Select
    End Select
    Set record = CreateObject("Scripting.Dictionary")
    record.Add "EventCode", id & "_" & code: record.Add "OutcomeCode", code
    record.Add "Severity", severity: record.Add "DataEffect", effect
    record.Add "UserMessage", message: record.Add "NextStep", ""
    Set Outcome = record
End Function
