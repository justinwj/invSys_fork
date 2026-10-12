Attribute VB_Name = "modAdminVerifyCodes"
Option Explicit
Option Private Module

Public Function Control(ByVal id As String) As Object
    Dim record As Object
    If id <> "ADMIN_VERIFY_ADDINS" Then Exit Function
    Set record = CreateObject("Scripting.Dictionary")
    record.Add "ControlId", id: record.Add "OwnerId", "ADMIN_ADDINS_VERIFY"
    record.Add "Class", "Command": record.Add "Role", "Admin"
    record.Add "Caption", "Verify Add-ins Published": record.Add "Surface", "Admin"
    record.Add "Capability", "ADMIN_MAINT": record.Add "CodePrefix", id & "_"
    Set Control = record
End Function

Public Function Outcome(ByVal id As String, ByVal code As String) As Object
    Dim record As Object, severity As String, message As String
    If id <> "ADMIN_VERIFY_ADDINS" Then Exit Function
    severity = "Info"
    Select Case code
        Case "REQUESTED": message = "Required add-in file verification requested."
        Case "VERIFIED": message = "The five required add-in files are present and nonempty. Package integrity and deployment are not asserted."
        Case "FAILED"
            severity = "Error"
            message = "Required add-in file verification did not succeed. Inspect the verification result."
        Case Else: Exit Function
    End Select
    Set record = CreateObject("Scripting.Dictionary")
    record.Add "EventCode", id & "_" & code: record.Add "OutcomeCode", code
    record.Add "Severity", severity: record.Add "DataEffect", "Unknown"
    record.Add "UserMessage", message: record.Add "NextStep", ""
    Set Outcome = record
End Function
