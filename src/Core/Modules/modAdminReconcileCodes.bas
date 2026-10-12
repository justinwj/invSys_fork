Attribute VB_Name = "modAdminReconcileCodes"
Option Explicit
Option Private Module

Public Function ControlIds() As Variant
    ControlIds = Array("ADMIN_RECONCILE_OPEN", "ADMIN_RECONCILE_INSPECT", "ADMIN_RECONCILE_CANCEL")
End Function

Public Function Control(ByVal id As String) As Object
    Dim record As Object, caption As String, surface As String, kind As String
    surface = "Admin > Shipment Reconcile": kind = "Navigation"
    Select Case id
        Case "ADMIN_RECONCILE_OPEN": caption = "Shipment Reconcile": surface = "Admin"
        Case "ADMIN_RECONCILE_INSPECT": caption = "Inspect": kind = "Command"
        Case "ADMIN_RECONCILE_CANCEL": caption = "Cancel"
        Case Else: Exit Function
    End Select
    Set record = CreateObject("Scripting.Dictionary")
    record.Add "ControlId", id: record.Add "OwnerId", "ADMIN_SHIPMENT_RECONCILE"
    record.Add "Class", kind: record.Add "Role", "Admin"
    record.Add "Caption", caption: record.Add "Surface", surface
    record.Add "Capability", "ADMIN_MAINT": record.Add "CodePrefix", id & "_"
    Set Control = record
End Function

Public Function Outcome(ByVal id As String, ByVal code As String) As Object
    Dim record As Object, severity As String, effect As String, message As String
    Set record = Control(id)
    If record Is Nothing Then Exit Function
    severity = "Info": effect = "Unchanged"
    Select Case code
        Case "REQUESTED": effect = "Unknown": message = "Shipment Reconcile action requested."
        Case "OPENED"
            If id <> "ADMIN_RECONCILE_OPEN" Then Exit Function
            message = "Shipment history opened. No correction was submitted."
        Case "NO_CORRECTION"
            If id <> "ADMIN_RECONCILE_INSPECT" Then Exit Function
            message = "Inspection required no correction. No event was submitted."
        Case "CANCELLED"
            If id = "ADMIN_RECONCILE_OPEN" Then Exit Function
            severity = "Notice": message = "Shipment Reconcile cancelled. No correction was submitted."
        Case "REJECTED"
            If id = "ADMIN_RECONCILE_CANCEL" Then Exit Function
            severity = "Warning": message = "Shipment Reconcile validation prevented submission."
        Case "DENIED"
            If id <> "ADMIN_RECONCILE_INSPECT" Then Exit Function
            severity = "Blocked": message = "Shipment Reconcile authorization refused submission."
        Case "CONFIRMED"
            If id <> "ADMIN_RECONCILE_INSPECT" Then Exit Function
            effect = "Unknown": message = "Correction queued. Inventory application requires published Domain evidence."
        Case "FAILED"
            severity = "Error": effect = "Unknown"
            message = "Shipment Reconcile could not finish. Inspect the current state before retrying."
        Case Else: Exit Function
    End Select
    Set record = CreateObject("Scripting.Dictionary")
    record.Add "EventCode", id & "_" & code: record.Add "OutcomeCode", code
    record.Add "Severity", severity: record.Add "DataEffect", effect
    record.Add "UserMessage", message: record.Add "NextStep", ""
    Set Outcome = record
End Function
