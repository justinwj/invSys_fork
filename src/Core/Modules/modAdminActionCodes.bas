Attribute VB_Name = "modAdminActionCodes"
Option Explicit
Option Private Module

Public Function Control(ByVal id As String) As Object
    Dim record As Object, owner As String, actionClass As String, caption As String
    Select Case id
        Case "ADMIN_WAREHOUSE_ROOT_ADD"
            owner = "ADMIN_WAREHOUSE_ROOT_PREFERENCE": actionClass = "Command": caption = "Add Warehouse Root"
        Case "ADMIN_VERIFY_ADDINS"
            owner = "ADMIN_ADDINS_VERIFY": actionClass = "Command": caption = "Verify Add-ins Published"
        Case "ADMIN_CONSOLE_OPEN"
            owner = "ADMIN_CONSOLE": actionClass = "Navigation": caption = "Admin Console"
        Case "ADMIN_DIRECTORY_OPEN"
            owner = "ADMIN_WAREHOUSE_DIRECTORY": actionClass = "Navigation": caption = "View Warehouses"
        Case "ADMIN_DEMO_SEED"
            owner = "ADMIN_DEMO_INVENTORY": actionClass = "Command": caption = "Seed Demo Inventory"
        Case "ADMIN_USER_SAVE", "ADMIN_USER_DEACTIVATE", "ADMIN_USER_DELETE"
            owner = "ADMIN_USERS_ROLES": actionClass = "Command"
            If id = "ADMIN_USER_SAVE" Then caption = "Create / Update"
            If id = "ADMIN_USER_DEACTIVATE" Then caption = "Deactivate"
            If id = "ADMIN_USER_DELETE" Then caption = "Delete"
        Case Else: Exit Function
    End Select
    Set record = CreateObject("Scripting.Dictionary")
    record.Add "ControlId", id: record.Add "OwnerId", owner
    record.Add "Class", actionClass: record.Add "Role", "Admin"
    record.Add "Caption", caption: record.Add "Surface", "Admin"
    If id = "ADMIN_DEMO_SEED" Then record("Surface") = "Admin > Demo Inventory"
    If owner = "ADMIN_USERS_ROLES" Then record("Surface") = "Admin > Users & Roles"
    record.Add "Capability", "ADMIN_MAINT": record.Add "CodePrefix", id & "_"
    Set Control = record
End Function

Public Function Outcome(ByVal id As String, ByVal code As String) As Object
    Dim record As Object, severity As String, message As String, effect As String
    severity = "Info": effect = "Unknown"
    Select Case id
        Case "ADMIN_WAREHOUSE_ROOT_ADD"
            Select Case code
                Case "REQUESTED": message = "Add Warehouse Root requested."
                Case "CANCELLED": effect = "Unchanged": message = "Add Warehouse Root was cancelled before writing or refreshing."
                Case "REMEMBERED_REFRESHED": message = "The root preference was confirmed and the directory refreshed. Warehouse creation and target selection are not asserted."
                Case "REMEMBERED_REFRESH_FAILED": severity = "Error": message = "The root preference was confirmed, but the directory could not be refreshed."
                Case "UNCONFIRMED_REFRESHED": severity = "Warning": message = "The root preference could not be confirmed. The directory refreshed."
                Case "UNCONFIRMED_REFRESH_FAILED": severity = "Error": message = "The root preference could not be confirmed and the directory could not be refreshed. Inspect the current state before retrying."
                Case "FAILED": message = "Add Warehouse Root could not finish. Inspect the current state before retrying."
                Case Else: Exit Function
            End Select
        Case "ADMIN_USER_SAVE", "ADMIN_USER_DEACTIVATE", "ADMIN_USER_DELETE"
            Select Case code
                Case "REQUESTED": message = "Users & Roles command requested."
                Case "SAVED": message = "The Users & Roles command was saved. Domain application is not asserted."
                Case "REJECTED": effect = "Unchanged": severity = "Warning": message = "The Users & Roles command was refused before writing."
                Case "DENIED": severity = "Warning": message = "The Users & Roles command lacked current authority. Inspect the current state before retrying."
                Case "FAILED": message = "The Users & Roles command could not finish. Inspect the current Auth state before retrying."
                Case "CANCELLED"
                    If id <> "ADMIN_USER_DELETE" Then Exit Function
                    effect = "Unchanged": message = "User deletion was cancelled."
                Case Else: Exit Function
            End Select
        Case "ADMIN_VERIFY_ADDINS"
            Select Case code
                Case "REQUESTED": message = "Required add-in file verification requested."
                Case "VERIFIED": message = "The five required add-in files are present and nonempty. Package integrity and deployment are not asserted."
                Case "FAILED": message = "Required add-in file verification did not succeed. Inspect the verification result."
                Case Else: Exit Function
            End Select
        Case "ADMIN_CONSOLE_OPEN"
            Select Case code
                Case "REQUESTED": message = "Admin Console opening requested."
                Case "OPENED": message = "Admin Console opened. Processing and Domain application are not asserted."
                Case "FAILED": message = "Admin Console could not be opened."
                Case Else: Exit Function
            End Select
        Case "ADMIN_DIRECTORY_OPEN"
            Select Case code
                Case "REQUESTED": message = "Warehouse directory opening requested."
                Case "OPENED": message = "Warehouse directory opened. Warehouse creation and target selection are not asserted."
                Case "FAILED": message = "Warehouse directory could not be opened."
                Case Else: Exit Function
            End Select
        Case "ADMIN_DEMO_SEED"
            Select Case code
                Case "REQUESTED": message = "Demo inventory Seed requested."
                Case "APPLIED": effect = "Changed": message = "This Seed event has a saved Domain application receipt. Publication and refresh are separate."
                Case "UNCHANGED": effect = "Unchanged": message = "The requested demo inventory is already present. No new inventory event was submitted."
                Case "PENDING": severity = "Notice": message = "The Seed event was submitted; its application is not confirmed."
                Case "REJECTED": effect = "Unchanged": severity = "Warning": message = "Seed validation refused the request before submission."
                Case "FAILED": message = "Seed could not finish. Inspect the current state before retrying."
                Case Else: Exit Function
            End Select
        Case Else: Exit Function
    End Select
    If code = "FAILED" Then severity = "Error"
    Set record = CreateObject("Scripting.Dictionary")
    record.Add "EventCode", id & "_" & code: record.Add "OutcomeCode", code
    record.Add "Severity", severity: record.Add "DataEffect", effect
    record.Add "UserMessage", message: record.Add "NextStep", ""
    Set Outcome = record
End Function
