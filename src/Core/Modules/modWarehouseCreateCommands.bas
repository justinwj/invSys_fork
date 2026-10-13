Attribute VB_Name = "modWarehouseCreateCommands"
Option Explicit

Private mWriting As Boolean

' D8-CREATE-01: transient command scope also protects downstream diagnostics.
Public Function DiagnosticText(ByVal original As String) As String
    If mWriting Then
        DiagnosticText = "Warehouse command diagnostic."
    Else
        DiagnosticText = original
    End If
End Function

Public Function CreateWarehouse(ByVal context As String, ByVal warehouseId As String, _
    ByVal warehouseName As String, ByVal stationId As String, ByVal adminUser As String, _
    ByVal localRoot As String, ByVal publishRoot As String, ByVal purpose As String, _
    ByRef report As String, ByRef operatorPath As String, Optional ByRef outcome As String = "") As Boolean
    On Error GoTo Failed
    report = "Another warehouse command is still running.": operatorPath = "": outcome = "FAILED"
    If mWriting Then Exit Function
    mWriting = True
    If Not Authorize(context, report) Then outcome = "DENIED": GoTo Finished
    If Not modWarehouseBootstrap.BootstrapWarehouseLocalValues(warehouseId, warehouseName, _
            stationId, adminUser, localRoot, publishRoot, purpose) Then
        report = "Warehouse creation could not finish. No rollback is claimed.": GoTo Finished
    End If
    operatorPath = modWarehouseBootstrap.GetLastWarehouseOperatorWorkbookPath()
    CreateWarehouse = True: report = "Warehouse hub creation complete.": outcome = "CREATED_LOCAL"
Finished:
    mWriting = False
    Exit Function
Failed:
    report = "Warehouse creation could not finish. No rollback is claimed.": outcome = "FAILED"
    Resume Finished
End Function

' Both same-instance retry and explicit reopened publication use saved identity.
Public Function PublishExisting(ByVal context As String, ByVal localRoot As String, _
    ByVal publishRoot As String, ByRef report As String, Optional ByVal expectedId As String = "", _
    Optional ByRef outcome As String = "") As Boolean
    Dim spec As WarehouseSpec
    On Error GoTo Failed
    report = "Another warehouse command is still running.": outcome = "FAILED"
    If mWriting Then Exit Function
    mWriting = True
    If Not Authorize(context, report) Then outcome = "DENIED": GoTo Finished
    report = "Select an existing valid warehouse hub and a publication destination.": outcome = "REJECTED"
    If Trim$(publishRoot) = "" Then GoTo Finished
    If Not modWarehousePublishSource.ReadExisting(localRoot, spec) Then GoTo Finished
    If expectedId <> "" And StrComp(expectedId, spec.WarehouseId, vbBinaryCompare) <> 0 Then GoTo Finished
    spec.PathSharePoint = publishRoot
    ' Selection/validation cannot extend the life of the captured authorization.
    If Not Authorize(context, report) Then outcome = "DENIED": GoTo Finished
    outcome = "FAILED"
    If Not modWarehouseBootstrap.PublishInitialArtifacts(spec) Then
        report = "Initial publication could not finish. The existing warehouse remains created.": GoTo Finished
    End If
    PublishExisting = True: report = "Initial publication complete.": outcome = "PUBLISHED"
Finished:
    mWriting = False
    Exit Function
Failed:
    report = "Initial publication could not finish. No rollback is claimed.": outcome = "FAILED"
    Resume Finished
End Function

Private Function Authorize(ByVal context As String, ByRef report As String) As Boolean
    Dim target As WarehouseTarget, wb As Workbook, actor As String
    On Error GoTo Denied
    report = "The invSys session, warehouse or Admin rights changed. Reopen Create Warehouse."
    If context = "" Or context <> modAuthUserCommands.CaptureContext() Then Exit Function
    Set target = modNasConnection.GetCurrentTarget()
    If target Is Nothing Then Exit Function
    ' A dirty borrowed authority cannot provide a new persisted permission grant.
    For Each wb In Application.Workbooks
        If StrComp(wb.FullName, target.AuthPath, vbTextCompare) = 0 Then
            If Not wb.Saved Then Exit Function
        End If
    Next wb
    If Not modAuth.LoadAuth(target.WarehouseId) Then Exit Function
    actor = modAuth.GetCurrentUserId()
    If Not modAuth.CanPerform("ADMIN_MAINT", actor, target.WarehouseId, target.StationId, "CREATE_WAREHOUSE") Then Exit Function
    If context <> modAuthUserCommands.CaptureContext() Then Exit Function
    Authorize = True: report = ""
Denied:
End Function
