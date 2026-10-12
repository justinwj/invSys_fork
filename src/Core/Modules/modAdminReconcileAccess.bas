Attribute VB_Name = "modAdminReconcileAccess"
Option Explicit

' D8-RECON-01: transient interactive binding; never persisted or used by processors.
Public Function CaptureContext() As String
    On Error GoTo Unavailable
    Dim target As WarehouseTarget
    If Not modAuth.IsSignedIn() Or Not modNasConnection.IsCurrentTargetAllowed(False) Then Exit Function
    Set target = modNasConnection.GetCurrentTarget()
    CaptureContext = modTrainingJson.Quote(target.WarehouseId) & modTrainingJson.Quote(target.StationId) & _
        modTrainingJson.Quote(target.RuntimeRoot) & modTrainingJson.Quote(target.ConfigPath) & _
        modTrainingJson.Quote(target.HubRoot) & CStr(target.SourceType) & _
        modTrainingJson.Quote(modAuth.GetCurrentUserId()) & CStr(modAuthSession.Version())
Unavailable:
End Function

Public Function ValidateSubmission(ByVal context As String, ByVal warehouseId As String, _
                                   ByVal stationId As String, ByVal runtimeRoot As String, _
                                   ByVal userId As String, ByRef report As String) As Boolean
    On Error GoTo Changed
    Dim target As WarehouseTarget
    report = "Current invSys user is not signed in."
    If Not modAuth.IsSignedIn() Then Exit Function
    If context = "" Or context <> CaptureContext() Then GoTo Changed
    Set target = modNasConnection.GetCurrentTarget()
    If StrComp(warehouseId, target.WarehouseId, vbBinaryCompare) <> 0 Then GoTo Changed
    If StrComp(stationId, target.StationId, vbBinaryCompare) <> 0 Then GoTo Changed
    If StrComp(runtimeRoot, target.RuntimeRoot, vbBinaryCompare) <> 0 Then GoTo Changed
    If StrComp(userId, modAuth.GetCurrentUserId(), vbBinaryCompare) <> 0 Then GoTo Changed
    report = "Current user lacks ADMIN_MAINT capability."
    If Not modAuth.CanPerform("ADMIN_MAINT", userId, warehouseId, stationId, "SHIPMENT_RECONCILE") Then Exit Function
    If context <> CaptureContext() Then GoTo Changed
    report = ""
    ValidateSubmission = True
    Exit Function
Changed:
    report = "The invSys session or warehouse changed. Reopen Shipment Reconcile."
End Function
