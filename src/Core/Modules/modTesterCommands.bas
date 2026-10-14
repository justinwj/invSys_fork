Attribute VB_Name = "modTesterCommands"
Option Explicit

Private mWriting As Boolean
Private mContext As String
Private mReceipt As Object
Private mOwned As Collection
Private mBorrowed As Collection
Private mDenied As Boolean
Private mCleaning As Boolean
Private mExpectedRoot As String

Public Function CaptureForm(Optional ByVal workbookName As String = "") As String
    CaptureForm = modTesterAuthority.Capture(workbookName)
End Function

Public Sub ReleaseForm(ByVal context As String)
    modTesterAuthority.Release context
End Sub

Public Function CheckForm(ByVal context As String, ByRef report As String) As Boolean
    CheckForm = modTesterAuthority.Check(context)
    If Not CheckForm Then report = "The invSys session, source warehouse or Admin rights changed. Reopen Test Environment Setup."
End Function

Public Function IsWriting() As Boolean
    IsWriting = mWriting
End Function

Public Function DiagnosticText(ByVal original As String) As String
    If mWriting Then DiagnosticText = "Tester command diagnostic." Else DiagnosticText = original
End Function

Public Function Setup(ByVal context As String, ByVal userId As String, ByVal credential As String, _
    ByVal warehouse As String, ByVal station As String, ByVal root As String, ByVal publishRoot As String, _
    ByRef report As String, ByRef operatorPath As String, Optional ByRef outcome As String = "") As Boolean
    Dim spec As TesterProvisionSpec, events As Boolean, alerts As Boolean
    On Error GoTo Failed
    outcome = "REJECTED": operatorPath = ""
    report = "Another Tester command is still running."
    If mWriting Then Exit Function
    BeginCommand context, events, alerts
    If Not CheckForm(context, report) Then outcome = "DENIED": GoTo Finished
    spec.UserId = Trim$(userId): spec.PinHash = credential
    spec.WarehouseId = Trim$(warehouse): spec.StationId = Trim$(station)
    spec.PathLocal = modTesterOwnership.CanonicalRoot(root)
    spec.PathSharePointRoot = modTesterOwnership.CanonicalRoot(publishRoot)
    report = "The destination is unavailable, conflicting or not verifiably owned. Setup did not complete."
    If spec.UserId = "" Or credential = "" Or spec.PathLocal = "" Then GoTo Finished
    If publishRoot <> "" And spec.PathSharePointRoot = "" Then GoTo Finished
    If Not FilesAvailable(modTesterOwnership.Artifacts(spec)) Then GoTo Finished
    Set mReceipt = modTesterOwnership.Resolve(context, spec, True)
    If mReceipt Is Nothing Then GoTo Finished
    outcome = "FAILED": report = "Setup could not finish. Ownership is retained for retry. No rollback is claimed."
    If Not CheckPhase() Then GoTo Finished
    If Not modTesterProvisioning.SetupTesterStation(spec) Then GoTo Finished
    If Not CheckPhase() Then GoTo Finished
    operatorPath = spec.PathLocal & "\" & spec.WarehouseId & ".Receiving.Operator.xlsm"
    Setup = True: outcome = "COMPLETED": report = "Setup complete. Open the operator workbook to begin."
Finished:
    If mDenied Then outcome = "DENIED": report = "Authority or destination scope changed. Work stopped; ownership is retained for retry."
    EndCommand events, alerts
    Exit Function
Failed:
    outcome = "FAILED": report = "Setup could not finish. No rollback is claimed."
    Resume Finished
End Function

Public Function DeleteGenerated(ByVal context As String, ByVal warehouse As String, ByVal station As String, _
    ByVal root As String, ByVal publishRoot As String, ByVal confirmed As Boolean, _
    ByRef report As String, Optional ByRef outcome As String = "") As Boolean
    Dim spec As TesterProvisionSpec, events As Boolean, alerts As Boolean, item As Variant, fso As Object
    On Error GoTo Failed
    outcome = "REJECTED": report = "Another Tester command is still running."
    If mWriting Then Exit Function
    BeginCommand context, events, alerts
    mCleaning = True
    If Not CheckForm(context, report) Then outcome = "DENIED": GoTo Finished
    If Not confirmed Then outcome = "CANCELLED": report = "Cleanup cancelled.": GoTo Finished
    spec.WarehouseId = warehouse: spec.StationId = station
    spec.PathLocal = root: spec.PathSharePointRoot = publishRoot
    report = "Cleanup refused: the destination is not verifiably owned or an artifact is open in Excel."
    Set mReceipt = modTesterOwnership.Resolve(context, spec, False)
    If mReceipt Is Nothing Then GoTo Finished
    If Not FilesAvailable(mReceipt("Artifacts")) Then GoTo Finished
    Set fso = CreateObject("Scripting.FileSystemObject")
    outcome = "FAILED": report = "Cleanup could not finish. Ownership is retained for retry. No rollback is claimed."
    For Each item In mReceipt("Artifacts")
        If Not CheckPhase() Then GoTo Finished
        If fso.FileExists(CStr(item)) Then fso.DeleteFile CStr(item), False
    Next item
    For Each item In mReceipt("Directories")
        If Not CheckPhase() Then GoTo Finished
        If fso.FolderExists(CStr(item)) Then
            If fso.GetFolder(CStr(item)).Files.Count = 0 And fso.GetFolder(CStr(item)).SubFolders.Count = 0 Then RmDir CStr(item)
        End If
    Next item
    ' Verification has its own authority boundary and performs no repair or save.
    If Not CheckPhase() Then GoTo Finished
    For Each item In mReceipt("Artifacts")
        If fso.FileExists(CStr(item)) Then GoTo Finished
    Next item
    If Not CheckPhase() Then GoTo Finished
    If spec.PathSharePointRoot <> "" Then
        If fso.FileExists(spec.PathSharePointRoot & "\" & modTesterOwnership.RECEIPT_NAME) Then _
            fso.DeleteFile spec.PathSharePointRoot & "\" & modTesterOwnership.RECEIPT_NAME, False
    End If
    If Not CheckPhase() Then GoTo Finished
    fso.DeleteFile spec.PathLocal & "\" & modTesterOwnership.RECEIPT_NAME, False
    DeleteGenerated = True: outcome = "COMPLETED": report = "Test Environment generated files deleted. Unknown files were preserved."
Finished:
    If mDenied Then outcome = "DENIED": report = "Authority or destination scope changed. Cleanup stopped; ownership is retained for retry."
    EndCommand events, alerts
    Exit Function
Failed:
    outcome = "FAILED": report = "Cleanup could not finish. Ownership is retained for retry. No rollback is claimed."
    Resume Finished
End Function

' Called by the shared provisioning owner at each write phase. Fixtures stay separate.
Public Function CheckPhase() As Boolean
    On Error GoTo Denied
    If Not mWriting Then CheckPhase = True: Exit Function
    If mDenied Then Exit Function
    If StrComp(NormalizeFolderPathTesterSetup(modRuntimeWorkbooks.GetCoreDataRootOverride()), mExpectedRoot, vbTextCompare) <> 0 Then GoTo Denied
    If Not modTesterAuthority.Check(mContext, mExpectedRoot) Then GoTo Denied
    If mReceipt Is Nothing Then GoTo Denied
    If Not modTesterOwnership.Unchanged(mReceipt, mCleaning) Then GoTo Denied
    If Not FilesAvailable(mReceipt("Artifacts")) Then GoTo Denied
    CheckPhase = True
    Exit Function
Denied:
    mDenied = True
End Function

Public Function EnterProvisioningRoot(ByVal root As String) As Boolean
    If mWriting Then
        If Not CheckPhase() Then Exit Function
        If root <> mReceipt("Root") Then Exit Function
        mExpectedRoot = NormalizeFolderPathTesterSetup(root)
    End If
    modRuntimeWorkbooks.SetCoreDataRootOverride root
    EnterProvisioningRoot = True
End Function

Public Sub LeaveProvisioningRoot(ByVal priorRoot As String)
    If mWriting Then
        If StrComp(NormalizeFolderPathTesterSetup(modRuntimeWorkbooks.GetCoreDataRootOverride()), mExpectedRoot, vbTextCompare) <> 0 Then
            mDenied = True
            Exit Sub
        End If
        mExpectedRoot = NormalizeFolderPathTesterSetup(priorRoot)
    End If
    RestoreCoreRootOverrideTesterSetup priorRoot
End Sub

Public Sub RegisterOwned(ByVal wb As Workbook)
    Dim item As Workbook
    If Not mWriting Or wb Is Nothing Then Exit Sub
    For Each item In mBorrowed
        If item Is wb Then Exit Sub
    Next item
    For Each item In mOwned
        If item Is wb Then Exit Sub
    Next item
    mOwned.Add wb
End Sub

Private Function FilesAvailable(ByVal paths As Collection) As Boolean
    Dim wb As Workbook, item As Variant, owned As Workbook, mayUse As Boolean
    For Each wb In Application.Workbooks
        For Each item In paths
            If StrComp(wb.FullName, CStr(item), vbTextCompare) = 0 Then
                mayUse = False
                If Not mOwned Is Nothing Then
                    For Each owned In mOwned
                        If wb Is owned Then mayUse = True
                    Next owned
                End If
                If Not mayUse Then Exit Function
            End If
        Next item
    Next wb
    FilesAvailable = True
End Function

Private Sub BeginCommand(ByVal context As String, ByRef events As Boolean, ByRef alerts As Boolean)
    Dim wb As Workbook
    events = Application.EnableEvents: alerts = Application.DisplayAlerts
    mWriting = True: mContext = context: mDenied = False: mCleaning = False
    mExpectedRoot = NormalizeFolderPathTesterSetup(modRuntimeWorkbooks.GetCoreDataRootOverride())
    Set mOwned = New Collection: Set mBorrowed = New Collection: Set mReceipt = Nothing
    For Each wb In Application.Workbooks: mBorrowed.Add wb: Next wb
    Application.EnableEvents = False: Application.DisplayAlerts = False
End Sub

Private Sub EndCommand(ByVal events As Boolean, ByVal alerts As Boolean)
    Dim wb As Workbook, candidate As Workbook
    On Error Resume Next
    For Each wb In mOwned
        For Each candidate In Application.Workbooks
            If candidate Is wb Then candidate.Close SaveChanges:=False: Exit For
        Next candidate
    Next wb
    Application.EnableEvents = events: Application.DisplayAlerts = alerts
    Set mOwned = Nothing: Set mBorrowed = Nothing: Set mReceipt = Nothing
    mContext = "": mWriting = False
    On Error GoTo 0
End Sub
