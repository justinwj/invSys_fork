# Read-only authority and cancellation acceptance through the packaged callback.
# All mutations below prepare/restore disposable owner-generated fixtures only.
function Install-AdminReconcileAuthorityProbe($Module) {
    $Module.InsertLines($Module.CountOfDeclarationLines+1,@'
Private mBoundaryCase As String
Private mBoundaryPath As String
Private mBoundarySnapshot As String
Private mBoundaryNotice As String
Private mMutationSecret As String, mOtherRoot As String, mOtherSecret As String
Private mMutationObserved As Boolean
'@)
    $Module.AddFromString(@'
Public Function PrepareAuthoritySuite() As Boolean
    Dim item As Object
    Set item = HistoryItem(mKeyA, 3, "CURRENT-A", "DAMAGED")
    If Not ApplyFixtureEvent("ADMIN_INVENTORY_ADJUST", modRoleEventWriter.CreateSystemKey(), item) Then Exit Function
    mSelection = 1
    PrepareAuthoritySuite = (EntityQty(mKeyA) = 21 And EntityQty(mKeyB) = 55 And StageCanary())
    mBoundaryPath = mInventory.FullName
End Function

Public Function PrepareAuthorityCase(ByVal scenario As String) As Boolean
    Dim lo As ListObject, r As Long
    mBoundaryCase = scenario
    mMutationObserved = False
    If Not mInventory.Saved Then Err.Raise 5, , "Authority fixture baseline must be saved"
    Set lo = FindTable(mInventory, "tblInventoryEntities")
    Select Case scenario
        Case "BorrowedDirty"
            lo.DataBodyRange.Cells(1, lo.ListColumns("Operator_Identity_Canary").Index).Value2 = "Unsaved generated annotation"
        Case "InvalidSchema"
            lo.ListColumns("QtyOnHand").Name = "Fixture_MissingQuantity"
        Case "AmbiguousKey"
            For r = 1 To lo.ListRows.Count
                If CStr(Value(lo, r, "System_Key")) = mKeyB Then lo.DataBodyRange.Cells(r, lo.ListColumns("System_Key").Index).Value2 = mKeyA
            Next r
        Case "Closed", "Missing"
            mInventory.Close False
            Set mInventory = Nothing
            If scenario = "Missing" Then
                If Len(Dir$(mBoundaryPath & ".reconcile-held")) <> 0 Then Err.Raise 5, , "Held fixture already exists"
                Name mBoundaryPath As mBoundaryPath & ".reconcile-held"
            End If
        Case "BorrowedClean", "CancelSelection", "CancelNarrative", "InvalidSelection", "ReaderDenied", "SignedOut", "SignedOutStaleIdentity", "LocalSignedIn"
        Case "SignOutDuringSelection", "SignOutDuringNarrative", "ReauthenticateDuringNarrative", "SwitchUserDuringNarrative", "SwitchTargetDuringNarrative"
        Case Else: Err.Raise 5, , "Unknown authority fixture"
    End Select
    mBoundarySnapshot = AuthorityMemorySignature()
    PrepareAuthorityCase = True
End Function

Public Function AuthoritySelectionAnswer() As String
    If mBoundaryCase = "SignOutDuringSelection" Then MutateAuthoritySession
    Select Case mBoundaryCase
        Case "CancelSelection": AuthoritySelectionAnswer = ""
        Case "InvalidSelection": AuthoritySelectionAnswer = "1.5"
        Case Else: AuthoritySelectionAnswer = CStr(mSelection)
    End Select
End Function

Public Function AuthorityNarrativeAnswer() As String
    If mBoundaryCase = "" Or mBoundaryCase = "ReaderDenied" Or mBoundaryCase = "SignedOut" Or mBoundaryCase = "SignedOutStaleIdentity" Then AuthorityNarrativeAnswer = "Generated training identity correction"
    If mBoundaryCase = "LocalSignedIn" Then AuthorityNarrativeAnswer = "Generated training local correction"
    If InStr(1, mBoundaryCase, "During", vbBinaryCompare) > 0 Then
        If mBoundaryCase <> "SignOutDuringSelection" Then MutateAuthoritySession
        AuthorityNarrativeAnswer = "Generated training session-change correction"
    End If
End Function

Public Sub ConfigureAuthorityMutations(ByVal secret As String, ByVal otherRoot As String, ByVal otherSecret As String)
    mMutationSecret = secret: mOtherRoot = otherRoot: mOtherSecret = otherSecret
End Sub

Private Sub MutateAuthoritySession()
    Dim status As String, target As WarehouseTarget
    If Not modAuth.IsSignedIn() Then Err.Raise 5, , "Mutation requires the original live session"
    Select Case mBoundaryCase
        Case "SignOutDuringSelection", "SignOutDuringNarrative"
            modAuth.SignOut
            mMutationObserved = Not modAuth.IsSignedIn()
        Case "ReauthenticateDuringNarrative"
            modAuth.SignOut
            status = modAuth.SignInCurrentTargetForAutomation("config-admin", mMutationSecret, "")
            mMutationObserved = (Left$(status, 3) = "OK|" And modAuth.IsSignedIn() And modAuth.GetCurrentUserId() = "config-admin")
        Case "SwitchUserDuringNarrative"
            status = modAuth.SignInCurrentTargetForAutomation("config-reader", mMutationSecret, "")
            mMutationObserved = (Left$(status, 3) = "OK|" And modAuth.IsSignedIn() And modAuth.GetCurrentUserId() = "config-reader")
        Case "SwitchTargetDuringNarrative"
            modRuntimeWorkbooks.SetCoreDataRootOverride mOtherRoot
            status = modNasConnection.SelectWarehouseTargetForAutomation(mOtherRoot, mOtherRoot, "S1", False)
            If Left$(status, 3) <> "OK|" Then Err.Raise 5, , "Other generated target unavailable"
            status = modAuth.SignInCurrentTargetForAutomation("config-admin", mOtherSecret, "")
            Set target = modNasConnection.GetCurrentTarget()
            mMutationObserved = (Left$(status, 3) = "OK|" And modAuth.IsSignedIn() And target.WarehouseId <> mWarehouse)
    End Select
    If Not mMutationObserved Then Err.Raise 5, , "Live authority mutation did not occur"
End Sub

Public Function AuthorityMutationObserved() As Boolean
    AuthorityMutationObserved = mMutationObserved
End Function

Public Function AuthorityObservation() As String
    Dim expectedPrompts As Long, wb As Workbook, openCount As Long, memory As Boolean, saved As Boolean
    Dim noticeMatches As Boolean
    expectedPrompts = 2
    Select Case mBoundaryCase
        Case "Missing": expectedPrompts = 0
        Case "InvalidSchema", "AmbiguousKey", "CancelSelection", "InvalidSelection": expectedPrompts = 1
    End Select
    For Each wb In Application.Workbooks
        If StrComp(wb.FullName, mBoundaryPath, vbTextCompare) = 0 Then openCount = openCount + 1
    Next wb
    If mBoundaryCase = "Closed" Or mBoundaryCase = "Missing" Then
        memory = (openCount = 0): saved = memory
    Else
        memory = (openCount = 1 And AuthorityMemorySignature() = mBoundarySnapshot)
        saved = mInventory.Saved
        If mBoundaryCase = "BorrowedDirty" Or mBoundaryCase = "InvalidSchema" Or mBoundaryCase = "AmbiguousKey" Then saved = Not saved
    End If
    Select Case mBoundaryCase
        Case "Missing": noticeMatches = (mBoundaryNotice = "No valid server shipment deduction entries are available in the selected warehouse.")
        Case "InvalidSchema", "AmbiguousKey": noticeMatches = (mBoundaryNotice = "The selected shipment and exact inventory entity could not be validated.")
        Case "CancelSelection": noticeMatches = (mBoundaryNotice = "")
        Case "InvalidSelection": noticeMatches = (InStr(1, mBoundaryNotice, "Enter a whole number from 1 to ", vbBinaryCompare) = 1)
        Case "ReaderDenied": noticeMatches = (InStr(1, mBoundaryNotice, "Current user lacks ADMIN_MAINT capability.", vbBinaryCompare) = 1)
        Case "SignedOut", "SignedOutStaleIdentity", "SignOutDuringSelection", "SignOutDuringNarrative": noticeMatches = (mBoundaryNotice = "Current invSys user is not signed in.")
        Case "ReauthenticateDuringNarrative", "SwitchUserDuringNarrative", "SwitchTargetDuringNarrative": noticeMatches = (mBoundaryNotice = "The invSys session or warehouse changed. Reopen Shipment Reconcile.")
        Case Else: noticeMatches = (mBoundaryNotice = "Shipment reconcile cancelled. A human repair narrative is required.")
    End Select
    If mBoundaryCase = "LocalSignedIn" Then
        noticeMatches = mQueuedNotice
    Else
        noticeMatches = noticeMatches And Not mQueuedNotice
    End If
    AuthorityObservation = CStr(mPrompts = expectedPrompts And noticeMatches) & "|" & CStr(memory) & "|" & CStr(saved)
End Function

Private Function AuthorityMemorySignature() As String
    Dim ws As Worksheet, lo As ListObject, cell As Range, text As String, part As String
    If mInventory Is Nothing Then Exit Function
    For Each ws In mInventory.Worksheets
        For Each lo In ws.ListObjects
            text = text & lo.Name & ":" & CStr(lo.ListRows.Count) & ":" & CStr(lo.ListColumns.Count) & "|"
            For Each cell In lo.Range.Cells
                part = CStr(cell.Formula)
                text = text & Len(part) & ":" & part
            Next cell
        Next lo
    Next ws
    AuthorityMemorySignature = text
End Function

Public Sub RestoreAuthorityFixture()
    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If StrComp(wb.FullName, mBoundaryPath, vbTextCompare) = 0 Then wb.Close False: Exit For
    Next wb
    Set mInventory = Nothing
    If mBoundaryCase = "Missing" Then Name mBoundaryPath & ".reconcile-held" As mBoundaryPath
    Set mInventory = Application.Workbooks.Open(mBoundaryPath, UpdateLinks:=0, ReadOnly:=False)
    mBoundaryCase = ""
End Sub
'@)
    $domain=$packages['invSys.Inventory.Domain.xlam'].VBProject
    $probe=$domain.VBComponents.Add(1);$probe.Name='TestReconcileReadLease'
    $probe.CodeModule.AddFromString(@'
Option Explicit
Private mSeen As Long
Private mAllReadOnly As Boolean
Private mLastStateRequest As String
Public Sub Reset()
    mSeen = 0: mAllReadOnly = True: mLastStateRequest = ""
End Sub
Public Sub Observe(ByVal wb As Workbook, ByVal request As String)
    mSeen = mSeen + 1
    mAllReadOnly = mAllReadOnly And wb.ReadOnly
    If Left$(request, 6) = "STATE" & vbTab Then mLastStateRequest = request
End Sub
Public Function ExactStateRequest(ByVal eventId As String, ByVal systemKey As String) As Boolean
    ExactStateRequest = (StrComp(mLastStateRequest, "STATE" & vbTab & eventId & vbTab & systemKey, vbBinaryCompare) = 0)
End Function
Public Function ReadOnlyLeaseObserved() As Boolean
    ReadOnlyLeaseObserved = (mSeen > 0 And mAllReadOnly)
End Function
'@)
    $api=$domain.VBComponents.Item('modInventoryBridgeApi').CodeModule
    $line=$api.ProcBodyLine('ReadShipmentReconcileBridgeResult',0)
    $api.InsertLines($line+1,'    TestReconcileReadLease.Observe inventoryWb, request')
}

function Test-AdminReconcileAuthority($Events,$State,$Fixture,$OtherFixture) {
    $ready=[bool](Run 'invSys.Admin.xlam' 'TestReconcileIdentity.PrepareAuthoritySuite')
    Check 'Reconcile.Authority.OwnerGeneratedBaseline' $ready
    if(-not $ready){throw 'Authority fixture baseline failed; not behavioral RED.'}
    [void](Run 'invSys.Admin.xlam' 'TestReconcileIdentity.ConfigureAuthorityMutations' @($Fixture.Secret,$OtherFixture.Root,$OtherFixture.Secret))
    foreach($case in @('BorrowedClean','BorrowedDirty','InvalidSchema','AmbiguousKey','Closed','Missing','CancelSelection','CancelNarrative','InvalidSelection','ReaderDenied','SignedOut','SignedOutStaleIdentity','SignOutDuringSelection','SignOutDuringNarrative','ReauthenticateDuringNarrative','SwitchUserDuringNarrative','SwitchTargetDuringNarrative','LocalSignedIn')){
        if($case -ceq 'ReaderDenied'){SelectTarget $Fixture 'config-reader'}
        if($case -ceq 'LocalSignedIn'){
            [void](Run 'invSys.Core.xlam' 'modNasConnection.SetCurrentTargetPathsForTest' @($Fixture.Root,$Fixture.Root))
            [void](Run 'invSys.Core.xlam' 'modNasConnection.SetCurrentTargetSourceTypeForTest' @(3))
            $allowed=[bool](Run 'invSys.Core.xlam' 'modNasConnection.IsCurrentTargetAllowed' @($false))
            $nas=[bool](Run 'invSys.Core.xlam' 'modNasConnection.IsCurrentTargetAllowed' @($true))
            Check 'Reconcile.Authority.LocalSignedIn.LocalFallbackObserved' ($allowed -and -not $nas)
            if(-not $allowed -or $nas){throw 'Local fallback fixture unavailable; not behavioral RED.'}
        }
        if($case -in @('SignedOut','SignedOutStaleIdentity')){
            [void](Run 'invSys.Core.xlam' 'modAuth.SignOut')
            if($case -ceq 'SignedOutStaleIdentity'){
                # A provisioned local user-id hint is not an authenticated session.
                [void](Run 'invSys.Core.xlam' 'modRoleEventWriter.SetCurrentUserId' @('config-admin'))
            }else{
                # Pin the denied hint too; never depend on the Windows account name.
                [void](Run 'invSys.Core.xlam' 'modRoleEventWriter.SetCurrentUserId' @('config-reader'))
            }
            if([bool](Run 'invSys.Core.xlam' 'modAuth.IsSignedIn')){throw 'Signed-out fixture precondition failed; not behavioral RED.'}
        }
        $inventoryBefore=Hash $State.InventoryPath
        $otherBefore=Hash $OtherFixture.Config
        $rowsBefore=$Events.ListRows.Count
        try{
            $prepared=[bool](Run 'invSys.Admin.xlam' 'TestReconcileIdentity.PrepareAuthorityCase' @($case))
            if(-not $prepared){throw 'Authority scenario preparation failed; not behavioral RED.'}
            [void](Run 'invSys.Inventory.Domain.xlam' 'TestReconcileReadLease.Reset')
            [void](Run 'invSys.Admin.xlam' 'TestReconcileIdentity.InvokeCallback')
            if($case -like '*During*'){
                $mutated=[bool](Run 'invSys.Admin.xlam' 'TestReconcileIdentity.AuthorityMutationObserved')
                Check ('Reconcile.Authority.'+$case+'.LiveMutationObserved') $mutated
                if(-not $mutated){throw 'Authority mutation failed; not behavioral RED.'}
            }
            $values=([string](Run 'invSys.Admin.xlam' 'TestReconcileIdentity.AuthorityObservation')) -split '\|'
            if($values.Count -ne 3){throw 'Invalid authority observation envelope.'}
            Check ('Reconcile.Authority.'+$case+'.ExpectedCallbackConclusion') ([bool]::Parse($values[0]))
            if($case -ceq 'LocalSignedIn'){
                Check 'Reconcile.Authority.LocalSignedIn.ExactlyOneQueuedEvent' ($Events.ListRows.Count -eq $rowsBefore+1)
            }else{
                Check ('Reconcile.Authority.'+$case+'.NoQueuedEvent') ($Events.ListRows.Count -eq $rowsBefore)
            }
            Check ('Reconcile.Authority.'+$case+'.MemoryAndLifetimePreserved') ([bool]::Parse($values[1]))
            Check ('Reconcile.Authority.'+$case+'.SavedStatePreserved') ([bool]::Parse($values[2]))
            if($case -ceq 'Closed'){
                Check 'Reconcile.Authority.Closed.DomainReceivesReadOnlyLease' ([bool](Run 'invSys.Inventory.Domain.xlam' 'TestReconcileReadLease.ReadOnlyLeaseObserved'))
            }
            if($case -ceq 'Missing'){
                Check 'Reconcile.Authority.Missing.AuthorityNotCreated' (-not (Test-Path -LiteralPath $State.InventoryPath))
                $bytes=Hash ($State.InventoryPath+'.reconcile-held')
            }else{$bytes=Hash $State.InventoryPath}
            Check ('Reconcile.Authority.'+$case+'.FileBytesPreserved') ($bytes -ceq $inventoryBefore)
            Check ('Reconcile.Authority.'+$case+'.OtherWarehousePreserved') ((Hash $OtherFixture.Config) -ceq $otherBefore)
        }finally{
            [void](Run 'invSys.Admin.xlam' 'TestReconcileIdentity.RestoreAuthorityFixture')
            if($case -in @('ReaderDenied','SignedOut','SignedOutStaleIdentity','LocalSignedIn') -or $case -like '*During*'){SelectTarget $Fixture}
        }
    }
    [void](Run 'invSys.Admin.xlam' 'TestReconcileIdentity.ConfigureAuthorityMutations' @('','',''))
}
