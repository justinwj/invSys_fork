# D18 observations of the existing five native workflows. No activity is fabricated.
. (Join-Path $PSScriptRoot 'Slice4beActivityAssertions.ps1')
. (Join-Path $PSScriptRoot 'Slice4beAdminReconcileGuide.ps1')
function Install-AdminReconcileRecordingProbe {
    $module=$packages['invSys.Core.xlam'].VBProject.VBComponents.Add(1)
    $module.Name='TestReconcileRecording'
    $module.CodeModule.AddFromString(@'
Option Explicit
Private mEntrySecret As String, mEntryChanged As Boolean
Private mFaultMode As String, mFaultCalls As Long
Public Sub ArmStoreFailure(ByVal mode As String)
    mFaultMode = mode: mFaultCalls = 0
End Sub
Public Sub StoreWritePath(ByVal body As String, ByRef pending As String)
    Dim record As Object
    If mFaultMode = "" Then Exit Sub
    Set record = modTrainingJson.DecodeObject(body)
    If Left$(CStr(record("ControlId")), 16) <> "ADMIN_RECONCILE_" Then Exit Sub
    If CStr(record("OutcomeCode")) <> mFaultMode Then Exit Sub
    mFaultCalls = mFaultCalls + 1
    ' The real Open statement fails on a nonexistent directory beneath the generated path.
    pending = pending & "\unavailable.pending"
End Sub
Public Function StoreFaultCalls() As Long
    StoreFaultCalls = mFaultCalls
End Function
Public Sub ArmEntryChange(ByVal secret As String)
    mEntrySecret = secret: mEntryChanged = False
End Sub
Public Sub ChangeAtEntry(ByVal controlId As String)
    Dim secret As String, status As String, version As Long
    If controlId <> "ADMIN_RECONCILE_OPEN" Or mEntrySecret = "" Then Exit Sub
    secret = mEntrySecret: mEntrySecret = ""
    version = modAuthSession.Version()
    modAuth.SignOut
    status = modAuth.SignInCurrentTargetForAutomation("config-admin", secret, "")
    secret = ""
    mEntryChanged = (Left$(status, 3) = "OK|" And modAuth.IsSignedIn() And modAuthSession.Version() <> version)
End Sub
Public Function EntryChanged() As Boolean
    EntryChanged = mEntryChanged
End Function
Public Function EnablePolicy(Optional ByVal mode As String = "Enabled") As Boolean
    Dim context As String, version As Long, request As String, report As String
    Dim model As Object, row As Variant, users As Collection, user As Object
    context = modActivity.CaptureContext()
    If context = "" Or Not modNasConnection.IsCurrentTargetAllowed(True) Then Exit Function
    If Not modTrackingPolicySettings.ReadEditor(context, version, request, report) Then Exit Function
    Set model = modTrackingPolicyModel.Defaults(mode <> "Disabled")
    model("ViewerActionPathCaptureEnabled") = (mode = "Capture")
    For Each row In model("Controls")
        row("Collect") = (mode <> "Disabled") And Not (mode = "ControlsOff" And Left$(row("ControlId"), 16) = "ADMIN_RECONCILE_")
    Next row
    If mode = "UserOff" Then
        Set user = CreateObject("Scripting.Dictionary")
        user.Add "UserId", "config-admin": user.Add "Record", False
        Set users = model("Users"): users.Add user
    End If
    request = modTrainingJson.EncodeObject(model)
    EnablePolicy = modTrackingPolicySettings.SavePolicy(context, version, request, report)
End Function
Public Function Definition(ByVal id As String, Optional ByVal version As Long = 0) As String
    Dim value As Object
    If version = 0 Then version = modActivityCatalog.CATALOG_VERSION
    Set value = modActivityCatalog.Control(id, version)
    If Not value Is Nothing Then Definition = modTrainingJson.EncodeObject(value)
End Function
Public Function CurrentCatalogVersion() As Long
    CurrentCatalogVersion = modActivityCatalog.CATALOG_VERSION
End Function
Public Function PriorDefinitionsPreserved() As Boolean
    Dim id As Variant, a As Object, b As Object, ids As Variant
    ids = modActivityCatalog.ControlIds(33)
    If UBound(ids) - LBound(ids) + 1 <> 151 Then Exit Function
    For Each id In ids
        Set a = modActivityCatalog.Control(CStr(id), 33)
        Set b = modActivityCatalog.Control(CStr(id))
        If a Is Nothing Or b Is Nothing Then Exit Function
        If modTrainingJson.EncodeObject(a) <> modTrainingJson.EncodeObject(b) Then Exit Function
    Next id
    PriorDefinitionsPreserved = True
End Function
Public Function Concludes(ByVal wire As String) As Boolean
    Concludes = modEvaluationMatches.CommandCompleted(modTrainingJson.DecodeObject(wire))
End Function
Public Function ReferencesAccepted(ByVal warehouse As String, ByVal controlId As String, ByVal outcome As String, ByVal wire As String) As Boolean
    Dim refs As Collection
    Set refs = modActivityReferences.Decode(warehouse, controlId, outcome, wire)
    ReferencesAccepted = Not refs Is Nothing
End Function
'@)
    $activity=$packages['invSys.Core.xlam'].VBProject.VBComponents.Item('modActivity').CodeModule
    $line=$activity.ProcBodyLine('BeginAction',0)
    while($activity.Lines($line,1).Trim() -cne 'notice = ""'){$line++;if($line -gt $activity.CountOfLines){throw 'Recording entry seam missing'}}
    $activity.InsertLines($line,'    TestReconcileRecording.ChangeAtEntry controlId')
    $store=$packages['invSys.Core.xlam'].VBProject.VBComponents.Item('modActivityStore').CodeModule
    $line=$store.ProcBodyLine('Append',0)
    while($store.Lines($line,1).Trim() -cne 'Open pending For Binary Access Write Lock Read Write As #file'){$line++;if($line -gt $store.CountOfLines){throw 'Store write seam missing'}}
    $store.InsertLines($line,'    TestReconcileRecording.StoreWritePath body, pending')
    $admin=$packages['invSys.Admin.xlam'].VBProject.VBComponents.Item('TestReconcileIdentity').CodeModule
    $admin.InsertLines($admin.CountOfDeclarationLines+1,'Private mLastTrackingNotice As String')
    $line=$admin.ProcBodyLine('ObserveNotice',0)
    while($admin.Lines($line,1).Trim() -cne 'mTrackingNotices = mTrackingNotices + 1'){$line++;if($line -gt $admin.CountOfLines){throw 'Tracking notice seam missing'}}
    $admin.InsertLines($line+1,'        mLastTrackingNotice = text')
    $admin.AddFromString(@'
Public Function ExactStoreWarning() As Boolean
    ExactStoreWarning = (mTrackingNotices = 1 And mLastTrackingNotice = "Tracking unavailable: the training record could not be saved.")
End Function
Public Function RecordingContextRefused() As Boolean
    RecordingContextRefused = (mBoundaryNotice = "The invSys session or warehouse changed. Reopen Shipment Reconcile." And Not mQueuedNotice)
End Function
Public Function RecordingSubmission() As Boolean
    mBoundaryCase = "": mSelection = 1
    RecordingSubmission = InvokeCallback()
End Function
'@)
    if($AdminReconcileGuidePaths){
        Install-ReconcileGuideProbe
    }
}

function Test-AdminReconcileRecording($Fixture,[string[]]$Before,$Events,$State) {
    function Probe([string]$Name,[object[]]$Values=@()){Run 'invSys.Core.xlam' ('TestReconcileRecording.'+$Name) $Values}
    $ids=@('ADMIN_RECONCILE_OPEN','ADMIN_RECONCILE_INSPECT','ADMIN_RECONCILE_CANCEL')
    Check 'Reconcile.Recording.Prior151DefinitionsPreserved' ([bool](Probe 'PriorDefinitionsPreserved'))
    foreach($id in $ids){
        $wire=[string](Probe 'Definition' @($id));$definition=if($wire){$wire|ConvertFrom-Json}else{$null}
        $class=if($id -ceq 'ADMIN_RECONCILE_INSPECT'){'Command'}else{'Navigation'}
        Check ('Reconcile.Recording.Catalog.'+$id) ($null -ne $definition -and $definition.OwnerId -ceq 'ADMIN_SHIPMENT_RECONCILE' -and $definition.Class -ceq $class -and $definition.Role -ceq 'Admin' -and $definition.Capability -ceq 'ADMIN_MAINT')
        Check ('Reconcile.Recording.PriorCatalogExcludes.'+$id) ([string](Probe 'Definition' @($id,33)) -ceq '')
    }
    $records=@(Get-Slice4beActivityFiles $Fixture|Where-Object {$_ -cnotin $Before}|ForEach-Object {Get-Content $_ -Raw|ConvertFrom-Json}|Where-Object {$_.ControlId -cin $ids})
    $attempts=@($records|Where-Object OutcomeCode -CEQ 'REQUESTED');$outcomes=@($records|Where-Object OutcomeCode -CNE 'REQUESTED')
    Check 'Reconcile.Recording.ExactTenActionPairs' ($records.Count -eq 20 -and $attempts.Count -eq 10 -and $outcomes.Count -eq 10)
    $expected=@{'ADMIN_RECONCILE_OPEN|OPENED'=5;'ADMIN_RECONCILE_CANCEL|CANCELLED'=2;'ADMIN_RECONCILE_INSPECT|NO_CORRECTION'=1;'ADMIN_RECONCILE_INSPECT|CANCELLED'=1;'ADMIN_RECONCILE_INSPECT|CONFIRMED'=1}
    $counts=$outcomes.Count -eq 10
    foreach($key in $expected.Keys){$counts=$counts -and @($outcomes|Where-Object {($_.ControlId+'|'+$_.OutcomeCode) -ceq $key}).Count -eq $expected[$key]}
    Check 'Reconcile.Recording.ExactOwnerOutcomes' $counts
    $pairs=$records.Count -eq 20 -and @($records.RecordId|Sort-Object -Unique).Count -eq 20 -and @($attempts.ActivityId|Sort-Object -Unique).Count -eq 10
    foreach($attempt in $attempts){$last=@($outcomes|Where-Object ActivityId -CEQ $attempt.ActivityId);$pairs=$pairs -and $last.Count -eq 1;if($last.Count -eq 1){$pairs=$pairs -and $last[0].ControlId -ceq $attempt.ControlId}}
    Check 'Reconcile.Recording.DistinctCorrelatedIdentities' $pairs
    $context=$records.Count -eq 20;$safe=$context;$conclusion=$context;$effects=$context;$references=$context
    $catalogVersion=[int](Probe 'CurrentCatalogVersion')
    $last=$Events.ListRows.Item($Events.ListRows.Count)
    $newEvent=[string]$last.Range.Cells.Item(1,$Events.ListColumns.Item('EventID').Index).Value2
    foreach($record in $records){
        $wire=ConvertTo-Json $record -Depth 20 -Compress
        $context=$context -and $record.WarehouseId -ceq $Fixture.Warehouse -and $record.StationId -ceq 'S1' -and $record.UserId -ceq 'config-admin' -and $record.OwnerId -ceq 'ADMIN_SHIPMENT_RECONCILE' -and $record.CatalogVersion -eq $catalogVersion
        foreach($value in @($Fixture.Secret,(CredentialHash $Fixture.Secret),$Fixture.Root,$State.KeyA,$State.KeyB,'Generated native training correction','PinHash','OpenShipmentReconcileTool')){$safe=$safe -and -not $wire.Contains($value)}
        $positive=$record.OutcomeCode -cin @('OPENED','CONFIRMED')
        $conclusion=$conclusion -and ([bool](Probe 'Concludes' @($wire)) -eq $positive)
        $effect=if($record.OutcomeCode -cin @('REQUESTED','CONFIRMED')){'Unknown'}else{'Unchanged'}
        $effects=$effects -and $record.DataEffect -ceq $effect -and $record.EventCode -ceq ($record.ControlId+'_'+$record.OutcomeCode)
        $refs=@($record.SourceEventRefs)
        if($record.OutcomeCode -ceq 'CONFIRMED'){
            $references=$references -and $refs.Count -eq 1
            if($refs.Count -eq 1){$references=$references -and $refs[0].EventId -ceq $newEvent -and $refs[0].WarehouseId -ceq $Fixture.Warehouse -and $refs[0].SourceKind -ceq 'Inventory' -and $refs[0].SubmissionState -ceq 'Submitted'}
        }else{$references=$references -and $refs.Count -eq 0}
    }
    Check 'Reconcile.Recording.CapturedContext' $context
    Check 'Reconcile.Recording.NoEnteredValues' $safe
    Check 'Reconcile.Recording.ExactCommandConclusions' $conclusion
    Check 'Reconcile.Recording.OwnerEffects' $effects
    Check 'Reconcile.Recording.OnlyNewCorrectionReference' $references
    Check 'Reconcile.Recording.AllFiveContextChangesObserved' ($script:reconcileContextNotices.Count -eq 5)
    foreach($notice in $script:reconcileContextNotices){Check ('Reconcile.Recording.ContextNotice.'+$notice.Case) ($notice.Count -eq 1)}
    $one=@{WarehouseId=$Fixture.Warehouse;SourceKind='Inventory';EventId=$newEvent;SubmissionState='Submitted'}
    $unknown=@{WarehouseId=$Fixture.Warehouse;SourceKind='Inventory';EventId=$newEvent;SubmissionState='Unknown'}
    $other=@{WarehouseId=$Fixture.Warehouse;SourceKind='Inventory';EventId=$State.ShipId;SubmissionState='Submitted'}
    $otherUnknown=@{WarehouseId=$Fixture.Warehouse;SourceKind='Inventory';EventId=$State.ShipId;SubmissionState='Unknown'}
    $guards=@(
        @{Name='ConfirmedOne';Code='CONFIRMED';Refs=@($one);Expected=$true},
        @{Name='ConfirmedMissing';Code='CONFIRMED';Refs=@();Expected=$false},
        @{Name='ConfirmedMultiple';Code='CONFIRMED';Refs=@($one,$other);Expected=$false},
        @{Name='FailedUnknown';Code='FAILED';Refs=@($unknown);Expected=$true},
        @{Name='FailedSubmitted';Code='FAILED';Refs=@($one);Expected=$false},
        @{Name='FailedMultiple';Code='FAILED';Refs=@($unknown,$otherUnknown);Expected=$false},
        @{Name='CancelledReference';Code='CANCELLED';Refs=@($one);Expected=$false},
        @{Name='NoCorrectionReference';Code='NO_CORRECTION';Refs=@($one);Expected=$false})
    foreach($guard in $guards){
        $wire=ConvertTo-Json -InputObject @($guard.Refs) -Depth 8 -Compress
        $accepted=[bool](Probe 'ReferencesAccepted' @($Fixture.Warehouse,'ADMIN_RECONCILE_INSPECT',$guard.Code,$wire))
        Check ('Reconcile.Recording.ReferenceGuard.'+$guard.Name) ($accepted -eq $guard.Expected)
    }
    $bytes=Hash $State.InventoryPath
    foreach($mode in @('Disabled','ControlsOff','UserOff')){
        if(-not [bool](Probe 'EnablePolicy' @($mode))){throw 'Suppression policy fixture unavailable; not behavioral RED.'}
        $before=@(Get-Slice4beActivityFiles $Fixture);$rows=$Events.ListRows.Count
        $called=[bool](Run 'invSys.Admin.xlam' 'TestReconcileIdentity.RecordingSubmission')
        Check ('Reconcile.Recording.'+$mode+'.CommandStillQueues') ($called -and $Events.ListRows.Count -eq $rows+1)
        Check ('Reconcile.Recording.'+$mode+'.NoActivity') (@(Get-Slice4beActivityFiles $Fixture).Count -eq $before.Count)
        Check ('Reconcile.Recording.'+$mode+'.NoFalseWarning') ([int](Run 'invSys.Admin.xlam' 'TestReconcileIdentity.TrackingNoticeCount') -eq 0)
        Check ('Reconcile.Recording.'+$mode+'.InventoryPreserved') (((Hash $State.InventoryPath) -ceq $bytes) -and [bool](Run 'invSys.Admin.xlam' 'TestReconcileIdentity.NativeFixturePreserved' @(4)))
    }
    if(-not [bool](Probe 'EnablePolicy')){throw 'Could not restore the generated enabled policy.'}
    $before=@(Get-Slice4beActivityFiles $Fixture);$rows=$Events.ListRows.Count
    Probe 'ArmEntryChange' @($Fixture.Secret)
    $called=[bool](Run 'invSys.Admin.xlam' 'TestReconcileIdentity.RecordingSubmission')
    if(-not [bool](Probe 'EntryChanged')){throw 'Entry reauthentication fixture did not occur; not behavioral RED.'}
    Check 'Reconcile.Recording.EntryChange.NoQueue' (-not $called -and $Events.ListRows.Count -eq $rows)
    Check 'Reconcile.Recording.EntryChange.OriginalContextRefused' ([bool](Run 'invSys.Admin.xlam' 'TestReconcileIdentity.RecordingContextRefused'))
    Check 'Reconcile.Recording.EntryChange.NoLaterAttribution' (@(Get-Slice4beActivityFiles $Fixture).Count -eq $before.Count)
    Check 'Reconcile.Recording.EntryChange.InventoryPreserved' (((Hash $State.InventoryPath) -ceq $bytes) -and [bool](Run 'invSys.Admin.xlam' 'TestReconcileIdentity.NativeFixturePreserved' @(4)))
    foreach($mode in @('REQUESTED','CONFIRMED')){
        $before=@(Get-Slice4beActivityFiles $Fixture);$prior=@{};foreach($path in $before){$prior[$path]=Hash $path}
        $rows=$Events.ListRows.Count
        Probe 'ArmStoreFailure' @($mode)
        try{
            $queued=[bool](Run 'invSys.Admin.xlam' 'TestReconcileIdentity.RecordingSubmission')
            $faults=[int](Probe 'StoreFaultCalls')
        }finally{Probe 'ArmStoreFailure' @('')}
        if($faults -eq 0){throw 'Store fault did not reach the real write boundary; not behavioral RED.'}
        Check ('Reconcile.Recording.StoreFailure.'+$mode+'.NoCommandRetry') ($queued -and $Events.ListRows.Count -eq $rows+1 -and $faults -eq $(if($mode -ceq 'REQUESTED'){2}else{1}))
        Check ('Reconcile.Recording.StoreFailure.'+$mode+'.ExactWarningOnce') ([bool](Run 'invSys.Admin.xlam' 'TestReconcileIdentity.ExactStoreWarning'))
        $new=@(Get-Slice4beActivityFiles $Fixture|Where-Object {$_ -cnotin $before}|ForEach-Object {Get-Content $_ -Raw|ConvertFrom-Json})
        $honest=if($mode -ceq 'REQUESTED'){$new.Count -eq 0}else{
            $new.Count -eq 2 -and @($new|Where-Object {$_.ControlId -ceq 'ADMIN_RECONCILE_OPEN' -and $_.OutcomeCode -ceq 'REQUESTED'}).Count -eq 1 -and
            @($new|Where-Object {$_.ControlId -ceq 'ADMIN_RECONCILE_INSPECT' -and $_.OutcomeCode -ceq 'REQUESTED'}).Count -eq 1 -and @($new|Where-Object {@($_.SourceEventRefs).Count -ne 0}).Count -eq 0
        }
        Check ('Reconcile.Recording.StoreFailure.'+$mode+'.NoInventedCompletion') $honest
        Check ('Reconcile.Recording.StoreFailure.'+$mode+'.PriorRecordsImmutable') (@($prior.Keys|Where-Object {(Hash $_) -cne $prior[$_]}).Count -eq 0)
        Check ('Reconcile.Recording.StoreFailure.'+$mode+'.InventoryPreserved') (((Hash $State.InventoryPath) -ceq $bytes) -and [bool](Run 'invSys.Admin.xlam' 'TestReconcileIdentity.NativeFixturePreserved' @(4)))
    }
    if($AdminReconcileGuidePaths){Test-ReconcileGuide $Fixture $Events $State}
}
