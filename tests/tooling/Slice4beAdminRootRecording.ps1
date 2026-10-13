# D18 discovered control: actual callback and confirmed owner facts, no path payloads.
. (Join-Path $PSScriptRoot 'Slice4beActivityAssertions.ps1')
. (Join-Path $PSScriptRoot 'Slice4beAdminRootBoundaries.ps1')
. (Join-Path $PSScriptRoot 'Slice4beAdminRootGuide.ps1')
. (Join-Path $PSScriptRoot 'Slice4beAdminRootScope.ps1')
function Install-AdminRootRecordingProbe {
    $admin=$packages['invSys.Admin.xlam'].VBProject
    $rootProbe=$admin.VBComponents.Item('TestDirectoryColumns').CodeModule
    $rootProbe.InsertLines($rootProbe.CountOfDeclarationLines+1,'Private mRootScopeFault As Boolean, mRootScopeFaultReached As Boolean')
    $rootProbe.AddFromString(@'
Public Sub ArmRootScopeFailure(ByVal enabled As Boolean)
    mRootScopeFault = enabled: mRootScopeFaultReached = False
End Sub
Public Sub BeforeRootScopeCapture()
    If RootAction And mRootScopeFault And Not mRootScopeFaultReached Then
        mRootScopeFaultReached = True
        Err.Raise 70, , "Generated optional discovery capture failure"
    End If
End Sub
Public Function RootScopeFailureObserved() As Boolean
    RootScopeFailureObserved = mRootScopeFaultReached
End Function
'@)
    $console=$admin.VBComponents.Item('modAdminConsole').CodeModule
    $captureLines=@(for($line=1;$line -le $console.CountOfLines;$line++){if($console.Lines($line,1) -match '^Public Function CaptureWarehouseDirectoryInputs\('){$line}})
    if($captureLines.Count -gt 1){throw 'Ambiguous root scope capture boundary'}
    if($captureLines.Count){$console.InsertLines($captureLines[0]+1,'    TestDirectoryColumns.BeforeRootScopeCapture')}
    $ribbon=@($admin.VBComponents|Where-Object {$_.CodeModule.CountOfLines -gt 0 -and $_.CodeModule.Lines(1,$_.CodeModule.CountOfLines) -match 'Public Sub RibbonOnActionAdmin\('})
    if($ribbon.Count -ne 1){throw 'Root Ribbon callback error boundary unavailable'}
    $ribbon[0].CodeModule.AddFromString(@'
Private Function MsgBox(ByVal Prompt As String, Optional ByVal Buttons As VbMsgBoxStyle = vbOKOnly, Optional ByVal Title As String = "") As VbMsgBoxResult
    MsgBox = TestDirectoryColumns.ObserveMessage(Prompt, Buttons)
End Function
'@)
    $core=$packages['invSys.Core.xlam'].VBProject
    $probe=$core.VBComponents.Add(1);$probe.Name='TestRootActivity'
    $probe.CodeModule.AddFromString(@'
Option Explicit
Public Denials As Long
Private mRecordingContext As String
Public Function StartRecording() As Boolean
    Dim report As String
    mRecordingContext = modActivity.CaptureContext()
    StartRecording = modActionRecording.Control(mRecordingContext, "Start", report)
End Function
Public Function RecordingActive() As Boolean
    Dim canStart As Boolean, report As String
    report = modActionRecording.Status(mRecordingContext, RecordingActive, canStart)
End Function
Public Sub CancelRecording()
    Dim report As String
    modActionRecording.Control mRecordingContext, "Cancel", report
End Sub
Public Function DenialCount() As Long
    DenialCount = Denials
End Function
Public Function SetPolicy(ByVal mode As String) As Boolean
    Dim context As String, version As Long, wire As String, notice As String
    Dim model As Object, row As Variant, users As Collection, user As Object
    context = modActivity.CaptureContext()
    If Not modTrackingPolicySettings.ReadEditor(context, version, wire, notice) Then Exit Function
    Set model = modTrackingPolicyModel.Defaults(mode <> "Disabled")
    model("ViewerActionPathCaptureEnabled") = (mode = "Capture")
    For Each row In model("Controls")
        If row("ControlId") = "ADMIN_WAREHOUSE_ROOT_ADD" Then row("Collect") = (mode <> "Disabled" And mode <> "ControlOff")
    Next row
    If mode = "UserOff" Then
        Set user = CreateObject("Scripting.Dictionary")
        user.Add "UserId", "config-admin": user.Add "Record", False
        Set users = model("Users"): users.Add user
    End If
    SetPolicy = modTrackingPolicySettings.SavePolicy(context, version, modTrainingJson.EncodeObject(model), notice)
End Function
Public Function Definition(ByVal version As Long) As String
    Dim record As Object
    Set record = modActivityCatalog.Control("ADMIN_WAREHOUSE_ROOT_ADD", version)
    If Not record Is Nothing Then Definition = modTrainingJson.EncodeObject(record)
End Function
Public Function PriorDefinitionsPreserved() As Boolean
    Dim id As Variant, before As Object, after As Object
    For Each id In modActivityCatalog.ControlIds(39)
        Set before = modActivityCatalog.Control(CStr(id), 39)
        Set after = modActivityCatalog.Control(CStr(id))
        If before Is Nothing Or after Is Nothing Then Exit Function
        If modTrainingJson.EncodeObject(before) <> modTrainingJson.EncodeObject(after) Then Exit Function
    Next id
    PriorDefinitionsPreserved = True
End Function
Public Function Concludes(ByVal wire As String) As Boolean
    Concludes = modEvaluationMatches.CommandCompleted(modTrainingJson.DecodeObject(wire))
End Function
'@)
    $core.VBComponents.Item('modRoleUiAccess').CodeModule.AddFromString(@'
Private Function MsgBox(ByVal Prompt As String, Optional ByVal Buttons As VbMsgBoxStyle = vbOKOnly, Optional ByVal Title As String = "") As VbMsgBoxResult
    TestRootActivity.Denials = TestRootActivity.Denials + 1
    MsgBox = vbOK
End Function
'@)
    $api=$core.VBComponents.Item('modActionRecording').CodeModule
    $gapBody=''
    if($api.Lines(1,$api.CountOfLines) -match 'Public Sub ObservationUnavailable\('){
        $gapBody=@'
    modActionRecording.ObservationUnavailable ""
    modActionRecording.ObservationUnavailable mRecordingContext & "unrelated"
    IgnoreOtherContextGaps = RecordingActive()
'@
    }
    elseif($api.Lines(1,$api.CountOfLines) -match 'Optional ByVal trackingUnavailable As Boolean'){
        $gapBody=@'
    modActionRecording.ContextClosed "", True
    modActionRecording.ContextClosed mRecordingContext & "unrelated", True
    IgnoreOtherContextGaps = RecordingActive()
'@
    }
    $probe.CodeModule.AddFromString("Public Function IgnoreOtherContextGaps() As Boolean`r`n"+$gapBody+"`r`nEnd Function")
    Install-AdminDirectoryBoundaries $probe.CodeModule $rootProbe -ControlId 'ADMIN_WAREHOUSE_ROOT_ADD' -Outcome 'REMEMBERED_REFRESHED' -ProbeName 'TestRootActivity' -RootAction
    Install-AdminRootUnexpected $rootProbe
    $scopeAvailable=@($admin.VBComponents|Where-Object Name -CEQ 'modAdminDirectoryScope').Count -eq 1
    $rootProbe.AddFromString("Public Function RootReadScopeAvailable() As Boolean`r`n    RootReadScopeAvailable = "+[string]$scopeAvailable+"`r`nEnd Function")
    if($scopeAvailable){Install-AdminRootScope $rootProbe $console $script:OriginalRootScanForTest}
    if($AdminWarehouseRootGuidePaths){
        . (Join-Path $PSScriptRoot 'Slice4beAdminVerifyGuide.ps1')
        Install-AdminVerifyGuideProbe
    }
}

function Test-AdminRootRecording($Fixture,$Other,$HostBook,$OtherBook) {
    function Probe([string]$Name,[object[]]$Values=@()){Run 'invSys.Core.xlam' ('TestRootActivity.'+$Name) $Values}
    function Retained($Before){foreach($path in $Before.Keys){if(-not(Test-Path -LiteralPath $path)){return $false};$s=[IO.File]::Open($path,'Open','Read','ReadWrite');try{$hash=(Get-FileHash -InputStream $s).Hash}finally{$s.Dispose()};if($hash -cne $Before[$path]){return $false}};return $true}
    SelectTarget $Fixture
    $wire=[string](Probe 'Definition' @(40));$definition=if($wire){$wire|ConvertFrom-Json}else{$null}
    Check 'WarehouseRoot.Recording.RegisteredDefinition' ($null -ne $definition -and $definition.OwnerId -ceq 'ADMIN_WAREHOUSE_ROOT_PREFERENCE' -and $definition.Class -ceq 'Command' -and $definition.Caption -ceq 'Add Warehouse Root' -and $definition.Capability -ceq 'ADMIN_MAINT')
    Check 'WarehouseRoot.Recording.AbsentFrom39' ([string](Probe 'Definition' @(39)) -ceq '')
    Check 'WarehouseRoot.Recording.PriorDefinitionsPreserved' ([bool](Probe 'PriorDefinitionsPreserved'))
    foreach($case in @('Opened','Cancel','WriteFailure','ReadFailure','ReadMismatch','DirectoryFailure','WriteAndDirectoryFailure','Disabled','ControlOff','UserOff','Reader','SignedOut','ScopeUnavailable')){
        SelectTarget $Fixture
        $mode=if($case -cin @('Disabled','ControlOff','UserOff')){$case}else{'Enabled'}
        if(-not [bool](Probe 'SetPolicy' @($mode))){throw 'Root recording policy fixture unavailable; not behavioral RED'}
        $before=@(Get-Slice4beActivityFiles $Fixture);$authority=Pins $Fixture.Root;$otherAuthority=Pins $Other.Root
        $denied=$case -cin @('Reader','SignedOut');$denials=[int](Probe 'DenialCount')
        if($case -ceq 'Reader'){SelectTarget $Fixture 'config-reader'}
        if($case -ceq 'SignedOut'){[void](Run 'invSys.Core.xlam' 'modAuth.SignOut')}
        $cancel=$case -ceq 'Cancel';$writeFault=$case -cin @('WriteFailure','WriteAndDirectoryFailure')
        $directoryFault=$case -like '*DirectoryFailure'
        $readMode=if($case -cin @('ReadFailure','ReadMismatch')){$case}else{''}
        $entered=if($cancel){''}else{Join-Path $runRoot ('recorded-root-'+$case)}
        $scopeFault=$case -ceq 'ScopeUnavailable'
        [void](Invoke 'ArmRootScopeFailure' @($scopeFault))
        try{$facts=([string](Invoke 'InvokeRoot' @($entered,$writeFault,$readMode,$directoryFault))) -split '\|'}finally{
            $scopeFaultReached=[bool](Invoke 'RootScopeFailureObserved')
            [void](Invoke 'ArmRootScopeFailure' @($false))
        }
        if($facts.Count -ne 14){throw 'Root callback facts unavailable; not behavioral RED'}
        $prefix='WarehouseRoot.Recording.'+$case+'.';$called=-not ($denied -or $cancel)
        $confirmed=$called -and -not ($writeFault -or $readMode);$opened=$called -and -not $directoryFault
        Check ($prefix+'OriginalOwnerCounts') ($facts[0] -ceq $(if($called){'1'}else{'0'}) -and $facts[4] -ceq $(if($called){'1'}else{'0'}) -and $facts[3] -ceq $(if($denied){'0'}else{'1'}))
        Check ($prefix+'OriginalPresentationAndPermission') ($facts[1] -ceq [string]$opened -and $facts[8] -ceq 'True' -and ([int](Probe 'DenialCount')-$denials) -eq $(if($denied){1}else{0}))
        Check ($prefix+'TruthfulFeedbackCount') ($facts[11] -ceq $(if($called){'1'}else{'0'}) -and $facts[2] -ceq $(if($confirmed -and $opened){'1'}else{'0'}) -and [int](Invoke 'RootTrackingNotices') -eq $(if($scopeFault){1}else{0}))
        if($scopeFault){Check ($prefix+'ActualOptionalCaptureFailure') $scopeFaultReached}
        $preserved=$true;foreach($part in @('Managed','Headers','Text','Formula','Tail','Notes')){$preserved=$preserved -and [bool](Invoke 'Preserved' @($part))}
        Check ($prefix+'WorkbookAndAuthorityPreserved') ($preserved -and (Retained $authority) -and (Retained $otherAuthority) -and [string]$OtherBook.Worksheets.Item(1).Cells.Item(1,1).Value2 -ceq 'Other workbook canary')
        $records=@(Get-Slice4beActivityFiles $Fixture|Where-Object {$_ -cnotin $before}|ForEach-Object {Get-Content $_ -Raw|ConvertFrom-Json}|Where-Object ControlId -CEQ 'ADMIN_WAREHOUSE_ROOT_ADD')
        if($denied -or $mode -cne 'Enabled' -or $scopeFault){Check ($prefix+'NoFabricatedAction') ($records.Count -eq 0);continue}
        $expected=if($cancel){'CANCELLED'}elseif($confirmed){if($opened){'REMEMBERED_REFRESHED'}else{'REMEMBERED_REFRESH_FAILED'}}else{if($opened){'UNCONFIRMED_REFRESHED'}else{'UNCONFIRMED_REFRESH_FAILED'}}
        $requests=@($records|Where-Object OutcomeCode -CEQ 'REQUESTED');$outcomes=@($records|Where-Object OutcomeCode -CEQ $expected)
        $pair=$records.Count -eq 2 -and $requests.Count -eq 1 -and $outcomes.Count -eq 1
        Check ($prefix+'ExactCorrelatedPair') ($pair -and $requests[0].ActivityId -ceq $outcomes[0].ActivityId -and $requests[0].RecordId -cne $outcomes[0].RecordId)
        $owned=$pair;$safe=$pair;$conclusion=$pair
        foreach($record in $records){
            $owned=$owned -and $record.OwnerId -ceq 'ADMIN_WAREHOUSE_ROOT_PREFERENCE' -and $record.WarehouseId -ceq $Fixture.Warehouse -and $record.UserId -ceq 'config-admin' -and $record.StationId -ceq 'S1' -and $record.CatalogVersion -eq 40
            $effect=if($record.OutcomeCode -ceq 'CANCELLED'){'Unchanged'}else{'Unknown'}
            $severity=if($record.OutcomeCode -like '*REFRESH_FAILED'){'Error'}elseif($record.OutcomeCode -ceq 'UNCONFIRMED_REFRESHED'){'Warning'}else{'Info'}
            $owned=$owned -and $record.DataEffect -ceq $effect -and $record.Severity -ceq $severity
            $json=$record|ConvertTo-Json -Depth 20 -Compress
            $safe=$safe -and @($record.SourceEventRefs).Count -eq 0
            foreach($value in @($Fixture.Root,$Other.Root,$Fixture.Secret,$entered,'Keep this local note')){if($value){$safe=$safe -and $json.IndexOf($value,[StringComparison]::OrdinalIgnoreCase) -lt 0}}
            $conclusion=$conclusion -and [bool](Probe 'Concludes' @($json)) -eq ($record.OutcomeCode -ceq 'REMEMBERED_REFRESHED')
        }
        Check ($prefix+'CapturedContextAndHonestEffects') $owned
        Check ($prefix+'NoPayloadOrSourceReferences') $safe
        Check ($prefix+'OnlyConfirmedWholeCommandConcludes') $conclusion
    }
    Test-AdminRootUnexpected $Fixture
    Test-AdminRootScope $Fixture $Other
    Test-AdminRootBoundaries $Fixture $Other $OtherBook
    SelectTarget $Fixture
    if(-not [bool](Probe 'SetPolicy' @('Capture'))){throw 'Active root recording policy prerequisite unavailable.'}
    $journalRoot=Join-Path $Fixture.Root ('Training/ActionPaths/'+$Fixture.Warehouse)
    $prior=Pins $Fixture.Root;$before=@(Get-Slice4beActivityFiles $Fixture)
    if(-not [bool](Probe 'StartRecording')){throw 'Active root recording prerequisite unavailable; not behavioral RED.'}
    $starts=@(Get-ChildItem $journalRoot -File -Filter '*.json'|Where-Object {-not $prior.ContainsKey($_.FullName)}|ForEach-Object {Get-Content $_.FullName -Raw|ConvertFrom-Json}|Where-Object RecordType -CEQ 'Start')
    if($starts.Count -ne 1 -or -not [bool](Probe 'RecordingActive')){throw 'One active durable recording required.'}
    $sequence=$starts[0].SequenceId;$authority=Pins $Fixture.Root
    Check 'WarehouseRoot.Recording.ActiveScopeUnavailable.OtherContextCannotClose' ([bool](Probe 'IgnoreOtherContextGaps') -and (Same $authority (Pins $Fixture.Root)))
    [void](Invoke 'ArmRootScopeFailure' @($true))
    try{
        $facts=([string](Invoke 'InvokeRoot' @((Join-Path $runRoot 'active-scope-fault'),$false,'',$false))) -split '\|'
        Check 'WarehouseRoot.Recording.ActiveScopeUnavailable.FaultReached' ([bool](Invoke 'RootScopeFailureObserved'))
        Check 'WarehouseRoot.Recording.ActiveScopeUnavailable.OwnerAndFeedbackPreserved' ($facts.Count -eq 14 -and $facts[0] -ceq '1' -and $facts[1] -ceq 'True' -and $facts[2] -ceq '1' -and $facts[3] -ceq '1' -and $facts[4] -ceq '1' -and $facts[8] -ceq 'True' -and $facts[11] -ceq '1' -and [int](Invoke 'RootTrackingNotices') -eq 1)
        Check 'WarehouseRoot.Recording.ActiveScopeUnavailable.NoLongerActive' (-not [bool](Probe 'RecordingActive'))
        $closed=@(Get-ChildItem $journalRoot -File -Filter '*.json'|ForEach-Object {Get-Content $_.FullName -Raw|ConvertFrom-Json}|Where-Object {$_.SequenceId -ceq $sequence -and $_.RecordType -ceq 'Close'})
        $valid=$closed.Count -eq 1
        if($valid){$valid=$closed[0].Lifecycle -ceq 'Incomplete' -and $closed[0].ReasonCode -ceq 'TRACKING_UNAVAILABLE' -and $closed[0].ActionCount -eq 0 -and @($closed[0].Observations).Count -eq 0}
        Check 'WarehouseRoot.Recording.ActiveScopeUnavailable.DurableIncomplete' $valid
        Check 'WarehouseRoot.Recording.ActiveScopeUnavailable.NoFabricatedAction' (@(Get-Slice4beActivityFiles $Fixture|Where-Object {$_ -cnotin $before}).Count -eq 0)
        Check 'WarehouseRoot.Recording.ActiveScopeUnavailable.PriorAuthorityRetained' (Retained $authority)
    }finally{
        [void](Invoke 'ArmRootScopeFailure' @($false))
        if([bool](Probe 'RecordingActive')){[void](Probe 'CancelRecording')}
    }
    if($AdminWarehouseRootGuidePaths){Test-AdminRootGuide $Fixture $HostBook}
}
