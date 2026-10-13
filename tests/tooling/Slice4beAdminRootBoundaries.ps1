# Same public root callback; faults at optional entry and the actual store write.
function Install-AdminRootUnexpected($Probe) {
    $Probe.InsertLines($Probe.CountOfDeclarationLines+1,'Private mUnexpectedRootPrompt As Boolean, mUnexpectedRootReached As Boolean')
    $Probe.AddFromString(@'
Public Sub ArmUnexpectedRootPrompt(ByVal enabled As Boolean)
    mUnexpectedRootPrompt = enabled: mUnexpectedRootReached = False
End Sub
Public Sub RootPromptBoundary()
    If RootAction And mUnexpectedRootPrompt Then
        mUnexpectedRootReached = True
        Err.Raise 70, , "Generated unexpected root prompt failure"
    End If
End Sub
Public Function UnexpectedRootPromptReached() As Boolean
    UnexpectedRootPromptReached = mUnexpectedRootReached
End Function
'@)
    $start=$Probe.ProcStartLine('ObservePrompt',0);$count=$Probe.ProcCountLines('ObservePrompt',0)
    $source=$Probe.Lines($start,$count)
    $changed=$source.Replace('    mPrompts = mPrompts + 1',"    mPrompts = mPrompts + 1`r`n    RootPromptBoundary")
    if($source -ceq $changed){throw 'Unexpected root prompt boundary unavailable.'}
    $Probe.DeleteLines($start,$count);$Probe.InsertLines($start,$changed)
}

function Test-AdminRootUnexpected($Fixture) {
    SelectTarget $Fixture
    if(-not [bool](Probe 'SetPolicy' @('Enabled'))){throw 'Unexpected root policy prerequisite failed.'}
    $before=@(Get-Slice4beActivityFiles $Fixture);$authority=Pins $Fixture.Root
    [void](Invoke 'ArmUnexpectedRootPrompt' @($true))
    try{
        $facts=([string](Invoke 'InvokeRoot' @((Join-Path $runRoot 'unexpected-root'),$false,'',$false))) -split '\|'
        Check 'WarehouseRoot.Unexpected.ActualPromptFailure' ([bool](Invoke 'UnexpectedRootPromptReached'))
        Check 'WarehouseRoot.Unexpected.NoWriteOrDirectoryAndOriginalFeedback' ($facts.Count -eq 14 -and $facts[0] -ceq '0' -and $facts[1] -ceq 'False' -and $facts[2] -ceq '0' -and $facts[3] -ceq '1' -and $facts[4] -ceq '0' -and $facts[8] -ceq 'True' -and $facts[11] -ceq '1' -and $facts[13].Contains('Generated unexpected root prompt failure') -and [int](Invoke 'RootTrackingNotices') -eq 0)
        $records=@(Get-Slice4beActivityFiles $Fixture|Where-Object {$_ -cnotin $before}|ForEach-Object {Get-Content $_ -Raw|ConvertFrom-Json}|Where-Object ControlId -CEQ 'ADMIN_WAREHOUSE_ROOT_ADD')
        $request=@($records|Where-Object OutcomeCode -CEQ 'REQUESTED');$failed=@($records|Where-Object OutcomeCode -CEQ 'FAILED')
        $pair=$records.Count -eq 2 -and $request.Count -eq 1 -and $failed.Count -eq 1
        Check 'WarehouseRoot.Unexpected.CorrelatedFailedOutcome' ($pair -and $request[0].ActivityId -ceq $failed[0].ActivityId -and $failed[0].Severity -ceq 'Error' -and $failed[0].DataEffect -ceq 'Unknown')
        $safe=$pair
        foreach($record in $records){$wire=$record|ConvertTo-Json -Depth 20 -Compress;$safe=$safe -and @($record.SourceEventRefs).Count -eq 0 -and -not [bool](Probe 'Concludes' @($wire)) -and -not $wire.Contains('Generated unexpected root prompt failure')}
        Check 'WarehouseRoot.Unexpected.NoFalseConclusionOrDiagnosticPayload' $safe
        $preserved=$true;foreach($part in @('Managed','Headers','Text','Formula','Tail','Notes')){$preserved=$preserved -and [bool](Invoke 'Preserved' @($part))}
        Check 'WarehouseRoot.Unexpected.WorkbookAndSavedAuthorityPreserved' ($preserved -and (Retained $authority))
    }finally{[void](Invoke 'ArmUnexpectedRootPrompt' @($false))}
}

function Test-AdminRootBoundaries($Fixture,$Other,$OtherBook) {
    foreach($case in @('SwitchWorkbook','SignOut','SwitchTarget','AttemptWriteFailure','OutcomeWriteFailure')){
        SelectTarget $Fixture
        if(-not [bool](Probe 'SetPolicy' @('Enabled'))){throw 'Root boundary policy prerequisite failed.'}
        $before=@(Get-Slice4beActivityFiles $Fixture);$otherBefore=@(Get-Slice4beActivityFiles $Other)
        $authority=Pins $Fixture.Root;$otherAuthority=Pins $Other.Root
        [void](Probe 'ArmBoundary' @($case,$Other.Root,$Other.Secret,$OtherBook.Name))
        try{
            $facts=([string](Invoke 'InvokeRoot' @((Join-Path $runRoot ('root-boundary-'+$case)),$false,'',$false))) -split '\|'
            if($facts.Count -ne 14){throw 'Root boundary callback facts unavailable.'}
            $reached=if($case -like '*WriteFailure'){[int](Probe 'WriteFailures') -eq 1}else{[bool](Probe 'EntryObserved')}
            if(-not $reached){throw 'Root actual optional boundary was not reached; not behavioral RED.'}
            $prefix='WarehouseRoot.Boundary.'+$case+'.'
            Check ($prefix+'ActualFaultReached') $reached
            Check ($prefix+'OwnerOnceWithOriginalPresentation') ($facts[0] -ceq '1' -and $facts[1] -ceq 'True' -and $facts[2] -ceq '1' -and $facts[3] -ceq '1' -and $facts[4] -ceq '1' -and $facts[8] -ceq 'True' -and $facts[11] -ceq '1')
            Check ($prefix+'OneTruthfulNoticeOrNone') ([int](Invoke 'RootTrackingNotices') -eq $(if($case -ceq 'SwitchWorkbook'){0}else{1}))
            Check ($prefix+'OriginalReadScopeRetained') ([bool](Invoke 'Preserved' @('Managed')))
            $preserved=$true;foreach($part in @('Headers','Text','Formula','Tail','Notes')){$preserved=$preserved -and [bool](Invoke 'Preserved' @($part))}
            Check ($prefix+'LocalAndOtherWorkbookPreserved') ($preserved -and [string]$OtherBook.Worksheets.Item(1).Cells.Item(1,1).Value2 -ceq 'Other workbook canary')
            Check ($prefix+'SavedAuthorityAndHistoryPreserved') ((Retained $authority) -and (Retained $otherAuthority))
            $records=@(Get-Slice4beActivityFiles $Fixture|Where-Object {$_ -cnotin $before}|ForEach-Object {Get-Content $_ -Raw|ConvertFrom-Json}|Where-Object ControlId -CEQ 'ADMIN_WAREHOUSE_ROOT_ADD')
            $honest=if($case -ceq 'SwitchWorkbook'){$records.Count -eq 2 -and @($records|Where-Object OutcomeCode -CEQ 'REMEMBERED_REFRESHED').Count -eq 1}elseif($case -ceq 'OutcomeWriteFailure'){$records.Count -eq 1 -and $records[0].OutcomeCode -ceq 'REQUESTED'}else{$records.Count -eq 0}
            Check ($prefix+'HonestOriginalEvidence') $honest
            Check ($prefix+'NoEvidenceInOtherWarehouse') ((@(Get-Slice4beActivityFiles $Other) -join '|') -ceq ($otherBefore -join '|'))
        }finally{[void](Probe 'ArmBoundary' @('','','',''));SelectTarget $Fixture}
    }
}
