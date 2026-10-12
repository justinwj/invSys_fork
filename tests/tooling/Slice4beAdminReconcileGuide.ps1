# Actual visible Reconcile journeys; shared authoring/evaluation handlers consume their records.
. (Join-Path $PSScriptRoot 'Slice4beAdminReconcileAppliedGuide.ps1')
function Install-ReconcileGuideProbe {
    . (Join-Path $PSScriptRoot 'Slice4beViewerPublishedRead.ps1')
    Test-Slice4beViewerPublishedRead $null $null $true
    . (Join-Path $PSScriptRoot 'Slice4beActionRecording.ps1')
    Test-Slice4beActionRecording $null $true
    . (Join-Path $PSScriptRoot 'Slice4beRecordingReader.ps1')
    Test-Slice4beRecordingReader $null $null $true
    . (Join-Path $PSScriptRoot 'Slice4beRecordingEvaluation.ps1')
    Install-RecordingEvaluationProbe
    . (Join-Path $PSScriptRoot 'Slice4beGuideDraft.ps1')
    Install-GuideDraftProbe
    if($AdminReconcileAppliedGuide){Install-ReconcileAppliedGuideProbe}
}

function Test-ReconcileGuide($Fixture,$Events,$State) {
    . (Join-Path $PSScriptRoot 'Slice4beRecordingFixture.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beRecordingEvaluation.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beGuideTestActions.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beGeneralSettingsGuidePaths.ps1')
    $steps=@(@('Shipment Reconcile','ADMIN_RECONCILE_OPEN','OPENED'),@('Inspect and confirm','ADMIN_RECONCILE_INSPECT','CONFIRMED'))
    $runs=@();$hostBook=$null;$originalReportRoot=$reportRoot
    $inventoryHash=Hash $State.InventoryPath
    $first=[string](Run 'invSys.Admin.xlam' 'TestReconcileIdentity.NativeFirstChoiceEvent')
    try{
        if(-not [bool](Run 'invSys.Core.xlam' 'TestReconcileRecording.EnablePolicy' @('Capture'))){throw 'Recording policy prerequisite unavailable.'}
        $hostBook=$excel.Workbooks.Add();$hostBook.Windows.Item(1).Visible=$true;$excel.Visible=$true
        $hostBook.Worksheets.Item(1).Cells.Item(1,1).Value2='Reconcile guide canary'
        foreach($kind in @('Source','Observed','Interrupted')){
            CloseRecordingViewer;SelectTarget $Fixture 'config-admin';$hostBook.Activate();OpenRecordingViewer
            $prior=BoundPins $journalRoot;$oldActivity=ActivityPins
            Check ('Reconcile.Recorded.'+$kind+'.ActualStart') ((RecordingControl 'Start Recording' 'Click') -ceq 'DELIVERED')
            $starts=@(Get-ChildItem $journalRoot -File -Filter '*.json'|Where-Object {-not $prior.ContainsKey($_.FullName)}|ForEach-Object {Get-Content $_.FullName -Raw|ConvertFrom-Json}|Where-Object RecordType -CEQ 'Start')
            if($starts.Count -ne 1){throw 'One durable recording Start required.'}
            $sequence=$starts[0].SequenceId;$before=@(Get-Slice4beActivityFiles $Fixture);$rows=$Events.ListRows.Count
            $reportRoot=Join-Path $originalReportRoot ('guide-'+$kind.ToLowerInvariant());New-Item -ItemType Directory -Path $reportRoot|Out-Null
            if($kind -ceq 'Interrupted'){[void](Run 'invSys.Core.xlam' 'TestReconcileRecording.ArmStoreFailure' @('CONFIRMED'))}
            try{$native=Invoke-ReconcileSelectionNative -Confirmation Submit -ExpectedEvent $first -ExpectedKey $State.KeyA}
            finally{[void](Run 'invSys.Core.xlam' 'TestReconcileRecording.ArmStoreFailure' @(''));$reportRoot=$originalReportRoot}
            Check ('Reconcile.Recorded.'+$kind+'.NativeSubmissionOnce') ($native.Returned -and $native.Confirmation.ActionDispatched -and $native.Confirmation.NarrativeExact -and $native.Confirmation.ExactPreview -and $native.Confirmation.TextFits -and $Events.ListRows.Count -eq $rows+1)
            $records=@(Get-Slice4beActivityFiles $Fixture|Where-Object {$_ -cnotin $before}|ForEach-Object {Get-Content $_ -Raw|ConvertFrom-Json})
            $success=$kind -cne 'Interrupted';$expectedCount=if($success){4}else{3}
            $valid=$records.Count -eq $expectedCount -and @($records|Where-Object {$_.Ordinal -eq 1 -and $_.ControlId -ceq $steps[0][1] -and $_.OutcomeCode -ceq 'OPENED'}).Count -eq 1 -and @($records|Where-Object {$_.Ordinal -eq 2 -and $_.ControlId -ceq $steps[1][1] -and $_.OutcomeCode -ceq 'REQUESTED'}).Count -eq 1
            $catalogVersion=[int](Run 'invSys.Core.xlam' 'TestReconcileRecording.CurrentCatalogVersion')
            foreach($record in $records){$valid=$valid -and $record.SequenceId -ceq $sequence -and $record.WarehouseId -ceq $Fixture.Warehouse -and $record.UserId -ceq 'config-admin' -and $record.CatalogVersion -eq $catalogVersion}
            Check ('Reconcile.Recorded.'+$kind+'.ExactActionsAndContext') $valid
            $confirmed=@($records|Where-Object OutcomeCode -CEQ 'CONFIRMED');$valid=$confirmed.Count -eq $(if($success){1}else{0})
            if($success -and $confirmed.Count -eq 1){
                $newId=[string]$Events.ListRows.Item($Events.ListRows.Count).Range.Cells.Item(1,$Events.ListColumns.Item('EventID').Index).Value2
                $refs=@($confirmed[0].SourceEventRefs);$valid=$valid -and $refs.Count -eq 1
                if($refs.Count -eq 1){$valid=$valid -and $refs[0].EventId -ceq $newId -and $refs[0].SubmissionState -ceq 'Submitted' -and $newId -cne $first}
            }
            Check ('Reconcile.Recorded.'+$kind+'.OnlyFreshCorrectionReference') $valid
            if($success){Check ('Reconcile.Recorded.'+$kind+'.ActualStop') ((RecordingControl 'Stop Recording' 'Click') -ceq 'DELIVERED')}
            else{Check 'Reconcile.Recorded.Interrupted.ExactWarning' ([bool](Run 'invSys.Admin.xlam' 'TestReconcileIdentity.ExactStoreWarning'))}
            $closed=@(RecordingJournal $sequence|Where-Object RecordType -CEQ 'Close');$valid=$closed.Count -eq 1
            if($valid){$valid=$closed[0].ActionCount -eq 2 -and @($closed[0].Observations).Count -eq $expectedCount -and $closed[0].Lifecycle -ceq $(if($success){'Stopped'}else{'Incomplete'})}
            if($valid -and -not $success){$valid=$closed[0].ReasonCode -ceq 'TRACKING_UNAVAILABLE'}
            Check ('Reconcile.Recorded.'+$kind+'.ExactJournalLifecycle') $valid
            Check ('Reconcile.Recorded.'+$kind+'.ImmutableJournalChain') (JournalChain $sequence ($expectedCount+2))
            Check ('Reconcile.Recorded.'+$kind+'.PriorActivityImmutable') (PinsRetained $oldActivity)
            Check ('Reconcile.Recorded.'+$kind+'.InventoryPreserved') ((Hash $State.InventoryPath) -ceq $inventoryHash -and [bool](Run 'invSys.Admin.xlam' 'TestReconcileIdentity.NativeFixturePreserved' @(4)))
            if($closed.Count -ne 1){throw 'Closed recording prerequisite unavailable.'}
            $runs+=@{Kind=$kind;Journal=$closed[0];Records=$records;Success=$success;Incomplete=(-not $success)}
        }
        CloseRecordingViewer;SelectTarget $Fixture 'config-admin'
        if(-not [bool](Run 'invSys.Admin.xlam' 'modAdminConsole.PublishReadFixtureForTest')){throw 'Actual publication prerequisite failed.'}
        $publication=Get-Content (Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Snapshot.Events.json')) -Raw|ConvertFrom-Json
        $lines=@($publication.Groups|Where-Object Source -CEQ 'Activity'|ForEach-Object {$_.Lines})
        $exact=$true
        foreach($run in $runs){foreach($record in $run.Records){$matches=@($lines|Where-Object RecordId -CEQ $record.RecordId);$exact=$exact -and $matches.Count -eq 1;if($matches.Count -eq 1){$exact=$exact -and $matches[0].ActivityId -ceq $record.ActivityId -and $matches[0].OutcomeCode -ceq $record.OutcomeCode}}}
        Check 'Reconcile.Recorded.ExactPublishedRecords' $exact
        Test-GeneralSettingsGuidePaths $Fixture $runs $steps -Reconcile
        Check 'Reconcile.Recorded.FinalInventoryAndCanaryPreserved' ((Hash $State.InventoryPath) -ceq $inventoryHash -and [bool](Run 'invSys.Admin.xlam' 'TestReconcileIdentity.NativeFixturePreserved' @(4)) -and [string]$hostBook.Worksheets.Item(1).Cells.Item(1,1).Value2 -ceq 'Reconcile guide canary')
        if($AdminReconcileAppliedGuide){Test-ReconcileAppliedGuide $Fixture $runs $Events $State $hostBook}
    }finally{
        $reportRoot=$originalReportRoot;CloseRecordingViewer
        if($null -ne $hostBook){$hostBook.Close($false)}
        SelectTarget $Fixture 'config-admin'
    }
}
