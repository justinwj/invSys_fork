# Real callback records feed actual recording, guide and comparison controls.
function Install-AdminVerifyGuideProbe {
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
}

function Test-AdminVerifyGuide($Fixture,$HostBook,$OtherBook) {
    . (Join-Path $PSScriptRoot 'Slice4beRecordingFixture.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beRecordingEvaluation.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beGuideTestActions.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beGeneralSettingsGuidePaths.ps1')
    function Hash([string]$Path){FileHash $Path}
    $steps=@(@('Verify Add-ins Published','ADMIN_VERIFY_ADDINS','VERIFIED'),@('Verify Add-ins Published','ADMIN_VERIFY_ADDINS','VERIFIED'))
    $root=Join-Path $runRoot 'verify-guide';$addins=Join-Path $root 'Addins'
    New-Item -ItemType Directory -Path $addins|Out-Null
    foreach($name in @('Core','Inventory.Domain','Designs.Domain','Operations','Admin')){[IO.File]::WriteAllText((Join-Path $addins ('invSys.'+$name+'.xlam')),'Non-executable file-presence fixture')}
    $files=Pins $root;$runs=@();$inventory=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Data.Inventory.xlsb');$inventoryHash=Hash $inventory
    try{
        SelectTarget $Fixture
        $authorized=[bool](Run 'invSys.Core.xlam' 'modAuth.CanPerform' @('ACTION_PATH_MAINT','config-admin',$Fixture.Warehouse,'S1'))
        if(-not $authorized){throw 'Explicit guide-maintenance fixture grant unavailable; not product RED.'}
        Check 'AdminVerify.Recorded.ExplicitGuidePermission' $authorized
        $HostBook.Windows.Item(1).Visible=$true;$excel.Visible=$true
        Check 'AdminVerify.Recorded.VisibleHost' ([bool]$excel.Visible -and [bool]$HostBook.Windows.Item(1).Visible)
        if(-not [bool](Probe 'SetPolicy' @('Capture'))){throw 'Capture policy prerequisite failed.'}
        if(-not [bool](Run 'invSys.Admin.xlam' 'modAdminConsole.PublishReadFixtureForTest')){throw 'Initial Viewer publication unavailable.'}
        $kinds=if($AdminVerifyExpectationChoicesOnly){@('Source')}else{@('Source','Observed','Interrupted')}
        foreach($kind in $kinds){
            CloseRecordingViewer;SelectTarget $Fixture;OpenRecordingViewer
            $prior=BoundPins $journalRoot;$oldActivity=ActivityPins;$before=@(Get-Slice4beActivityFiles $Fixture)
            Check ('AdminVerify.Recorded.'+$kind+'.ActualStart') ((RecordingControl 'Start Recording' 'Click') -ceq 'DELIVERED')
            $starts=@(Get-ChildItem $journalRoot -File -Filter '*.json'|Where-Object {-not $prior.ContainsKey($_.FullName)}|ForEach-Object {Get-Content $_.FullName -Raw|ConvertFrom-Json}|Where-Object RecordType -CEQ 'Start')
            if($starts.Count -ne 1){throw 'One durable Start required.'}
            $sequence=$starts[0].SequenceId;$success=$kind -cne 'Interrupted'
            for($ordinal=1;$ordinal -le 2;$ordinal++){
                $fault=-not $success -and $ordinal -eq 2
                $mode=if($fault){'OutcomeWriteFailure'}else{''}
                [void](Probe 'ArmBoundary' @($mode,'','',''))
                try{
                    $result=([string](Run 'invSys.Admin.xlam' 'TestAdminVerify.Invoke' @($root,$HostBook.Name,$OtherBook.Name))) -split '\|'
                    $faults=[int](Probe 'WriteFailures')
                }finally{[void](Probe 'ArmBoundary' @('','','',''))}
                if($fault -and $faults -ne 1){throw 'Real outcome write failure prerequisite not reached.'}
                $notices=if($fault){'1'}else{'0'}
                Check ('AdminVerify.Recorded.'+$kind+'.OwnerResultOnce.'+$ordinal) ($result.Count -eq 7 -and $result[0] -ceq '1' -and $result[1] -ceq 'VERIFIED' -and $result[2] -ceq $notices -and $result[6] -ceq 'True')
            }
            $records=@(Get-Slice4beActivityFiles $Fixture|Where-Object {$_ -cnotin $before}|ForEach-Object {Get-Content $_ -Raw|ConvertFrom-Json}|Where-Object ControlId -CEQ 'ADMIN_VERIFY_ADDINS')
            $count=if($success){4}else{3};$valid=$records.Count -eq $count
            foreach($record in $records){$valid=$valid -and $record.SequenceId -ceq $sequence -and $record.CatalogVersion -eq 35 -and $record.WarehouseId -ceq $Fixture.Warehouse -and $record.UserId -ceq 'config-admin' -and $record.DataEffect -ceq 'Unknown' -and @($record.SourceEventRefs).Count -eq 0}
            foreach($ordinal in 1,2){$expected=if(-not $success -and $ordinal -eq 2){1}else{2};$part=@($records|Where-Object Ordinal -EQ $ordinal);$valid=$valid -and $part.Count -eq $expected -and @($part|Where-Object OutcomeCode -CEQ 'REQUESTED').Count -eq 1 -and @($part|Where-Object OutcomeCode -CEQ 'VERIFIED').Count -eq ($expected-1)}
            Check ('AdminVerify.Recorded.'+$kind+'.ExactActionsAndContext') $valid
            if($success){Check ('AdminVerify.Recorded.'+$kind+'.ActualStop') ((RecordingControl 'Stop Recording' 'Click') -ceq 'DELIVERED')}
            $closed=@(RecordingJournal $sequence|Where-Object RecordType -CEQ 'Close');$valid=$closed.Count -eq 1
            if($valid){$valid=$closed[0].ActionCount -eq 2 -and @($closed[0].Observations).Count -eq $count -and $closed[0].Lifecycle -ceq $(if($success){'Stopped'}else{'Incomplete'})}
            if($valid -and -not $success){$valid=$closed[0].ReasonCode -ceq 'TRACKING_UNAVAILABLE'}
            Check ('AdminVerify.Recorded.'+$kind+'.ExactJournalLifecycle') $valid
            Check ('AdminVerify.Recorded.'+$kind+'.ImmutableJournalChain') (JournalChain $sequence ($count+2))
            Check ('AdminVerify.Recorded.'+$kind+'.PriorActivityImmutable') (PinsRetained $oldActivity)
            if($closed.Count -ne 1){throw 'Closed recording unavailable.'}
            $runs+=@{Kind=$kind;Journal=$closed[0];Records=$records;Success=$success;Incomplete=(-not $success)}
        }
        CloseRecordingViewer;SelectTarget $Fixture
        if($AdminVerifyExpectationChoicesOnly){
            Test-AdminVerifyExpectationChoices $Fixture $runs[0].Journal.ActionPathId
            Check 'AdminVerify.Expectation.FilesInventoryAndCustomValuePreserved' ((SamePins $files (Pins $root)) -and (Hash $inventory) -ceq $inventoryHash -and [bool](Run 'invSys.Admin.xlam' 'TestAdminVerify.HostPreserved' @($HostBook.Name)))
            return
        }
        if(-not [bool](Run 'invSys.Admin.xlam' 'modAdminConsole.PublishReadFixtureForTest')){throw 'Recorded publication failed.'}
        $publication=Get-Content (Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Snapshot.Events.json')) -Raw|ConvertFrom-Json
        $lines=@($publication.Groups|Where-Object Source -CEQ 'Activity'|ForEach-Object {$_.Lines});$exact=$true
        foreach($run in $runs){foreach($record in $run.Records){$matches=@($lines|Where-Object RecordId -CEQ $record.RecordId);$exact=$exact -and $matches.Count -eq 1;if($matches.Count -eq 1){$exact=$exact -and $matches[0].ActivityId -ceq $record.ActivityId -and $matches[0].OutcomeCode -ceq $record.OutcomeCode}}}
        Check 'AdminVerify.Recorded.ExactPublishedRecords' $exact
        Test-GeneralSettingsGuidePaths $Fixture $runs $steps -VerifyAddins
        Check 'AdminVerify.Recorded.FilesInventoryAndCustomValuePreserved' ((SamePins $files (Pins $root)) -and (Hash $inventory) -ceq $inventoryHash -and [bool](Run 'invSys.Admin.xlam' 'TestAdminVerify.HostPreserved' @($HostBook.Name)))
    }finally{[void](Probe 'ArmBoundary' @('','','',''));CloseRecordingViewer;SelectTarget $Fixture}
}

function Test-AdminVerifyExpectationChoices($Fixture,[string]$PathId) {
    $pins=BoundPins $journalRoot
    OpenRecordingViewer
    if((BoundLibrary 'Open') -cne 'DELIVERED' -or (BoundLibrary 'Select' $PathId) -cne 'SELECTED'){throw 'Saved recording selection prerequisite unavailable.'}
    if((BoundControl 'btnExpectedConclusion' 'Click' '' 'frmActionPaths') -cne 'DELIVERED'){throw 'Existing Expected conclusion handler unavailable.'}
    try{
        $selected=BoundControl 'cboExpectedControl' 'Write' 'ADMIN_VERIFY_ADDINS' 'frmActionPathExpectation'
        Check 'AdminVerify.Expectation.ActualControlSelection' ($selected -ceq 'DELIVERED')
        if($selected -cne 'DELIVERED'){throw 'Registered control selection unavailable.'}
        $choices=@((BoundControl 'cboExpectedOutcome' 'Values' '' 'frmActionPathExpectation') -split "`n"|Where-Object {$_ -cne ''})
        Check 'AdminVerify.Expectation.ExistingOutcomesPresent' ('REQUESTED' -cin $choices -and 'FAILED' -cin $choices)
        Check 'AdminVerify.Expectation.VerifiedOutcomeOffered' ('VERIFIED' -cin $choices)
        Check 'AdminVerify.Expectation.NoUnregisteredOutcomes' (@($choices|Where-Object {$_ -cnotin @('REQUESTED','VERIFIED','FAILED')}).Count -eq 0)
    }finally{
        Check 'AdminVerify.Expectation.ActualCancel' ((BoundControl 'btnCancelExpectation' 'Click' '' 'frmActionPathExpectation') -ceq 'DELIVERED')
        CloseRecordingViewer
    }
    Check 'AdminVerify.Expectation.CancelPreservesJournal' (BoundSame $pins (BoundPins $journalRoot))
}
