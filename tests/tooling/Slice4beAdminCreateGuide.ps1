# Four independent real creation/publication journeys; no fabricated outcome or journal.
function Test-AdminCreateGuide($Fixture,$HostBook){
    . (Join-Path $PSScriptRoot 'Slice4beRecordingFixture.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beRecordingEvaluation.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beGuideTestActions.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beGeneralSettingsGuidePaths.ps1')
    function Hash([string]$Path){$stream=[IO.File]::Open($Path,'Open','Read','ReadWrite');try{(Get-FileHash -InputStream $stream).Hash}finally{$stream.Dispose()}}
    $steps=@(@('Create','ADMIN_CREATE_SUBMIT','CREATED_PUBLISHED'),@('Create','ADMIN_CREATE_SUBMIT','CREATED_PUBLISHED'))
    $runs=@();$inventory=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Data.Inventory.xlsb');$inventoryHash=Hash $inventory
    try{
        SelectTarget $Fixture
        $registered=[string](Probe 'Definition' @('ADMIN_CREATE_SUBMIT',41)) -cne ''
        Check 'Create.Recorded.CaptureControlRegistered' $registered
        if(-not $registered){return}
        if(-not [bool](Run 'invSys.Core.xlam' 'modAuth.CanPerform' @('ACTION_PATH_MAINT','config-admin',$Fixture.Warehouse,'S1'))){throw 'Explicit Create guide permission unavailable.'}
        if(-not [bool](Probe 'SetPolicy' @('CreateOnly'))){throw 'Create capture policy prerequisite failed.'}
        $HostBook.Windows.Item(1).Visible=$true;$HostBook.Activate();$excel.Visible=$true
        if(-not [bool](Run 'invSys.Admin.xlam' 'modAdminConsole.PublishReadFixtureForTest')){throw 'Initial Create Viewer publication unavailable.'}
        foreach($kind in @('Source','Observed','Interrupted','Unconfirmed')){
            CloseRecordingViewer;SelectTarget $Fixture;OpenRecordingViewer
            $prior=BoundPins $journalRoot;$oldActivity=ActivityPins;$before=@(Get-Slice4beActivityFiles $Fixture)
            Check ('Create.Recorded.'+$kind+'.ActualStart') ((RecordingControl 'Start Recording' 'Click') -ceq 'DELIVERED')
            $starts=@(Get-ChildItem $journalRoot -File -Filter '*.json'|Where-Object{-not $prior.ContainsKey($_.FullName)}|ForEach-Object{Get-Content $_.FullName -Raw|ConvertFrom-Json}|Where-Object RecordType -CEQ 'Start')
            if($starts.Count -ne 1){throw 'One durable Create recording Start required.'}
            $sequence=$starts[0].SequenceId;$incomplete=$kind -ceq 'Interrupted';$success=$kind -cin @('Source','Observed')
            for($ordinal=1;$ordinal -le 2;$ordinal++){
                [void](Invoke 'OpenForm' @($HostBook.Name))
                $id='CG'+[guid]::NewGuid().ToString('N').Substring(0,8)
                $root=Join-Path $runRoot ('guide-create-'+$kind+'-'+$ordinal);$share=Join-Path $runRoot ('guide-share-'+$kind+'-'+$ordinal)
                New-Item -ItemType Directory -Path $share|Out-Null
                $partial=$kind -ceq 'Unconfirmed' -and $ordinal -eq 2
                if($partial){[IO.File]::WriteAllText((Join-Path $share $id),'Generated publication obstruction')}
                [void](Invoke 'Prepare' @($id,$root,$share,'Training'));[void](Invoke 'ResetTrackingNotices')
                $fault=$incomplete -and $ordinal -eq 2
                [void](Probe 'ArmBoundary' @($(if($fault){'OutcomeWriteFailure'}else{''}),'','',''))
                try{
                    [void](Invoke 'Create');$facts=[string](Invoke 'Facts');$faults=[int](Probe 'WriteFailures')
                    if($fault -and $faults -ne 1){throw 'Real Create outcome-write fault not reached.'}
                    Check ('Create.Recorded.'+$kind+'.OwnerTruth.'+$ordinal) ($facts -ceq $(if($partial){'True|False|Retry Publish|True|1|1'}else{'True|True|Close|True|1|1'}))
                    Check ('Create.Recorded.'+$kind+'.NoticeTruth.'+$ordinal) ([int](Invoke 'TrackingNotices') -eq $(if($fault){1}else{0}))
                }finally{[void](Probe 'ArmBoundary' @('','','',''));[void](Invoke 'CloseForm')}
            }
            $records=@(Get-Slice4beActivityFiles $Fixture|Where-Object{$_ -cnotin $before}|ForEach-Object{Get-Content $_ -Raw|ConvertFrom-Json}|Where-Object ControlId -CEQ 'ADMIN_CREATE_SUBMIT')
            $count=if($incomplete){3}else{4};$valid=$records.Count -eq $count
            foreach($record in $records){$valid=$valid -and $record.SequenceId -ceq $sequence -and $record.CatalogVersion -eq 41 -and $record.WarehouseId -ceq $Fixture.Warehouse -and $record.UserId -ceq 'config-admin' -and $record.DataEffect -ceq 'Unknown' -and @($record.SourceEventRefs).Count -eq 0}
            foreach($ordinal in 1,2){$expected=if($incomplete -and $ordinal -eq 2){1}else{2};$outcome=if($kind -ceq 'Unconfirmed' -and $ordinal -eq 2){'CREATED_PUBLISH_FAILED'}else{'CREATED_PUBLISHED'};$part=@($records|Where-Object Ordinal -EQ $ordinal);$valid=$valid -and $part.Count -eq $expected -and @($part|Where-Object OutcomeCode -CEQ 'REQUESTED').Count -eq 1 -and @($part|Where-Object OutcomeCode -CEQ $outcome).Count -eq ($expected-1)}
            Check ('Create.Recorded.'+$kind+'.ExactActionsAndContext') $valid
            if(-not $incomplete){Check ('Create.Recorded.'+$kind+'.ActualStop') ((RecordingControl 'Stop Recording' 'Click') -ceq 'DELIVERED')}
            $closed=@(RecordingJournal $sequence|Where-Object RecordType -CEQ 'Close');$valid=$closed.Count -eq 1
            if($valid){$valid=$closed[0].ActionCount -eq 2 -and @($closed[0].Observations).Count -eq $count -and $closed[0].Lifecycle -ceq $(if($incomplete){'Incomplete'}else{'Stopped'})}
            if($valid -and $incomplete){$valid=$closed[0].ReasonCode -ceq 'TRACKING_UNAVAILABLE'}
            Check ('Create.Recorded.'+$kind+'.ExactJournalLifecycle') $valid
            Check ('Create.Recorded.'+$kind+'.ImmutableJournalChain') (JournalChain $sequence ($count+2))
            Check ('Create.Recorded.'+$kind+'.PriorActivityImmutable') (PinsRetained $oldActivity)
            if($closed.Count -ne 1){throw 'Closed Create recording unavailable.'}
            $runs+=@{Kind=$kind;Journal=$closed[0];Records=$records;Success=$success;Incomplete=$incomplete}
        }
        CloseRecordingViewer;SelectTarget $Fixture
        if(-not [bool](Run 'invSys.Admin.xlam' 'modAdminConsole.PublishReadFixtureForTest')){throw 'Recorded Create publication failed.'}
        $publication=Get-Content (Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Snapshot.Events.json')) -Raw|ConvertFrom-Json
        $lines=@($publication.Groups|Where-Object Source -CEQ 'Activity'|ForEach-Object{$_.Lines});$exact=$true
        foreach($run in $runs){foreach($record in $run.Records){$matches=@($lines|Where-Object RecordId -CEQ $record.RecordId);$exact=$exact -and $matches.Count -eq 1;if($matches.Count -eq 1){$exact=$exact -and $matches[0].ActivityId -ceq $record.ActivityId -and $matches[0].OutcomeCode -ceq $record.OutcomeCode}}}
        Check 'Create.Recorded.ExactPublishedRecords' $exact
        Test-GeneralSettingsGuidePaths $Fixture $runs $steps -CreateActions
        Check 'Create.Recorded.SourceInventoryPreserved' ((Hash $inventory) -ceq $inventoryHash)
    }finally{[void](Probe 'ArmBoundary' @('','','',''));[void](Invoke 'CloseForm');CloseRecordingViewer;SelectTarget $Fixture}
}
