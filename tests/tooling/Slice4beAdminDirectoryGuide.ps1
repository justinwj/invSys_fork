# Independent real Directory journeys feed the existing guide/evaluation handlers.
function Test-AdminDirectoryGuide($Fixture,$HostBook) {
    function HostFacts([string]$Stage){
        [pscustomobject]@{Stage=$Stage;AppVisible=[bool]$excel.Visible;ArgumentWindows=[int]$HostBook.Windows.Count;ArgumentIsAddin=[bool]$HostBook.IsAddin;OwnerFacts=[string](Invoke 'HostWindowFacts')}|
            ConvertTo-Json -Compress|Add-Content (Join-Path $reportRoot 'directory-guide-host-facts.jsonl')
    }
    HostFacts 'Entry'
    . (Join-Path $PSScriptRoot 'Slice4beRecordingFixture.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beRecordingEvaluation.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beGuideTestActions.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beGeneralSettingsGuidePaths.ps1')
    function Hash([string]$Path){$stream=[IO.File]::Open($Path,'Open','Read','ReadWrite');try{(Get-FileHash -InputStream $stream).Hash}finally{$stream.Dispose()}}
    $steps=@(@('View Warehouses','ADMIN_DIRECTORY_OPEN','OPENED'),@('View Warehouses','ADMIN_DIRECTORY_OPEN','OPENED'))
    $runs=@();$inventory=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Data.Inventory.xlsb');$inventoryHash=Hash $inventory
    try{
        SelectTarget $Fixture
        HostFacts 'AfterTargetSelection'
        $authorized=[bool](Run 'invSys.Core.xlam' 'modAuth.CanPerform' @('ACTION_PATH_MAINT','config-admin',$Fixture.Warehouse,'S1'))
        if(-not $authorized){throw 'Explicit guide-maintenance fixture grant unavailable; not product RED.'}
        Check 'AdminDirectory.Recorded.ExplicitGuidePermission' $authorized
        HostFacts 'BeforeVisibility'
        $HostBook.Windows.Item(1).Visible=$true
        HostFacts 'AfterWindowVisibility'
        $HostBook.Activate()
        HostFacts 'AfterCapturedHostActivation'
        if(-not $excel.Visible){$excel.Visible=$true}
        HostFacts 'AfterApplicationVisibility'
        Check 'AdminDirectory.Recorded.VisibleHost' ([bool]$excel.Visible -and [bool]$HostBook.Windows.Item(1).Visible)
        if(-not [bool](Probe 'SetPolicy' @('Capture'))){throw 'Directory capture policy prerequisite failed.'}
        if(-not [bool](Run 'invSys.Admin.xlam' 'modAdminConsole.PublishReadFixtureForTest')){throw 'Initial Viewer publication unavailable.'}
        foreach($kind in @('Source','Observed','Interrupted')){
            CloseRecordingViewer;SelectTarget $Fixture;OpenRecordingViewer
            $prior=BoundPins $journalRoot;$oldActivity=ActivityPins;$before=@(Get-Slice4beActivityFiles $Fixture)
            Check ('AdminDirectory.Recorded.'+$kind+'.ActualStart') ((RecordingControl 'Start Recording' 'Click') -ceq 'DELIVERED')
            $starts=@(Get-ChildItem $journalRoot -File -Filter '*.json'|Where-Object {-not $prior.ContainsKey($_.FullName)}|ForEach-Object {Get-Content $_.FullName -Raw|ConvertFrom-Json}|Where-Object RecordType -CEQ 'Start')
            if($starts.Count -ne 1){throw 'One durable Directory Start required.'}
            $sequence=$starts[0].SequenceId;$success=$kind -cne 'Interrupted'
            for($ordinal=1;$ordinal -le 2;$ordinal++){
                $fault=-not $success -and $ordinal -eq 2;$mode=if($fault){'OutcomeWriteFailure'}else{''}
                [void](Probe 'ArmBoundary' @($mode,'','',''))
                try{
                    $values=([string](Invoke 'Invoke')) -split '\|'
                    $faults=[int](Probe 'WriteFailures')
                }finally{[void](Probe 'ArmBoundary' @('','','',''))}
                if($fault -and $faults -ne 1){throw 'Real Directory outcome-write fault was not reached.'}
                $notices=if($fault){'1'}else{'0'}
                Check ('AdminDirectory.Recorded.'+$kind+'.OwnerOpensOnce.'+$ordinal) ($values.Count -eq 6 -and $values[0] -ceq '1' -and $values[1] -ceq 'True' -and $values[2] -ceq 'True' -and $values[3] -ceq '1' -and $values[4] -ceq '0' -and [string](Invoke 'TrackingNotices') -ceq $notices)
            }
            $records=@(Get-Slice4beActivityFiles $Fixture|Where-Object {$_ -cnotin $before}|ForEach-Object {Get-Content $_ -Raw|ConvertFrom-Json}|Where-Object ControlId -CEQ 'ADMIN_DIRECTORY_OPEN')
            $count=if($success){4}else{3};$valid=$records.Count -eq $count
            foreach($record in $records){$valid=$valid -and $record.SequenceId -ceq $sequence -and $record.CatalogVersion -eq 39 -and $record.WarehouseId -ceq $Fixture.Warehouse -and $record.UserId -ceq 'config-admin' -and $record.DataEffect -ceq 'Unknown' -and @($record.SourceEventRefs).Count -eq 0}
            foreach($ordinal in 1,2){$expected=if(-not $success -and $ordinal -eq 2){1}else{2};$part=@($records|Where-Object Ordinal -EQ $ordinal);$valid=$valid -and $part.Count -eq $expected -and @($part|Where-Object OutcomeCode -CEQ 'REQUESTED').Count -eq 1 -and @($part|Where-Object OutcomeCode -CEQ 'OPENED').Count -eq ($expected-1)}
            Check ('AdminDirectory.Recorded.'+$kind+'.ExactActionsAndContext') $valid
            if($success){Check ('AdminDirectory.Recorded.'+$kind+'.ActualStop') ((RecordingControl 'Stop Recording' 'Click') -ceq 'DELIVERED')}
            $closed=@(RecordingJournal $sequence|Where-Object RecordType -CEQ 'Close');$valid=$closed.Count -eq 1
            if($valid){$valid=$closed[0].ActionCount -eq 2 -and @($closed[0].Observations).Count -eq $count -and $closed[0].Lifecycle -ceq $(if($success){'Stopped'}else{'Incomplete'})}
            if($valid -and -not $success){$valid=$closed[0].ReasonCode -ceq 'TRACKING_UNAVAILABLE'}
            Check ('AdminDirectory.Recorded.'+$kind+'.ExactJournalLifecycle') $valid
            Check ('AdminDirectory.Recorded.'+$kind+'.ImmutableJournalChain') (JournalChain $sequence ($count+2))
            Check ('AdminDirectory.Recorded.'+$kind+'.PriorActivityImmutable') (PinsRetained $oldActivity)
            if($closed.Count -ne 1){throw 'Closed Directory recording unavailable.'}
            $runs+=@{Kind=$kind;Journal=$closed[0];Records=$records;Success=$success;Incomplete=(-not $success)}
        }
        CloseRecordingViewer;SelectTarget $Fixture
        if(-not [bool](Run 'invSys.Admin.xlam' 'modAdminConsole.PublishReadFixtureForTest')){throw 'Recorded Directory publication failed.'}
        $publication=Get-Content (Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Snapshot.Events.json')) -Raw|ConvertFrom-Json
        $lines=@($publication.Groups|Where-Object Source -CEQ 'Activity'|ForEach-Object {$_.Lines});$exact=$true
        foreach($run in $runs){foreach($record in $run.Records){$matches=@($lines|Where-Object RecordId -CEQ $record.RecordId);$exact=$exact -and $matches.Count -eq 1;if($matches.Count -eq 1){$exact=$exact -and $matches[0].ActivityId -ceq $record.ActivityId -and $matches[0].OutcomeCode -ceq $record.OutcomeCode}}}
        Check 'AdminDirectory.Recorded.ExactPublishedRecords' $exact
        Test-GeneralSettingsGuidePaths $Fixture $runs $steps -DirectoryActions
        Check 'AdminDirectory.Recorded.InventoryAndCustomValuePreserved' ((Hash $inventory) -ceq $inventoryHash -and [bool](Invoke 'Preserved' @('Text')))
    }finally{[void](Probe 'ArmBoundary' @('','','',''));CloseRecordingViewer;SelectTarget $Fixture}
}
