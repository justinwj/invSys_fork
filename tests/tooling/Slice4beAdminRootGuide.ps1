# Separate real root journeys feed the existing guide/evaluation form handlers.
function Test-AdminRootGuide($Fixture,$HostBook) {
    . (Join-Path $PSScriptRoot 'Slice4beRecordingFixture.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beRecordingEvaluation.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beGuideTestActions.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beGeneralSettingsGuidePaths.ps1')
    function Hash([string]$Path){$stream=[IO.File]::Open($Path,'Open','Read','ReadWrite');try{(Get-FileHash -InputStream $stream).Hash}finally{$stream.Dispose()}}
    $steps=@(@('Add Warehouse Root','ADMIN_WAREHOUSE_ROOT_ADD','REMEMBERED_REFRESHED'),@('Add Warehouse Root','ADMIN_WAREHOUSE_ROOT_ADD','REMEMBERED_REFRESHED'))
    $runs=@();$inventory=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Data.Inventory.xlsb');$inventoryHash=Hash $inventory
    try{
        SelectTarget $Fixture
        $authorized=[bool](Run 'invSys.Core.xlam' 'modAuth.CanPerform' @('ACTION_PATH_MAINT','config-admin',$Fixture.Warehouse,'S1'))
        if(-not $authorized){throw 'Explicit guide-maintenance fixture grant unavailable; not product RED.'}
        Check 'WarehouseRoot.Recorded.ExplicitGuidePermission' $authorized
        $HostBook.Windows.Item(1).Visible=$true;$HostBook.Activate();$excel.Visible=$true
        Check 'WarehouseRoot.Recorded.VisibleHost' ([bool]$excel.Visible -and [bool]$HostBook.Windows.Item(1).Visible)
        if(-not [bool](Probe 'SetPolicy' @('Capture'))){throw 'Root capture policy prerequisite failed.'}
        if(-not [bool](Run 'invSys.Admin.xlam' 'modAdminConsole.PublishReadFixtureForTest')){throw 'Initial Viewer publication unavailable.'}
        foreach($kind in @('Source','Observed','Interrupted','Unconfirmed')){
            CloseRecordingViewer;SelectTarget $Fixture;OpenRecordingViewer
            $prior=BoundPins $journalRoot;$oldActivity=ActivityPins;$before=@(Get-Slice4beActivityFiles $Fixture)
            Check ('WarehouseRoot.Recorded.'+$kind+'.ActualStart') ((RecordingControl 'Start Recording' 'Click') -ceq 'DELIVERED')
            $starts=@(Get-ChildItem $journalRoot -File -Filter '*.json'|Where-Object {-not $prior.ContainsKey($_.FullName)}|ForEach-Object {Get-Content $_.FullName -Raw|ConvertFrom-Json}|Where-Object RecordType -CEQ 'Start')
            if($starts.Count -ne 1){throw 'One durable root Start required.'}
            $sequence=$starts[0].SequenceId;$incomplete=$kind -ceq 'Interrupted';$success=$kind -cin @('Source','Observed')
            for($ordinal=1;$ordinal -le 2;$ordinal++){
                $fault=$incomplete -and $ordinal -eq 2;$mode=if($fault){'OutcomeWriteFailure'}else{''}
                $unconfirmed=$kind -ceq 'Unconfirmed' -and $ordinal -eq 2
                [void](Probe 'ArmBoundary' @($mode,'','',''))
                try{
                    $facts=([string](Invoke 'InvokeRoot' @((Join-Path $runRoot ('root-guide-'+$kind+'-'+$ordinal)),$unconfirmed,'',$false))) -split '\|'
                    $faults=[int](Probe 'WriteFailures')
                }finally{[void](Probe 'ArmBoundary' @('','','',''))}
                if($fault -and $faults -ne 1){throw 'Real root outcome-write fault was not reached.'}
                Check ('WarehouseRoot.Recorded.'+$kind+'.OwnerRunsOnce.'+$ordinal) ($facts.Count -eq 14 -and $facts[0] -ceq '1' -and $facts[1] -ceq 'True' -and $facts[2] -ceq $(if($unconfirmed){'0'}else{'1'}) -and $facts[3] -ceq '1' -and $facts[4] -ceq '1' -and $facts[8] -ceq 'True' -and $facts[11] -ceq '1' -and [int](Invoke 'RootTrackingNotices') -eq $(if($fault){1}else{0}))
                if($unconfirmed){Check 'WarehouseRoot.Recorded.Unconfirmed.ActualWriteRefusalAndWarning' ($facts[5] -ceq '1' -and $facts[13] -ceq 'Warehouse root preference could not be confirmed. Warehouse directory refreshed.')}
            }
            $records=@(Get-Slice4beActivityFiles $Fixture|Where-Object {$_ -cnotin $before}|ForEach-Object {Get-Content $_ -Raw|ConvertFrom-Json}|Where-Object ControlId -CEQ 'ADMIN_WAREHOUSE_ROOT_ADD')
            $count=if($incomplete){3}else{4};$valid=$records.Count -eq $count
            foreach($record in $records){$valid=$valid -and $record.SequenceId -ceq $sequence -and $record.CatalogVersion -eq 40 -and $record.WarehouseId -ceq $Fixture.Warehouse -and $record.UserId -ceq 'config-admin' -and $record.DataEffect -ceq 'Unknown' -and @($record.SourceEventRefs).Count -eq 0}
            foreach($ordinal in 1,2){$expected=if($incomplete -and $ordinal -eq 2){1}else{2};$outcome=if($kind -ceq 'Unconfirmed' -and $ordinal -eq 2){'UNCONFIRMED_REFRESHED'}else{'REMEMBERED_REFRESHED'};$part=@($records|Where-Object Ordinal -EQ $ordinal);$valid=$valid -and $part.Count -eq $expected -and @($part|Where-Object OutcomeCode -CEQ 'REQUESTED').Count -eq 1 -and @($part|Where-Object OutcomeCode -CEQ $outcome).Count -eq ($expected-1)}
            Check ('WarehouseRoot.Recorded.'+$kind+'.ExactActionsAndContext') $valid
            if(-not $incomplete){Check ('WarehouseRoot.Recorded.'+$kind+'.ActualStop') ((RecordingControl 'Stop Recording' 'Click') -ceq 'DELIVERED')}
            $closed=@(RecordingJournal $sequence|Where-Object RecordType -CEQ 'Close');$valid=$closed.Count -eq 1
            if($valid){$valid=$closed[0].ActionCount -eq 2 -and @($closed[0].Observations).Count -eq $count -and $closed[0].Lifecycle -ceq $(if($incomplete){'Incomplete'}else{'Stopped'})}
            if($valid -and $incomplete){$valid=$closed[0].ReasonCode -ceq 'TRACKING_UNAVAILABLE'}
            Check ('WarehouseRoot.Recorded.'+$kind+'.ExactJournalLifecycle') $valid
            Check ('WarehouseRoot.Recorded.'+$kind+'.ImmutableJournalChain') (JournalChain $sequence ($count+2))
            Check ('WarehouseRoot.Recorded.'+$kind+'.PriorActivityImmutable') (PinsRetained $oldActivity)
            if($closed.Count -ne 1){throw 'Closed root recording unavailable.'}
            $runs+=@{Kind=$kind;Journal=$closed[0];Records=$records;Success=$success;Incomplete=$incomplete}
        }
        CloseRecordingViewer;SelectTarget $Fixture
        if(-not [bool](Run 'invSys.Admin.xlam' 'modAdminConsole.PublishReadFixtureForTest')){throw 'Recorded root publication failed.'}
        $publication=Get-Content (Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Snapshot.Events.json')) -Raw|ConvertFrom-Json
        $lines=@($publication.Groups|Where-Object Source -CEQ 'Activity'|ForEach-Object {$_.Lines});$exact=$true
        foreach($run in $runs){foreach($record in $run.Records){$matches=@($lines|Where-Object RecordId -CEQ $record.RecordId);$exact=$exact -and $matches.Count -eq 1;if($matches.Count -eq 1){$exact=$exact -and $matches[0].ActivityId -ceq $record.ActivityId -and $matches[0].OutcomeCode -ceq $record.OutcomeCode}}}
        Check 'WarehouseRoot.Recorded.ExactPublishedRecords' $exact
        Test-GeneralSettingsGuidePaths $Fixture $runs $steps -RootActions
        Check 'WarehouseRoot.Recorded.InventoryAndCustomValuePreserved' ((Hash $inventory) -ceq $inventoryHash -and [bool](Invoke 'Preserved' @('Text')))
    }finally{[void](Probe 'ArmBoundary' @('','','',''));CloseRecordingViewer;SelectTarget $Fixture}
}
