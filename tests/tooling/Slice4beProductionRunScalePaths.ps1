# Independent worksheet recordings; reuse the real Scale owner and guide controls.
function Test-ProductionRunScalePaths($Fixture) {
    . (Join-Path $PSScriptRoot 'Slice4beRecordingFixture.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beRecordingEvaluation.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beEvaluationContracts.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beGuideTestActions.ps1')
    function Retained($Before,$After){foreach($file in $Before.Keys){if(-not $After.ContainsKey($file) -or $After[$file] -cne $Before[$file]){return $false}};return $true}
    $runs=@();$steps=@(@('Apply Scale','PRODUCTION_RUN_SCALE','STAGED'),@('Apply Scale again','PRODUCTION_RUN_SCALE','STAGED'))
    try {
        foreach($kind in @('Source','Observed','Interrupted')){
            # Keep the session captured when the outer fixture opened Production.
            # Only the final Interrupted run deliberately signs out.
            CloseRecordingViewer
            if(-not [bool](Probe 'RunScaleSuccessPrepare' @($canary))){throw 'Successful worksheet fixture unavailable; not behavioral RED.'}
            $label='WorksheetScaleRecorded.'+$kind
            Check ($label+'.WorksheetBranchAndDisabledFlags') ([bool](Probe 'RunScaleSheetBranch') -and -not [bool](Run 'invSys.Core.xlam' 'modConfig.GetBool' @('DesignsEnabled',$true)) -and -not [bool](Run 'invSys.Core.xlam' 'modConfig.GetBool' @('FF_DesignsEnabled',$true)))
            OpenRecordingViewer;$prior=BoundPins $journalRoot;$activity=ActivityPins
            Check ($label+'.ActualStart') ((RecordingControl 'Start Recording' 'Click') -ceq 'DELIVERED')
            $starts=@(Get-ChildItem -LiteralPath $journalRoot -Filter '*.json' -File|Where-Object {-not $prior.ContainsKey($_.FullName)}|ForEach-Object {Get-Content $_.FullName -Raw|ConvertFrom-Json}|Where-Object RecordType -CEQ 'Start')
            if($starts.Count -ne 1){throw 'One durable Start required.'}
            $sequence=$starts[0].SequenceId;$all=@()
            for($ordinal=1;$ordinal -le 2;$ordinal++){
                # Reset observation counters only; leave all owner work and results intact.
                [void](Probe 'RunScaleReset');$before=@(Files);$decoy.Activate()
                $notice=[string](Probe 'RunFaultAct' @('SCALE'))
                Check ($label+'.'+$ordinal+'.ActualOwnerAndSuccess') ($notice.StartsWith('Loaded recipe at 200.% and prepared output rows: ') -and [bool](Probe 'RunFaultGuards') -and [string](Probe 'RunScaleEvidence') -ceq '1|0|False|1|Captured|0' -and [string](Probe 'RunFaultEvidence') -ceq '0|False|False')
                foreach($fact in @('ScaledChooser','ExactInputIdentity','OutputPrepared','RecipeChooser_generated','InventoryPalette_generated','ProductionOutput')){Check ($label+'.'+$ordinal+'.'+$fact) ([bool](Probe 'RunScaleSuccessFact' @($fact)))}
                Check ($label+'.'+$ordinal+'.CanonicalSourcePreserved') ([bool](Probe 'RunWorksheetSource'))
                Pair $before 'SCALE' 'STAGED' ($label+'.'+$ordinal)
                $records=@(Files|Where-Object {$_ -cnotin $before}|ForEach-Object {Get-Content $_ -Raw|ConvertFrom-Json})
                $valid=$records.Count -eq 2
                foreach($record in $records){$valid=$valid -and $record.SequenceId -ceq $sequence -and $record.Ordinal -eq $ordinal}
                Check ($label+'.'+$ordinal+'.ExactRecordingSequenceAndOrdinal') $valid
                $all+=@($records)
                if($kind -ceq 'Interrupted'){[void](Run 'invSys.Core.xlam' 'modAuth.SignOut');break}
            }
            if($kind -cne 'Interrupted'){Check ($label+'.ActualStop') ((RecordingControl 'Stop Recording' 'Click') -ceq 'DELIVERED')}
            $closed=@(RecordingJournal $sequence|Where-Object RecordType -CEQ 'Close')
            $count=if($kind -ceq 'Interrupted'){1}else{2};$valid=$closed.Count -eq 1
            if($valid){$valid=$closed[0].ActionCount -eq $count -and @($closed[0].Observations).Count -eq 2*$count -and $closed[0].Lifecycle -ceq $(if($kind -ceq 'Interrupted'){'Incomplete'}else{'Stopped'})}
            if($valid -and $kind -ceq 'Interrupted'){$valid=$closed[0].ReasonCode -ceq 'SESSION_CHANGED'}
            Check ($label+'.ExactJournalLifecycle') $valid
            Check ($label+'.ImmutableJournalChain') (JournalChain $sequence (2*$count+2))
            Check ($label+'.PriorActivityImmutable') (PinsRetained $activity)
            if($closed.Count -ne 1){throw 'Closed recording prerequisite missing.'}
            $runs+=@{Kind=$kind;Journal=$closed[0];Records=$all;Success=($kind -cne 'Interrupted');Incomplete=($kind -ceq 'Interrupted')}
        }
        CloseRecordingViewer;SelectTarget $Fixture 'config-admin'
        if(-not [bool](Run 'invSys.Admin.xlam' 'modAdminConsole.PublishReadFixtureForTest')){throw 'Actual publication prerequisite failed.'}
        $publication=Get-Content (Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Snapshot.Events.json')) -Raw|ConvertFrom-Json
        OpenRecordingViewer
        foreach($run in $runs){
            foreach($record in $run.Records){
                $group=@($publication.Groups|Where-Object {$_.Source -ceq 'Activity' -and $_.SourceId -ceq $record.ActivityId})
                $lines=@($group|ForEach-Object Lines|Where-Object RecordId -CEQ $record.RecordId)
                Check ('WorksheetScalePublished.'+$run.Kind+'.'+$record.Ordinal+'.'+$record.OutcomeCode) ($lines.Count -eq 1 -and ($lines[0]|ConvertTo-Json -Depth 20 -Compress) -ceq ($record|ConvertTo-Json -Depth 20 -Compress))
            }
        }
        foreach($record in @($runs|Where-Object Kind -CEQ 'Observed')[0].Records|Where-Object OutcomeCode -CEQ 'STAGED'){
            $id=[string]$record.ActivityId
            Check ('WorksheetScaleDetail.'+$record.Ordinal) ([bool](Run 'invSys.Operations.xlam' 'modInventoryViewer.PublishedReadActionForTest' @('SelectSource',$id)) -and [bool](Run 'invSys.Operations.xlam' 'modInventoryViewer.PublishedReadDetailForTest' @('Source event / activity ID',$id)))
        }
        $training=BoundPins $journalRoot
        . (Join-Path $PSScriptRoot 'Slice4beGeneralSettingsGuidePaths.ps1')
        Test-GeneralSettingsGuidePaths $Fixture $runs $steps -WorksheetScale
        Check 'WorksheetScaleRecorded.TrainingEvidenceImmutableThroughEvaluation' (Retained $training (BoundPins $journalRoot))
    }finally{CloseRecordingViewer;SelectTarget $Fixture 'config-producer'}
}
