# Versioned guide expectation uses exact independent published Domain evidence.
function Test-AdminSeedAppliedGuide($Fixture,$Runs,[string]$InventoryPath) {
    . (Join-Path $PSScriptRoot 'Slice4beEvaluationEvidence.ps1')
    function Deliver([string]$Form,[string]$Control,[string]$Action,[string]$Value=''){
        if((BoundControl $Control $Action $Value $Form) -cne 'DELIVERED'){throw ('Applied Seed guide control unavailable: '+$Control)}
    }
    function Pass([string]$Name,[bool]$Value){Check ('SeedAppliedGuide.'+$Name) $Value|Out-Host}
    function Retained($Before){foreach($path in $Before.Keys){if(-not(Test-Path -LiteralPath $path) -or (Hash $path) -cne $Before[$path]){return $false}};return $true}
    $source=@($Runs|Where-Object Kind -CEQ 'Source')[0]
    $guideRoot=Join-Path $journalRoot 'Guides';$eventsPath=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Snapshot.Events.json')
    $originals=@(Get-ChildItem $guideRoot -File -Filter '*.json'|ForEach-Object {ReadGuideExpectationRecord $_.FullName}|Where-Object {$_.SourceRun.RecordId -ceq $source.Journal.RecordId})
    if($originals.Count -ne 1){throw 'One original Seed guide required.'}
    $original=$originals[0];$training=BoundPins $journalRoot;$activity=ActivityPins
    $inventoryHash=Hash $InventoryPath;$publicationHash=Hash $eventsPath
    $publication=Get-Content $eventsPath -Raw|ConvertFrom-Json
    try{
        CloseRecordingViewer;SelectTarget $Fixture 'config-admin';OpenRecordingViewer
        if((BoundLibrary 'Open') -cne 'DELIVERED'){throw 'Recording library unavailable.'}
        BoundOpen;BoundSelect $original
        Deliver 'frmActionPathLibrary' 'btnEditPublishedGuide' 'Click'
        Deliver 'frmActionPathGuide' 'btnGuideExpectedConclusion' 'Click'
        Deliver 'frmActionPathExpectation' 'cboTerminalKind' 'Write' 'SourceEventsApplied'
        Deliver 'frmActionPathExpectation' 'btnUseExpectation' 'Click'
        Deliver 'frmActionPathGuide' 'txtGuideInstructions' 'Write' "Seed built-in demo stock, then the selected training CSV.`r`nRead the owner results and wait for publication. This version requires the terminal command's exact published inventory event."
        Deliver 'frmActionPathGuide' 'btnSaveGuide' 'Click'
        $guide=ReadGuideExpectationRecord (Join-Path $guideRoot ($original.ActionPathId+'.2.json'))
        if($null -eq $guide){throw 'Saved version-two Seed guide unavailable.'}
        Pass 'ActualEditAppendsAppliedExpectation' ($guide.Version -eq 2 -and $guide.PreviousRecordId -ceq $original.RecordId -and $guide.PreviousSha256 -ceq $original.ContentSha256 -and $guide.ExpectedConclusion.TerminalKind -ceq 'SourceEventsApplied' -and $guide.ExpectedConclusion.TerminalStepId -ceq $original.ExpectedConclusion.TerminalStepId)
        Pass 'OriginalGuideAndSourcePreserved' ((Retained $training) -and (PinsRetained $activity) -and ($guide.Observations|ConvertTo-Json -Depth 20 -Compress) -ceq ($original.Observations|ConvertTo-Json -Depth 20 -Compress) -and ($guide.SourceRun|ConvertTo-Json -Depth 12 -Compress) -ceq ($original.SourceRun|ConvertTo-Json -Depth 12 -Compress))
        foreach($kind in @('Observed','Interrupted')){
            $run=@($Runs|Where-Object Kind -CEQ $kind)[0];$eventId=$run.EventIds[1]
            $groups=@($publication.Groups|Where-Object {$_.Source -ceq 'Inventory' -and $_.SourceId -ceq $eventId})
            Pass ($kind+'.ExactDomainEventAvailable') ($groups.Count -eq 1)
            CloseRecordingViewer;SelectTarget $Fixture 'config-reader';OpenRecordingViewer
            if((BoundLibrary 'Open') -cne 'DELIVERED' -or (BoundLibrary 'Select' $run.Journal.ActionPathId) -cne 'SELECTED'){throw 'Independent Seed recording unavailable.'}
            BoundOpen;BoundSelect $guide
            Deliver 'frmActionPathLibrary' 'btnUseGuideForRun' 'Click';Deliver 'frmActionPathLibrary' 'btnCloseGuides' 'Click'
            $before=@(EvaluationFiles|ForEach-Object FullName)
            Deliver 'frmActionPaths' 'btnEvaluatePath' 'Click'
            $fresh=@(EvaluationFiles|Where-Object {$_.FullName -cnotin $before})
            if($fresh.Count -ne 1){throw 'One saved applied-guide result required.'}
            $result=Get-Content $fresh[0].FullName -Raw|ConvertFrom-Json
            $expected=if($run.Incomplete){'Incomplete'}else{'Concluded'};$reason=if($run.Incomplete){'CAPTURE_INCOMPLETE'}else{'SOURCE_APPLIED'}
            Pass ($kind+'.HonestConclusion') ($result.ResultState -ceq $expected -and $reason -cin @($result.ReasonCodes) -and $result.ExpectedConclusion.TerminalKind -ceq 'SourceEventsApplied')
            Pass ($kind+'.ExactGuideJournalAndPublication') ($result.Guide.ContentSha256 -ceq $guide.ContentSha256 -and $result.Guide.Version -eq 2 -and $result.JournalRecordId -ceq $run.Journal.RecordId -and $result.JournalSha256 -ceq $run.Journal.ContentSha256 -and $result.Publication.PublicationId -ceq $publication.PublicationId -and $result.Publication.ContentSha256 -ceq $publication.ContentSha256)
            $refs=@($result.TerminalSources)
            if($run.Success){
                $exact=$refs.Count -eq 1 -and $groups.Count -eq 1
                if($exact){
                    $ref=$refs[0];$body=[ordered]@{Lines=@($groups[0].Lines)}|ConvertTo-Json -Depth 16 -Compress
                    $keys=@($groups[0].Lines|ForEach-Object {$_.System_Key}|Select-Object -Unique)
                    $exact=$ref.EventId -ceq $eventId -and $ref.SourceKind -ceq 'Inventory' -and $ref.WarehouseId -ceq $Fixture.Warehouse -and $ref.OwnerStatus -ceq 'Applied' -and $ref.LineCount -eq 1 -and $ref.LinesSha256 -ceq (EvaluationSha $body) -and @($ref.SystemKeys).Count -eq 1 -and $keys.Count -eq 1 -and $ref.SystemKeys[0] -ceq $keys[0]
                }
                Pass ($kind+'.ExactPublishedOwnerIdentityAndLines') $exact
            }else{
                Pass ($kind+'.PublishedEventCannotReplaceMissingObservation') ($refs.Count -eq 0 -and @($result.FailedSteps).Count -eq 1 -and $result.FailedSteps[0] -ceq $guide.ExpectedConclusion.TerminalStepId)
            }
            Deliver 'frmActionPaths' 'btnViewActionPath' 'Click'
            Deliver 'frmActionPathView' 'cboActionPathView' 'Write' 'Diagnostic'
            $text=BoundControl 'txtActionPathDiagnostic' 'Text' '' 'frmActionPathView'
            $caption=if($run.Incomplete){'Incomplete evidence'}else{'Conclusion observed'}
            $visible=$text.Contains($caption+"`r`nSaved diagnostic result: "+$result.EvaluationId)
            if($run.Success){$visible=$visible -and $text.Contains('Source event: '+$eventId+' - Applied')}
            Pass ($kind+'.VisibleSavedConclusion') $visible
            Pass ($kind+'.Layout') ((BoundControl '' 'Fit' 'Current' 'frmActionPathView') -ceq 'True')
            Deliver 'frmActionPathView' 'txtActionPathDiagnostic' 'ViewportBottom'
            CaptureOwnedFormByCaptionEvidence 'Action Path view' ('seed-guide-'+$kind.ToLowerInvariant()+'-source-result.png')
        }
        Pass 'TrainingPublicationAndInventoryPreserved' ((Retained $training) -and (PinsRetained $activity) -and (Hash $eventsPath) -ceq $publicationHash -and (Hash $InventoryPath) -ceq $inventoryHash)
    }finally{CloseRecordingViewer;SelectTarget $Fixture 'config-admin'}
}
