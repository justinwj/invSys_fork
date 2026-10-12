# Authored intent compared with independent recordings through packaged controls.
function Test-GeneralSettingsGuidePaths($Fixture,$Runs,$Steps,[switch]$CloseActions,[switch]$WorksheetScale,[switch]$SettingsNavigation,[switch]$Reconcile) {
    if(@($CloseActions,$WorksheetScale,$SettingsNavigation,$Reconcile|Where-Object{$_}).Count -gt 1){throw 'Choose one independent guide fixture.'}
    function Delivered([string]$Value){if($Value -cne 'DELIVERED'){throw 'Actual General guide control unavailable; not product RED.'}}
    function Author([string]$Name,[string]$Action,[string]$Value=''){BoundControl $Name $Action $Value 'frmActionPathGuide'}
    function View([string]$Name,[string]$Action,[string]$Value=''){BoundControl $Name $Action $Value 'frmActionPathView'}
    function Retained($Before,$After){foreach($path in $Before.Keys){if(-not $After.ContainsKey($path) -or $After[$path] -cne $Before[$path]){return $false}};return $true}
    $prefix=if($Reconcile){'ReconcileGuide'}elseif($SettingsNavigation){'SettingsNavigationGuide'}elseif($WorksheetScale){'WorksheetScaleGuide'}elseif($CloseActions){'AdminCloseGuide'}else{'GeneralGuide'}
    $capturePrefix=if($Reconcile){'reconcile-guide-'}elseif($SettingsNavigation){'settings-navigation-guide-'}elseif($WorksheetScale){'worksheet-scale-guide-'}elseif($CloseActions){'admin-close-guide-'}else{'general-guide-'}
    function Pass([string]$Name,[bool]$Value){Check ($prefix+'.'+$Name) $Value|Out-Host}
    $source=@($Runs|Where-Object Kind -CEQ 'Source')[0]
    $observed=@($Runs|Where-Object Kind -CEQ 'Observed')[0]
    $eventsPath=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Snapshot.Events.json')
    $publication=Get-Content $eventsPath -Raw|ConvertFrom-Json;$publicationHash=Hash $eventsPath
    $originalPins=BoundPins $journalRoot;$activityPins=ActivityPins
    $instruction="Use General Settings to inspect configuration and UOM choices, manage carriers, and save the connection option.`r`nCarrier and connection changes apply to this Windows user; they do not prove inventory changes.`r`nRead the observed result. Interrupted or denied actions cannot borrow completion from this guide."
    if($CloseActions){$instruction="Dismiss Settings with Close, reopen it, then use the window close button.`r`nUnsaved form edits are discarded. Dismissal proves neither a settings save nor Domain application.`r`nBoth gestures record the same Close control; the recording does not identify the gesture. Interrupted actions cannot borrow completion from this guide."}
    if($WorksheetScale){$instruction="With Designs disabled and the local worksheet recipe available, select the recipe and enter 200 percent. Click Apply Scale and inspect the prepared rows.`r`nClick Apply Scale again. The recipe quantity remains four, not eight; existing System_Key and custom fields are preserved.`r`nThis concludes two local staging commands, not inventory application. An interrupted recording cannot borrow completion from this guide."}
    if($SettingsNavigation){$instruction="Open Settings, select Event Tracking, Event Detail, Action Paths, Tracking, and General, then Close.`r`nKeep Settings open throughout the tab changes. These observations prove presentation and dismissal, not a settings save or Domain application.`r`nAn interrupted recording cannot borrow completion from this guide."}
    if($Reconcile){$instruction="Open Shipment Reconcile in the training warehouse and inspect the selected exact inventory entity's shipment.`r`nRead the current quantity, balance after shipment and proposed correction. Supply a repair narrative and confirm only the intended correction.`r`nA queued correction proves command completion, not inventory application. An interrupted recording cannot borrow completion from this guide."}
    try {
        Pass 'IndependentSourceAndObservedRun' ($source.Journal.ActionPathId -cne $observed.Journal.ActionPathId -and @($observed.Records|Where-Object {$_.ActivityId -cin @($source.Records.ActivityId)}).Count -eq 0)
        CloseRecordingViewer;SelectTarget $Fixture 'config-admin';OpenRecordingViewer
        Delivered (BoundLibrary 'Open')
        if((BoundLibrary 'Select' $source.Journal.ActionPathId) -cne 'SELECTED'){throw 'Exact General guide source unavailable.'}
        Delivered (BoundControl 'btnCreateGuide' 'Click' '' 'frmActionPaths')
        Delivered (Author 'txtGuideName' 'Write' $(if($Reconcile){'Shipment Reconcile training'}elseif($SettingsNavigation){'Settings navigation training'}elseif($WorksheetScale){'Worksheet Scale training'}elseif($CloseActions){'Settings Close training'}else{'General Settings training'}))
        Delivered (Author 'txtGuideInstructions' 'Write' $instruction)
        Pass $(if($SettingsNavigation){'SevenObservedGuideSteps'}elseif($CloseActions -or $WorksheetScale -or $Reconcile){'TwoObservedGuideSteps'}else{'NineObservedGuideSteps'}) ((Author 'lstGuideSteps' 'Rows') -ceq [string]$Steps.Count)
        for($index=0;$index -lt $Steps.Count;$index++){
            if((Author 'lstGuideSteps' 'Select' ([string]$index)) -cne 'SELECTED'){throw 'General guide step unavailable.'}
            Delivered (Author 'txtGuideStepInstruction' 'Write' ('Use '+$Steps[$index][0]+'. Read the owner result before continuing.'))
        }
        Delivered (Author 'btnGuideExpectedConclusion' 'Click')
        $count=BoundControl 'lstExpectedSteps' 'Rows' '' 'frmActionPathExpectation'
        for($index=[int]$count-1;$index -ge 0;$index--){
            if((BoundControl 'lstExpectedSteps' 'Select' ([string]$index) 'frmActionPathExpectation') -cne 'SELECTED'){throw 'Expectation removal selection unavailable.'}
            Delivered (BoundControl 'btnRemoveExpectedStep' 'Click' '' 'frmActionPathExpectation')
        }
        foreach($step in $Steps){
            Delivered (BoundControl 'cboExpectedControl' 'Write' $step[1] 'frmActionPathExpectation')
            Delivered (BoundControl 'cboExpectedOutcome' 'Write' $step[2] 'frmActionPathExpectation')
            Delivered (BoundControl 'btnAddExpectedStep' 'Click' '' 'frmActionPathExpectation')
        }
        if((BoundControl 'cboTerminalStep' 'Select' ([string]($Steps.Count-1)) 'frmActionPathExpectation') -cne 'SELECTED'){throw 'Last expected terminal step unavailable.'}
        Delivered (BoundControl 'cboTerminalKind' 'Write' 'CommandCompleted' 'frmActionPathExpectation')
        Delivered (BoundControl 'btnUseExpectation' 'Click' '' 'frmActionPathExpectation')
        $directory=Join-Path $journalRoot 'Guides';$prior=BoundPins $directory
        Delivered (Author 'btnSaveGuide' 'Click')
        $fresh=@(Get-ChildItem -LiteralPath $directory -Filter '*.json' -File|Where-Object {-not $prior.ContainsKey($_.FullName)})
        if($fresh.Count -ne 1){throw 'Exactly one newly authored General guide required.'}
        $guide=ReadGuideExpectationRecord $fresh[0].FullName
        if($null -eq $guide){throw 'General guide integrity unavailable.'}
        $exact=$guide.SourceRun.RecordId -ceq $source.Journal.RecordId -and $guide.SourceRun.ContentSha256 -ceq $source.Journal.ContentSha256 -and @($guide.Steps).Count -eq $Steps.Count -and @($guide.ExpectedConclusion.Steps).Count -eq $Steps.Count -and $guide.ExpectedConclusion.TerminalKind -ceq 'CommandCompleted'
        for($index=0;$index -lt $Steps.Count;$index++){$exact=$exact -and $guide.ExpectedConclusion.Steps[$index].ControlId -ceq $Steps[$index][1] -and $guide.ExpectedConclusion.Steps[$index].RequiredOutcome -ceq $Steps[$index][2]}
        Pass 'ExactSourceAndAuthoredIntent' $exact
        CloseRecordingViewer;SelectTarget $Fixture 'config-reader';OpenRecordingViewer
        Pass 'ReaderCannotAuthor' (-not [bool](Run 'invSys.Core.xlam' 'modAuth.CanPerform' @('ACTION_PATH_MAINT','config-reader',$Fixture.Warehouse,'S1')))
        $kinds=if($CloseActions -or $WorksheetScale -or $SettingsNavigation -or $Reconcile){@('Observed','Interrupted')}else{@('Observed','Interrupted','Denied')}
        foreach($kind in $kinds){
            $run=@($Runs|Where-Object Kind -CEQ $kind)[0];$journal=$run.Journal
            Delivered (BoundLibrary 'Open')
            if((BoundLibrary 'Select' $journal.ActionPathId) -cne 'SELECTED'){throw 'Exact General comparison recording unavailable.'}
            BoundOpen;BoundSelect $guide
            Delivered (BoundControl 'btnUseGuideForRun' 'Click');Delivered (BoundControl 'btnCloseGuides' 'Click')
            $before=@(EvaluationFiles|ForEach-Object FullName)
            Delivered (BoundControl 'btnEvaluatePath' 'Click' '' 'frmActionPaths')
            $fresh=@(EvaluationFiles|Where-Object {$_.FullName -cnotin $before})
            if($fresh.Count -ne 1){throw 'Paired Evaluate must append exactly one result.'}
            $result=Get-Content $fresh[0].FullName -Raw|ConvertFrom-Json
            $state=if($run.Incomplete){'Incomplete'}elseif($run.Success){'Concluded'}else{'Failed'}
            $reason=if($run.Incomplete){'CAPTURE_INCOMPLETE'}elseif($run.Success){'COMMAND_COMPLETED'}else{'REQUIRED_OUTCOME_MISMATCH'}
            Pass ($kind+'.ExactDiagnosticState') ($result.ResultState -ceq $state -and $reason -cin @($result.ReasonCodes))
            Pass ($kind+'.ExactGuideAndObservedJournal') ($result.JournalRecordId -ceq $journal.RecordId -and $result.JournalSha256 -ceq $journal.ContentSha256 -and $result.Guide.ContentSha256 -ceq $guide.ContentSha256 -and $result.Guide.Version -eq $guide.Version)
            Pass ($kind+'.ExactPublication') ($result.Publication.PublicationId -ceq $publication.PublicationId -and $result.Publication.ContentSha256 -ceq $publication.ContentSha256 -and $result.Publication.WarehouseId -ceq $Fixture.Warehouse)
            Pass ($kind+'.NoDomainApplicationClaim') (@($result.TerminalSources).Count -eq 0 -and $result.ExpectedConclusion.TerminalKind -ceq 'CommandCompleted')
            # Incomplete/failed conclusions retain the actions actually completed.
            # Interruption follows Reload; denial still permits four selections.
            $ordinals=@(if($run.Success){1..$Steps.Count}elseif($run.Incomplete){1}else{2,3,6,8})
            $exactMatches=@($result.Matches).Count -eq $ordinals.Count
            foreach($ordinal in $ordinals){
                $expected=@($run.Records|Where-Object {$_.Ordinal -eq $ordinal -and $_.OutcomeCode -ceq $Steps[$ordinal-1][2]})
                $matches=@($result.Matches|Where-Object Ordinal -EQ $ordinal)
                $exactMatches=$exactMatches -and $expected.Count -eq 1 -and $matches.Count -eq 1
                if($expected.Count -eq 1 -and $matches.Count -eq 1){$exactMatches=$exactMatches -and $matches[0].ActivityId -ceq $expected[0].ActivityId -and $matches[0].StepId -ceq $guide.ExpectedConclusion.Steps[$ordinal-1].StepId}
            }
            Pass ($kind+'.OnlyObservedRunCanMatch') $exactMatches
            # Reconcile loses CONFIRMED after REQUESTED was saved; the attempted
            # terminal step has an unmatched outcome. Capture remains Incomplete.
            $missing=if($run.Incomplete -and -not $Reconcile){$Steps.Count-1}else{0}
            $failed=if($Reconcile -and $run.Incomplete){1}elseif(-not $run.Success -and -not $run.Incomplete){5}else{0}
            $unmatched=@($result.MissingSteps).Count -eq $missing -and @($result.FailedSteps).Count -eq $failed -and @($result.UnavailableSteps).Count -eq 0
            if($Reconcile -and $run.Incomplete){$unmatched=$unmatched -and $result.FailedSteps[0] -ceq $guide.ExpectedConclusion.TerminalStepId -and 'REQUIRED_OUTCOME_MISMATCH' -cin @($result.ReasonCodes)}
            Pass ($kind+'.ExactUnmatchedSteps') $unmatched
            $pins=BoundPins $journalRoot
            Delivered (BoundControl 'btnViewActionPath' 'Click' '' 'frmActionPaths')
            $pair=View 'lblActionPathPair' 'Label';$how=View 'txtActionPathHowTo' 'Text';$diagnostic=View 'txtActionPathDiagnostic' 'Text'
            Pass ($kind+'.ExactPairAndMultilineInstructions') ($pair.Contains($guide.ContentSha256) -and $pair.Contains($journal.ActionPathId) -and $how.Contains($instruction))
            $noBorrowing=$diagnostic.Contains($result.EvaluationId) -and $diagnostic.Contains($run.Records[0].ActivityId)
            foreach($id in @($source.Records.ActivityId|Select-Object -Unique)){$noBorrowing=$noBorrowing -and -not $diagnostic.Contains($id)}
            Pass ($kind+'.DiagnosticCannotBorrowGuideSourceSuccess') $noBorrowing
            $caption=if($run.Incomplete){'Incomplete evidence'}elseif($run.Success){'Conclusion observed. Command completed; Domain application not asserted'}else{'Failed'}
            Pass ($kind+'.DisplayedSavedConclusion') ($diagnostic.Contains($caption+"`r`nSaved diagnostic result: "+$result.EvaluationId))
            foreach($method in @('How-To','Diagnostic','Compare both')){
                Delivered (View 'cboActionPathView' 'Write' $method)
                Pass ($kind+'.View.'+$method) ((View 'cboActionPathView' 'Selected') -ceq $method -and ((View 'txtActionPathHowTo' 'State') -split '\|')[0] -ceq $(if($method -ceq 'Diagnostic'){'False'}else{'True'}) -and ((View 'txtActionPathDiagnostic' 'State') -split '\|')[0] -ceq $(if($method -ceq 'How-To'){'False'}else{'True'}))
                Pass ($kind+'.SameEvidence.'+$method) ((View 'lblActionPathPair' 'Label') -ceq $pair -and (View 'txtActionPathHowTo' 'Text') -ceq $how -and (View 'txtActionPathDiagnostic' 'Text') -ceq $diagnostic)
                Pass ($kind+'.Layout.'+$method) ((BoundControl '' 'Fit' 'Current' 'frmActionPathView') -ceq 'True')
                if($kind -ceq 'Observed' -or $method -ceq 'Compare both'){CaptureOwnedFormByCaptionEvidence 'Action Path view' ($capturePrefix+$kind.ToLowerInvariant()+'-'+$method.Replace(' ','-').ToLowerInvariant()+'.png')}
            }
            Delivered (View 'txtActionPathDiagnostic' 'ViewportText' ($caption+"`r`nSaved diagnostic result: "+$result.EvaluationId))
            CaptureOwnedFormByCaptionEvidence 'Action Path view' ($capturePrefix+$kind.ToLowerInvariant()+'-conclusion.png')
            Pass ($kind+'.ViewsPreserveTrainingEvidence') (BoundSame $pins (BoundPins $journalRoot))
            Delivered (View 'btnCloseActionPathView' 'Click')
        }
        Pass 'SourceJournalsAndActivityPreserved' ((Retained $originalPins (BoundPins $journalRoot)) -and (PinsRetained $activityPins) -and (Hash $eventsPath) -ceq $publicationHash)
    }finally{CloseRecordingViewer;SelectTarget $Fixture 'config-admin'}
}
