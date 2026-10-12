# Exact recorded correction -> owning Domain application -> published guide conclusion.
function Install-ReconcileAppliedGuideProbe {
    $packages['invSys.Admin.xlam'].VBProject.VBComponents.Item('TestReconcileIdentity').CodeModule.AddFromString(@'
Public Function ApplyRecordedCorrection(ByVal inboxName As String, ByVal eventId As String) As String
    Dim lo As ListObject, evt As Object, field As Variant, r As Long, selected As Long, count As Long
    Dim logCount As Long, appliedCount As Long, applied As Boolean, duplicate As Boolean
    Dim status As String, code As String, message As String, quantity As Boolean, attributes As Boolean
    If Not NativeFixturePreserved(4) Or Not CurrentHistoryAttributesPreserved() Then
        ApplyRecordedCorrection = "PRECONDITION": Exit Function
    End If
    Set lo = FindTable(Application.Workbooks(inboxName), "tblInboxShip")
    For r = 1 To lo.ListRows.Count
        If StrComp(CStr(Value(lo, r, "EventID")), eventId, vbBinaryCompare) = 0 Then selected = r: count = count + 1
    Next r
    If count <> 1 Then ApplyRecordedCorrection = "PRECONDITION": Exit Function
    Set evt = CreateObject("Scripting.Dictionary")
    For Each field In Array("EventID", "EventType", "WarehouseId", "StationId", "UserId", "CreatedAtUTC", "PayloadJson")
        evt(CStr(field)) = lo.DataBodyRange.Cells(selected, lo.ListColumns(CStr(field)).Index).Value
    Next field
    logCount = FindTable(mInventory, "tblInventoryLog").ListRows.Count
    appliedCount = FindTable(mInventory, "tblAppliedEvents").ListRows.Count
    applied = modInventoryDomainBridge.ApplyInventoryEventBridge(evt, mInventory, "", status, code, message)
    applied = applied And status = "APPLIED"
    status = "": code = "": message = ""
    duplicate = modInventoryDomainBridge.ApplyInventoryEventBridge(evt, mInventory, "", status, code, message)
    duplicate = duplicate And status = "SKIP_DUP"
    quantity = (EntityQty(mKeyA) = 1 And EntityQty(mKeyB) = 55)
    attributes = CurrentHistoryAttributesPreserved() And CanaryPreserved()
    ' The generated-fixture owner persists after both calls, including protection metadata.
    If applied And duplicate Then mInventory.Save
    ApplyRecordedCorrection = CStr(applied) & "|" & CStr(quantity) & "|" & CStr(attributes) & "|" & _
        CStr(FindTable(mInventory, "tblInventoryLog").ListRows.Count = logCount + 1 And _
             FindTable(mInventory, "tblAppliedEvents").ListRows.Count = appliedCount + 1) & "|" & _
        CStr(duplicate) & "|" & CStr(mInventory.Saved)
End Function
'@)
}

function Test-ReconcileAppliedGuide($Fixture,$Runs,$Events,$State,$HostBook) {
    . (Join-Path $PSScriptRoot 'Slice4beEvaluationEvidence.ps1')
    function Deliver([string]$Form,[string]$Control,[string]$Action,[string]$Value=''){
        if((BoundControl $Control $Action $Value $Form) -cne 'DELIVERED'){throw ('Applied guide control unavailable: '+$Control+'; not product RED.')}
    }
    function Retained($Before){foreach($path in $Before.Keys){if(-not(Test-Path -LiteralPath $path) -or (Hash $path) -cne $Before[$path]){return $false}};return $true}
    function Pass([string]$Name,[bool]$Value){Check ('ReconcileAppliedGuide.'+$Name) $Value|Out-Host}
    $observed=@($Runs|Where-Object Kind -CEQ 'Observed')[0];$source=@($Runs|Where-Object Kind -CEQ 'Source')[0]
    $terminal=@($observed.Records|Where-Object OutcomeCode -CEQ 'CONFIRMED')[0]
    $eventId=[string]$terminal.SourceEventRefs[0].EventId
    $guideRoot=Join-Path $journalRoot 'Guides';$eventsPath=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Snapshot.Events.json')
    $originals=@(Get-ChildItem -LiteralPath $guideRoot -File -Filter '*.json'|ForEach-Object {ReadGuideExpectationRecord $_.FullName}|Where-Object {$_.SourceRun.RecordId -ceq $source.Journal.RecordId})
    if($originals.Count -ne 1){throw 'One original authored guide required.'}
    $original=$originals[0];$training=BoundPins $journalRoot;$activity=ActivityPins;$inventoryHash=Hash $State.InventoryPath
    $queueCount=$Events.ListRows.Count
    try{
        CloseRecordingViewer;SelectTarget $Fixture 'config-admin';OpenRecordingViewer
        if((BoundLibrary 'Open') -cne 'DELIVERED'){throw 'Recording library unavailable.'}
        BoundOpen;BoundSelect $original
        Deliver 'frmActionPathLibrary' 'btnEditPublishedGuide' 'Click'
        Deliver 'frmActionPathGuide' 'btnGuideExpectedConclusion' 'Click'
        Deliver 'frmActionPathExpectation' 'cboTerminalKind' 'Write' 'SourceEventsApplied'
        Deliver 'frmActionPathExpectation' 'btnUseExpectation' 'Click'
        Deliver 'frmActionPathGuide' 'txtGuideInstructions' 'Write' "Inspect the exact shipment entity and confirm the intended correction.`r`nWait for the owning processor and publication. Queuing alone cannot conclude this guide."
        Deliver 'frmActionPathGuide' 'btnSaveGuide' 'Click'
        $guide=ReadGuideExpectationRecord (Join-Path $guideRoot ($original.ActionPathId+'.2.json'))
        if($null -eq $guide){throw 'Saved version-two guide prerequisite unavailable.'}
        Pass 'ActualEditAppendsExplicitAppliedExpectation' ($guide.Version -eq 2 -and $guide.PreviousRecordId -ceq $original.RecordId -and $guide.PreviousSha256 -ceq $original.ContentSha256 -and $guide.ExpectedConclusion.TerminalKind -ceq 'SourceEventsApplied' -and $guide.ExpectedConclusion.TerminalStepId -ceq $original.ExpectedConclusion.TerminalStepId)
        Pass 'OriginalGuideAndSourceEvidencePreserved' ((Retained $training) -and (PinsRetained $activity) -and ($guide.Observations|ConvertTo-Json -Depth 20 -Compress) -ceq ($original.Observations|ConvertTo-Json -Depth 20 -Compress) -and ($guide.SourceRun|ConvertTo-Json -Depth 12 -Compress) -ceq ($original.SourceRun|ConvertTo-Json -Depth 12 -Compress))
        Pass 'AuthoringDoesNotApplyCorrection' ((Hash $State.InventoryPath) -ceq $inventoryHash -and [bool](Run 'invSys.Admin.xlam' 'TestReconcileIdentity.NativeFixturePreserved' @(4)))
        CloseRecordingViewer
        $priorPublication='';$pendingFile='';$pendingHash=''
        foreach($stage in @('Pending','Applied')){
            if($stage -ceq 'Applied'){
                SelectTarget $Fixture 'config-admin'
                $values=([string](Run 'invSys.Admin.xlam' 'TestReconcileIdentity.ApplyRecordedCorrection' @($Events.Parent.Parent.Name,$eventId))) -split '\|'
                if($values.Count -ne 6){throw 'Exact queued-event application prerequisite unavailable; not product RED.'}
                foreach($pair in @(@('OwnerApplied',0),@('OnlyExactEntityQuantityChanged',1),@('AttributesAndCustomValuesPreserved',2),@('OneLogAndApplicationRecord',3),@('DuplicateDoesNotReapply',4),@('SavedAuthority',5))){Pass $pair[0] ($values[$pair[1]] -ceq 'True')}
                if(-not [bool](Run 'invSys.Admin.xlam' 'modAdminConsole.PublishReadFixtureForTest')){throw 'Owning publication prerequisite unavailable.'}
                Pass 'EarlierAwaitingResultImmutable' ((Hash $pendingFile) -ceq $pendingHash -and (Get-Content $pendingFile -Raw|ConvertFrom-Json).ResultState -ceq 'Awaiting')
            }
            $publication=Get-Content $eventsPath -Raw|ConvertFrom-Json
            $groups=@($publication.Groups|Where-Object {$_.Source -ceq 'Inventory' -and $_.SourceId -ceq $eventId})
            Pass ($stage+'.ExactApplicationPublication') $(if($stage -ceq 'Pending'){$groups.Count -eq 0}else{$groups.Count -eq 1 -and $publication.PublicationId -cne $priorPublication})
            SelectTarget $Fixture 'config-reader';OpenRecordingViewer
            if((BoundLibrary 'Open') -cne 'DELIVERED' -or (BoundLibrary 'Select' $observed.Journal.ActionPathId) -cne 'SELECTED'){throw 'Exact observed recording unavailable.'}
            BoundOpen;BoundSelect $guide
            Deliver 'frmActionPathLibrary' 'btnUseGuideForRun' 'Click';Deliver 'frmActionPathLibrary' 'btnCloseGuides' 'Click'
            $before=@(EvaluationFiles|ForEach-Object FullName)
            Deliver 'frmActionPaths' 'btnEvaluatePath' 'Click'
            $fresh=@(EvaluationFiles|Where-Object {$_.FullName -cnotin $before})
            if($fresh.Count -ne 1){throw 'Actual Evaluate must append one immutable result.'}
            $result=Get-Content $fresh[0].FullName -Raw|ConvertFrom-Json
            $expected=if($stage -ceq 'Applied'){'Concluded'}else{'Awaiting'}
            $reason=if($stage -ceq 'Applied'){'SOURCE_APPLIED'}else{'SOURCE_PENDING'}
            Pass ($stage+'.OwningEvidenceDeterminesConclusion') ($result.ResultState -ceq $expected -and $reason -cin @($result.ReasonCodes) -and $result.ExpectedConclusion.TerminalKind -ceq 'SourceEventsApplied')
            Pass ($stage+'.ExactGuideAndObservedJournal') ($result.ExpectationSource -ceq 'Guide expectation' -and $result.Guide.ContentSha256 -ceq $guide.ContentSha256 -and $result.Guide.Version -eq 2 -and $result.JournalRecordId -ceq $observed.Journal.RecordId -and $result.JournalSha256 -ceq $observed.Journal.ContentSha256 -and @($result.Matches).Count -eq 2)
            Pass ($stage+'.ExactLoadedPublication') ($result.Publication.PublicationId -ceq $publication.PublicationId -and $result.Publication.ContentSha256 -ceq $publication.ContentSha256 -and $result.Publication.WarehouseId -ceq $Fixture.Warehouse)
            $refs=@($result.TerminalSources);$exact=$refs.Count -eq 1
            if($exact){$ref=$refs[0];$exact=$ref.EventId -ceq $eventId -and $ref.SourceKind -ceq 'Inventory' -and $ref.WarehouseId -ceq $Fixture.Warehouse -and $ref.SubmissionState -ceq 'Submitted'}
            Pass ($stage+'.OnlyExactFreshCorrectionReference') $exact
            $evidence=$exact
            if($evidence -and $stage -ceq 'Applied'){
                $evidence=$groups.Count -eq 1
                if($evidence){
                    $body=[ordered]@{Lines=@($groups[0].Lines)}|ConvertTo-Json -Depth 16 -Compress
                    $evidence=$ref.OwnerStatus -ceq 'Applied' -and $ref.LineCount -eq 1 -and @($ref.SystemKeys).Count -eq 1 -and $ref.SystemKeys[0] -ceq $State.KeyA -and $ref.LinesSha256 -ceq (EvaluationSha $body)
                }
            }elseif($evidence){$evidence=$ref.OwnerStatus -ceq 'Awaiting' -and $ref.LineCount -eq 0 -and @($ref.SystemKeys).Count -eq 0 -and $ref.LinesSha256 -ceq ''}
            Pass ($stage+'.ExactOwnerLinesHashAndIdentity') $evidence
            $pins=BoundPins $journalRoot;$publicationHash=Hash $eventsPath;$authorityHash=Hash $State.InventoryPath
            Deliver 'frmActionPaths' 'btnViewActionPath' 'Click'
            Deliver 'frmActionPathView' 'cboActionPathView' 'Write' 'Diagnostic'
            $text=BoundControl 'txtActionPathDiagnostic' 'Text' '' 'frmActionPathView'
            $caption=if($stage -ceq 'Applied'){'Conclusion observed'}else{'Awaiting published result'}
            $ownerStatus=if($stage -ceq 'Applied'){'Applied'}else{'Awaiting'}
            Pass ($stage+'.VisibleSavedResultAndSource') ($text.Contains($caption+"`r`nSaved diagnostic result: "+$result.EvaluationId) -and $text.Contains('Source event: '+$eventId+' - '+$ownerStatus) -and ($stage -ceq 'Applied' -or -not $text.Contains('Conclusion observed')))
            Pass ($stage+'.Layout') ((BoundControl '' 'Fit' 'Current' 'frmActionPathView') -ceq 'True')
            Deliver 'frmActionPathView' 'txtActionPathDiagnostic' 'ViewportBottom'
            CaptureOwnedFormByCaptionEvidence 'Action Path view' ('reconcile-guide-'+$stage.ToLowerInvariant()+'-source-result.png')
            Pass ($stage+'.ViewPreservesEvidenceAndAuthority') ((BoundSame $pins (BoundPins $journalRoot)) -and (Hash $eventsPath) -ceq $publicationHash -and (Hash $State.InventoryPath) -ceq $authorityHash)
            CloseRecordingViewer
            if($stage -ceq 'Pending'){$pendingFile=$fresh[0].FullName;$pendingHash=Hash $pendingFile;$priorPublication=$publication.PublicationId}
        }
        Pass 'OriginalSubmittedObservationsAndJournalsImmutable' ((Retained $training) -and (PinsRetained $activity))
        Pass 'NoAdditionalQueueAndOperatorCanaryPreserved' ($Events.ListRows.Count -eq $queueCount -and [string]$HostBook.Worksheets.Item(1).Cells.Item(1,1).Value2 -ceq 'Reconcile guide canary')
    }finally{CloseRecordingViewer;SelectTarget $Fixture 'config-admin'}
}
