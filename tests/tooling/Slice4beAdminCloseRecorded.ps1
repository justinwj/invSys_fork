# Catalog33 regression: real button/native dismissal, independent recordings and authored intent.
function Test-AdminCloseRecorded($Fixture,$Other) {
    . (Join-Path $PSScriptRoot 'Slice4beRecordingFixture.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beRecordingEvaluation.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beEvaluationContracts.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beGuideTestActions.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beAdminCloseActivity.ps1')
    function Hash([string]$Path){(Get-FileHash -LiteralPath $Path).Hash}
    function Retained($Before,$After){foreach($path in $Before.Keys){if(-not $After.ContainsKey($path) -or $After[$path] -cne $Before[$path]){return $false}};return $true}
    function CloseSettings {[void](Run 'invSys.Admin.xlam' 'TestD5Commands.CloseSettings')}
    function Instances {[int](Run 'invSys.Admin.xlam' 'TestD5Commands.GeneralSettingsInstancesForTest')}
    $owned=[IO.Path]::GetFullPath($runRoot).TrimEnd('\')+'\'
    $root=[IO.Path]::GetFullPath($Fixture.Root).TrimEnd('\')+'\'
    $authPath=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Auth.xlsb')
    if(-not $root.StartsWith($owned,[StringComparison]::OrdinalIgnoreCase)){throw 'Generated recording fixture required.'}
    foreach($path in @($Fixture.Config,$authPath,$activityRoot,$journalRoot)){
        if(-not [IO.Path]::GetFullPath($path).StartsWith($root,[StringComparison]::OrdinalIgnoreCase)){throw 'Recording path escaped generated fixture.'}
    }
    $configBytes=[IO.File]::ReadAllBytes($Fixture.Config);$configHash=Hash $Fixture.Config
    $authHash=Hash $authPath;$otherPins=BoundPins $Other.Root
    $hostPath=Join-Path $runRoot 'admin-close-recorded-host.xlsb'
    if(Test-Path -LiteralPath $hostPath){throw 'Preserve existing recording host.'}
    $anchor=$null;$runs=@()
    $steps=@(@('Close button','ADMIN_SETTINGS_CLOSE','CLOSED'),@('Window close','ADMIN_SETTINGS_CLOSE','CLOSED'))
    try {
        $anchor=$excel.Workbooks.Add();$anchor.SaveAs($hostPath,50);$anchor.Close($false);$anchor=$null
        $hostHash=Hash $hostPath;$anchor=$excel.Workbooks.Open($hostPath,0,$false)
        SelectTarget $Fixture 'config-admin'
        if(-not [bool](Run 'invSys.Core.xlam' 'TestSettingsOwnerState.SetupRecordingPolicy')){throw 'Recording policy prerequisite unavailable.'}
        $enabledHash=Hash $Fixture.Config
        foreach($kind in @('Source','Observed','Interrupted')){
            $label='AdminCloseRecorded.'+$kind
            CloseSettings;CloseRecordingViewer;SelectTarget $Fixture 'config-admin';OpenRecordingViewer
            $prior=BoundPins $journalRoot;$activity=ActivityPins
            Check ($label+'.ActualStart') ((RecordingControl 'Start Recording' 'Click') -ceq 'DELIVERED')
            $starts=@(Get-ChildItem -LiteralPath $journalRoot -Filter '*.json' -File|Where-Object {-not $prior.ContainsKey($_.FullName)}|ForEach-Object {Get-Content $_.FullName -Raw|ConvertFrom-Json}|Where-Object RecordType -CEQ 'Start')
            if($starts.Count -ne 1){throw 'One durable Start required.'}
            $sequence=$starts[0].SequenceId;$ordinal=0;$all=@()
            foreach($step in $steps){
                $ordinal++;CloseSettings
                [void](Run 'invSys.Admin.xlam' 'TestD5Commands.OpenSettings')
                [void](Run 'invSys.Admin.xlam' 'TestD5Commands.ShowSettings')
                if((Instances) -ne 1){throw 'One visible Settings instance required.'}
                $before=@(Get-Slice4beActivityFiles $Fixture);$pins=ActivityPins
                if($ordinal -eq 1){$delivered=[string](Run 'invSys.Admin.xlam' 'TestD5Commands.GeneralSettingsAction' @('Close','')) -ceq 'DELIVERED'}
                else{$delivered=Close-AdminNativeForTest}
                Check ($label+'.'+$ordinal+'.ActualDismissalAndDisposal') ($delivered -and (Instances) -eq 0 -and [InvSysSettingsCapture]::OwnedVisibleForm('invSys Settings',[IntPtr]$excel.Hwnd) -eq [IntPtr]::Zero)
                $records=@(Get-Slice4beActivityFiles $Fixture|Where-Object {$_ -cnotin $before}|ForEach-Object {Get-Content $_ -Raw|ConvertFrom-Json})
                $first=@($records|Where-Object OutcomeCode -CEQ 'REQUESTED');$last=@($records|Where-Object OutcomeCode -CEQ 'CLOSED')
                $valid=$records.Count -eq 2 -and $first.Count -eq 1 -and $last.Count -eq 1
                if($valid){$valid=$first[0].ActivityId -ceq $last[0].ActivityId -and $first[0].RecordId -cne $last[0].RecordId}
                foreach($record in $records){$valid=$valid -and $record.ControlId -ceq 'ADMIN_SETTINGS_CLOSE' -and $record.OwnerId -ceq 'ADMIN_SETTINGS_UI' -and $record.SequenceId -ceq $sequence -and $record.Ordinal -eq $ordinal -and $record.CatalogVersion -eq 33 -and $record.UserId -ceq 'config-admin' -and $record.WarehouseId -ceq $Fixture.Warehouse -and @($record.SourceEventRefs).Count -eq 0}
                Check ($label+'.'+$ordinal+'.ExactRecordedPairAndContext') $valid
                $safe=$true
                foreach($record in $records){$raw=ConvertTo-Json $record -Depth 20 -Compress;foreach($value in @($Fixture.Secret,(CredentialHash $Fixture.Secret),$Fixture.Root,'PinHash')){$safe=$safe -and -not $raw.Contains($value)}}
                Check ($label+'.'+$ordinal+'.NoInputOrCredentialPayload') $safe
                Check ($label+'.'+$ordinal+'.PriorEvidenceAndAuthorityPreserved') ((PinsRetained $pins) -and (Hash $Fixture.Config) -ceq $enabledHash -and (Hash $authPath) -ceq $authHash)
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
        CloseSettings;CloseRecordingViewer;SelectTarget $Fixture 'config-admin'
        if(-not [bool](Run 'invSys.Admin.xlam' 'modAdminConsole.PublishReadFixtureForTest')){throw 'Actual publication prerequisite failed.'}
        $training=BoundPins $journalRoot
        . (Join-Path $PSScriptRoot 'Slice4beGeneralSettingsGuidePaths.ps1')
        Test-GeneralSettingsGuidePaths $Fixture $runs $steps -CloseActions
        Check 'AdminCloseRecorded.TrainingEvidenceImmutableThroughEvaluation' (Retained $training (BoundPins $journalRoot))
        Check 'AdminCloseRecorded.OtherWarehousePreserved' (BoundSame $otherPins (BoundPins $Other.Root))
    }finally{
        CloseSettings;CloseRecordingViewer
        if($null -ne $anchor){$anchor.Close($false)}
        [IO.File]::WriteAllBytes($Fixture.Config,$configBytes)
        SelectTarget $Fixture 'config-admin'
    }
    Check 'AdminCloseRecorded.FinalAuthorityAndHostPreserved' ((Hash $Fixture.Config) -ceq $configHash -and (Hash $authPath) -ceq $authHash -and (Hash $hostPath) -ceq $hostHash)
}
