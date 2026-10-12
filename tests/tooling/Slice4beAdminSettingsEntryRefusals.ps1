# Fault injection affects disposable collaborators only; the real owner publishes outcomes.
function Install-AdminSettingsEntryRefusalProbe($TestModule,$FormCode) {
    $TestModule.CodeModule.AddFromString(@'
Private mEntryFailureMode As String
Private mEntryOpeningFailures As Long
Public Sub EntryFailureModeForTest(ByVal value As String)
    mEntryFailureMode = value
End Sub
Public Function EntryConfigFaultForTest() As Boolean
    EntryConfigFaultForTest = (mEntryFailureMode = "Config")
    If EntryConfigFaultForTest Then mEntryFailureMode = ""
End Function
Public Function EntryOpeningFaultForTest() As Boolean
    EntryOpeningFaultForTest = (mEntryFailureMode = "Opening")
    If EntryOpeningFaultForTest Then
        mEntryFailureMode = "": mEntryOpeningFailures = mEntryOpeningFailures + 1
    End If
End Function
Public Function EntryOpeningFaultCountForTest() As Long
    EntryOpeningFaultCountForTest = mEntryOpeningFailures
End Function
'@)
    $config=$packages['invSys.Core.xlam'].VBProject.VBComponents.Item('modConfig').CodeModule
    $config.InsertLines($config.CountOfDeclarationLines+1,"Private mEntryRefuseNext As Boolean`r`nPrivate mEntryRefusalCount As Long")
    $config.AddFromString(@'
Public Sub EntryArmReadRefusalForTest()
    mIsLoaded = False
    mEntryRefuseNext = True
End Sub
Public Sub EntryClearReadRefusalForTest()
    mEntryRefuseNext = False
End Sub
Public Function EntryReadRefusalCountForTest() As Long
    EntryReadRefusalCountForTest = mEntryRefusalCount
End Function
'@)
    $start=$config.ProcStartLine('LoadConfig',0);$end=$start+$config.ProcCountLines('LoadConfig',0)
    $hits=@(for($line=$start;$line -lt $end;$line++){if($config.Lines($line,1).Trim() -eq 'InitializeState'){$line}})
    if($hits.Count -ne 1){throw 'Unique configuration initialization boundary required.'}
    $config.InsertLines($hits[0]+1,@'
    If mEntryRefuseNext Then
        mEntryRefuseNext = False: mEntryRefusalCount = mEntryRefusalCount + 1
        AddValidationIssue "ERROR", "ENTRY_TEST_LOAD", "Controlled configuration read refusal."
        Exit Function
    End If
'@)
    $owner=$packages['invSys.Admin.xlam'].VBProject.VBComponents.Item('cAdminSettingsEntry').CodeModule
    $anchor='Call modRoleWorkbookSurfaces.EnsureAdminLegacyWorkbookSurface(targetWb, report)'
    $hits=@(for($line=1;$line -le $owner.CountOfLines;$line++){if($owner.Lines($line,1).Trim() -eq $anchor){$line}})
    if($hits.Count -ne 1){throw 'Unique launcher configuration boundary required.'}
    $owner.InsertLines($hits[0]+1,'    If TestD5Commands.EntryConfigFaultForTest() Then modConfig.EntryArmReadRefusalForTest')
    $start=$FormCode.ProcBodyLine('UserForm_Initialize',0)
    $FormCode.InsertLines($start+1,'    If TestD5Commands.EntryOpeningFaultForTest() Then Err.Raise vbObjectError + 8121, , "CONTROLLED_ENTRY_INITIALIZATION_FAILURE"')
}

function Start-AdminEntryRefusalNotice([string]$Expected) {
    $id=[guid]::NewGuid().ToString('N');$ready=Join-Path $runRoot ('entry-refusal-'+$id+'.ready');$stop=Join-Path $runRoot ('entry-refusal-'+$id+'.stop')
    $job=Start-Job -ArgumentList $repo,([long]$excel.Hwnd),$ready,$stop,$Expected -ScriptBlock {
        param($repo,$handle,$ready,$stop,$expected)
        $ErrorActionPreference='Stop'
        . (Join-Path $repo 'tools/plan022-dialog-observer.ps1')
        Invoke-Plan022NativeDialogObservation -ProcessId 0 -TimeoutSeconds 0
        [uint32]$owner=0;[void][Plan022NativeDialogs]::GetWindowThreadProcessId([IntPtr]$handle,[ref]$owner)
        if(-not $owner){throw 'Owned refusal process unavailable.'}
        $created=(Get-Process -Id $owner).StartTime.ToUniversalTime().Ticks
        [IO.File]::WriteAllText($ready,'Ready');$seen=$false;$until=[DateTime]::UtcNow.AddSeconds(60)
        while(-not(Test-Path -LiteralPath $stop) -and [DateTime]::UtcNow -lt $until){
            if((Get-Process -Id $owner).StartTime.ToUniversalTime().Ticks -ne $created){throw 'Refusal process identity changed.'}
            $texts=@([Plan022NativeDialogs]::Poll($owner))
            if(@($texts|Where-Object{$_.Contains('ControlType.Text|'+$expected)}).Count){$seen=$true}
            Start-Sleep -Milliseconds 100
        }
        [pscustomobject]@{NoticeVisible=$seen}
    }
    $until=[DateTime]::UtcNow.AddSeconds(12)
    while(-not(Test-Path -LiteralPath $ready) -and $job.State -eq 'Running' -and [DateTime]::UtcNow -lt $until){Start-Sleep -Milliseconds 100}
    if(-not(Test-Path -LiteralPath $ready)){if($job.State -eq 'Running'){Stop-Job $job};Remove-Job $job;throw 'Refusal observer unavailable; not behavioral RED.'}
    [pscustomobject]@{Job=$job;Stop=$stop}
}

function Test-AdminSettingsEntryRefusals($Fixture,$Other) {
    . (Join-Path $PSScriptRoot 'Slice4beActivityAssertions.ps1')
    function Act([string]$name,[object[]]$values=@()){Run 'invSys.Admin.xlam' ('TestD5Commands.'+$name) $values}
    function Files {@(Get-Slice4beActivityFiles $Fixture)}
    $book=$null
    try {
        SelectTarget $Fixture
        if(-not [bool](Run 'invSys.Core.xlam' 'TestAdminEntry.EnablePolicy')){throw 'Refusal recording policy fixture unavailable.'}
        $pins=@{};foreach($file in Files){$pins[$file]=(Get-FileHash $file).Hash}
        if($pins.Count -lt 10){throw 'Nonempty boundary history required.'}
        $authority=@{};foreach($file in Get-ChildItem $Fixture.Root -File -Filter '*.xlsb'){$authority[$file.FullName]=(Get-FileHash $file.FullName).Hash}
        $otherPins=@{};foreach($file in Get-ChildItem $Other.Root -Recurse -File){$otherPins[$file.FullName]=(Get-FileHash $file.FullName).Hash}
        $book=$excel.Workbooks.Add();$book.Worksheets.Item(1).Cells.Item(1,1).Value2='Refusal operator canary'
        foreach($case in @(@('SignedOut','Current invSys user is not signed in.',''),@('Config','Canonical config could not be loaded.','REJECTED'),@('Opening','Settings could not be opened. Inspect the current window before retrying.','FAILED'))){
            SelectTarget $Fixture
            $mode=$case[0];$label='AdminEntryRefusal.'+$mode
            if([int](Act 'RemainingSettingsInstances') -ne 0){throw 'Refusal fixture contains a pre-existing form.'}
            if($mode -ceq 'SignedOut'){[void](Run 'invSys.Core.xlam' 'modAuth.SignOut')}
            $context=[string](Run 'invSys.Core.xlam' 'modActivity.CaptureContext')
            if(($mode -ceq 'SignedOut') -ne ($context -ceq '')){throw 'Refusal authentication fixture unavailable.'}
            $before=@(Files);$visible=[int](Act 'EntryBoundaryCount')
            $configFaults=[int](Run 'invSys.Core.xlam' 'modConfig.EntryReadRefusalCountForTest');$openingFaults=[int](Act 'EntryOpeningFaultCountForTest')
            [void](Act 'EntryFailureModeForTest' @($mode));[void](Act 'EntryBoundaryArm' @($true))
            $observer=Start-AdminEntryRefusalNotice $case[1]
            try {
                [void](Run 'invSys.Admin.xlam' 'modAdmin.Open_Settings')
                [IO.File]::WriteAllText($observer.Stop,'Stop');[void](Wait-Job $observer.Job -Timeout 5)
                if($observer.Job.State -ne 'Completed'){throw 'Refusal observer did not finish.'}
                $notice=Receive-Job $observer.Job -ErrorAction Stop
            }finally{[IO.File]::WriteAllText($observer.Stop,'Stop');if($observer.Job.State -eq 'Running'){Stop-Job $observer.Job};Remove-Job $observer.Job;[void](Act 'EntryBoundaryArm' @($false));[void](Act 'EntryFailureModeForTest' @(''))}
            if($mode -ceq 'Config' -and [int](Run 'invSys.Core.xlam' 'modConfig.EntryReadRefusalCountForTest') -ne ($configFaults+1)){throw 'One actual configuration-read refusal required; not behavioral RED.'}
            if($mode -ceq 'Opening' -and [int](Act 'EntryOpeningFaultCountForTest') -ne ($openingFaults+1)){throw 'One actual Initialize fault required; not behavioral RED.'}
            Check ($label+'.RequiredFailureReached') $true
            Check ($label+'.VisibleOwnerNotice') ([bool]$notice.NoticeVisible)
            Check ($label+'.NoSettingsPresentation') ([int](Act 'EntryBoundaryCount') -eq $visible)
            Check ($label+'.ContextNotReplaced') ([string](Run 'invSys.Core.xlam' 'modActivity.CaptureContext') -ceq $context)
            $rows=@(Files|Where-Object{$_ -cnotin $before}|ForEach-Object{Get-Content $_ -Raw|ConvertFrom-Json})
            if($mode -ceq 'SignedOut'){Check ($label+'.NoUnauthenticatedActivity') ($rows.Count -eq 0)}
            else {
                $first=@($rows|Where-Object OutcomeCode -CEQ 'REQUESTED');$last=@($rows|Where-Object OutcomeCode -CEQ $case[2]);$valid=$rows.Count -eq 2 -and $first.Count -eq 1 -and $last.Count -eq 1
                if($valid){$valid=$first[0].ActivityId -ceq $last[0].ActivityId -and $first[0].RecordId -cne $last[0].RecordId}
                foreach($row in $rows){
                    $terminal=$row.OutcomeCode -ceq $case[2];$severity=if(-not $terminal){'Info'}elseif($mode -ceq 'Config'){'Warning'}else{'Error'}
                    $effect=if($terminal -and $mode -ceq 'Config'){'Unchanged'}else{'Unknown'}
                    $valid=$valid -and $row.ControlId -ceq 'ADMIN_SETTINGS_OPEN' -and $row.OwnerId -ceq 'ADMIN_SETTINGS_UI' -and $row.WarehouseId -ceq $Fixture.Warehouse -and $row.UserId -ceq 'config-admin' -and $row.CatalogVersion -eq 33 -and $row.Severity -ceq $severity -and $row.DataEffect -ceq $effect -and $row.EventCode -ceq ('ADMIN_SETTINGS_OPEN_'+$row.OutcomeCode) -and @($row.SourceEventRefs).Count -eq 0
                    if($terminal){$message=if($mode -ceq 'Config'){'Settings could not be opened in the current context.'}else{'Settings navigation failed; inspect the form before retrying.'};$valid=$valid -and $row.UserMessage -ceq $message}
                    $wire=ConvertTo-Json $row -Depth 20 -Compress
                    $valid=$valid -and -not [bool](Run 'invSys.Core.xlam' 'TestAdminEntry.Concludes' @($wire))
                    foreach($value in @($Fixture.Secret,(CredentialHash $Fixture.Secret),$Fixture.Root,'ENTRY_TEST_LOAD','CONTROLLED_ENTRY_INITIALIZATION_FAILURE','Controlled configuration read refusal.')){$valid=$valid -and -not $wire.Contains($value)}
                }
                Check ($label+'.ExactNonconcludingRedactedPair') $valid
            }
            [void](Act 'CleanupDefaultSettings')
            Check ($label+'.NoHiddenInstanceAfterCleanup') ([int](Act 'RemainingSettingsInstances') -eq 0)
        }
        Check 'AdminEntryRefusal.AuthorityBytesPreserved' (@($authority.Keys|Where-Object{(Get-FileHash $_).Hash -cne $authority[$_]}).Count -eq 0)
        Check 'AdminEntryRefusal.OlderRecordsImmutable' (@($pins.Keys|Where-Object{(Get-FileHash $_).Hash -cne $pins[$_]}).Count -eq 0)
        Check 'AdminEntryRefusal.OtherWarehousePreserved' (@($otherPins.Keys|Where-Object{(Get-FileHash $_).Hash -cne $otherPins[$_]}).Count -eq 0)
        Check 'AdminEntryRefusal.OperatorCanaryPreserved' ([string]$book.Worksheets.Item(1).Cells.Item(1,1).Value2 -ceq 'Refusal operator canary')
    }finally{[void](Act 'EntryFailureModeForTest' @(''));[void](Act 'EntryBoundaryArm' @($false));[void](Run 'invSys.Core.xlam' 'modConfig.EntryClearReadRefusalForTest');[void](Act 'CleanupDefaultSettings');if($null -ne $book){$book.Close($false)};SelectTarget $Fixture}
}
