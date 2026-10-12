# Continuous real modal Settings journey; observations never create owner outcomes.
function Install-AdminNavigationRecordedProbe {
    $admin=$packages['invSys.Admin.xlam'].VBProject
    $form=$admin.VBComponents.Item('frmAdminSettings').CodeModule
    $test=$admin.VBComponents.Item('TestD5Commands').CodeModule
    $test.AddFromString(@'
Private mNavigationReady As String
Private mNavigationInterrupt As Boolean
Public Sub ArmNavigationRecording(ByVal path As String, ByVal interrupt As Boolean)
    mNavigationReady = path: mNavigationInterrupt = interrupt
End Sub
Public Function NavigationReadyPath() As String
    NavigationReadyPath = mNavigationReady
End Function
Public Sub NavigationPresented()
    If mNavigationInterrupt Then
        mNavigationInterrupt = False
        modAuth.SignOut
    End If
End Sub
Public Sub CleanupNavigationRecording()
    mNavigationReady = "": mNavigationInterrupt = False
    Unload frmAdminSettings
End Sub
'@)
    $form.AddFromString(@'
Public Sub NavigationWriteGeometry()
    Dim path As String, channel As Integer, line As String
    path = TestD5Commands.NavigationReadyPath()
    If path = "" Or Not Me.Visible Then Exit Sub
    line = CStr(CDbl(modUserFormResizeWin.GetUserFormWindowHandle(Me))) & "|" & _
        CStr(mPages.Value) & "|" & CStr(mTrackingSections.Value) & "|" & _
        CStr(mPages.Left) & "|" & CStr(mPages.Top) & "|" & _
        CStr(mPages.Left + mTrackingSections.Left) & "|" & _
        CStr(mPages.Top + 23 + mTrackingSections.Top) & "|" & _
        CStr(Me.InsideWidth) & "|" & CStr(Me.InsideHeight) & "|" & _
        CStr(mBtnClose.Left + mBtnClose.Width / 2) & "|" & CStr(mBtnClose.Top + mBtnClose.Height / 2)
    channel = FreeFile
    Open path For Output As #channel
    Print #channel, line
    Close #channel
End Sub
'@)
    # Add read-only form event listeners alongside existing owner class listeners.
    foreach($field in @('mPages','mTrackingSections')){
        $found=$false
        for($line=1;$line -le $form.CountOfDeclarationLines;$line++){
            if($form.Lines($line,1).Trim() -ceq ('Private '+$field+' As MSForms.MultiPage')){
                $form.ReplaceLine($line,('Private WithEvents '+$field+' As MSForms.MultiPage'));$found=$true;break
            }
        }
        if(-not $found){throw 'Expected uninstrumented MultiPage declaration.'}
        $form.AddFromString("Private Sub ${field}_Change()`r`n    NavigationWriteGeometry`r`nEnd Sub")
    }
    $start=$form.ProcStartLine('UserForm_Activate',0);$end=$start+$form.ProcCountLines('UserForm_Activate',0)
    $hits=@(for($line=$start;$line -lt $end;$line++){if($form.Lines($line,1).Trim() -ceq 'End Sub'){$line}})
    if($hits.Count -ne 1){throw 'Expected unique visible activation boundary.'}
    $form.InsertLines($hits[0],"    TestD5Commands.NavigationPresented`r`n    NavigationWriteGeometry")
}

function Start-SettingsModalInput([string]$Ready,[string]$Trace,[bool]$Interrupted) {
    Start-Job -ArgumentList $repo,([long]$excel.Hwnd),$Ready,$Trace,$Interrupted -ScriptBlock {
        param($repo,$excelHandle,$ready,$trace,$interrupted)
        $ErrorActionPreference='Stop'
        . (Join-Path $repo 'tests/tooling/Slice4beAdminSettingsNavigation.ps1')
        Initialize-SettingsNavigationInput
        [SettingsNavigationCoordinates]::TracePath=$trace
        Add-Type @'
using System;using System.Runtime.InteropServices;
public static class ModalRecordingWindow {
 [DllImport("user32.dll")]public static extern uint GetWindowThreadProcessId(IntPtr h,out uint p);
 [DllImport("user32.dll")]public static extern bool IsWindowVisible(IntPtr h);
 [DllImport("user32.dll")]public static extern bool PostMessage(IntPtr h,uint m,IntPtr w,IntPtr l);
}
'@
        [uint32]$owner=0;[void][ModalRecordingWindow]::GetWindowThreadProcessId([IntPtr]$excelHandle,[ref]$owner)
        $created=(Get-Process -Id $owner).StartTime.ToUniversalTime().Ticks
        function Geometry([int]$outer,[int]$inner){
            $until=[DateTime]::UtcNow.AddSeconds(20)
            do {
                if((Get-Process -Id $owner).StartTime.ToUniversalTime().Ticks -ne $created){throw 'Owned Excel identity changed.'}
                $g=@();try{if(Test-Path -LiteralPath $ready){$g=([IO.File]::ReadAllText($ready).Trim()).Split('|')}}catch{}
                if($g.Count -eq 11 -and [int]$g[1] -eq $outer -and [int]$g[2] -eq $inner){return ,$g}
                Start-Sleep -Milliseconds 100
            }while([DateTime]::UtcNow -lt $until)
            throw 'Modal geometry/selection acknowledgment unavailable; not behavioral RED.'
        }
        $form=[IntPtr]::Zero;$completed=$false
        try {
            $g=Geometry 0 0;$form=[IntPtr][long][double]$g[0]
            [uint32]$formOwner=0;[void][ModalRecordingWindow]::GetWindowThreadProcessId($form,[ref]$formOwner)
            if($formOwner -ne $owner -or -not [ModalRecordingWindow]::IsWindowVisible($form)){throw 'Owned visible Settings form required.'}
            if(-not $interrupted){
                foreach($step in @(@($false,88,1,0),@($true,75,1,1),@($true,140,1,2),@($true,22,1,0),@($false,22,0,0))){
                    $x=if($step[0]){[double]$g[5]}else{[double]$g[3]};$y=if($step[0]){[double]$g[6]}else{[double]$g[4]}
                    [SettingsNavigationCoordinates]::Click($form,[IntPtr]$excelHandle,($x+[double]$step[1]),($y+7),[double]$g[7],[double]$g[8])
                    $g=Geometry ([int]$step[2]) ([int]$step[3])
                }
            }
            [SettingsNavigationCoordinates]::Click($form,[IntPtr]$excelHandle,[double]$g[9],[double]$g[10],[double]$g[7],[double]$g[8])
            $until=[DateTime]::UtcNow.AddSeconds(10)
            while([ModalRecordingWindow]::IsWindowVisible($form) -and [DateTime]::UtcNow -lt $until){Start-Sleep -Milliseconds 100}
            if([ModalRecordingWindow]::IsWindowVisible($form)){throw 'Actual Close did not dismiss Settings.'}
            $completed=$true
            [pscustomobject]@{Completed=$true;Clicks=$(if($interrupted){1}else{6})}
        }finally{
            # Owned normal-window cleanup only after a failed fixture, never acceptance input.
            if(-not $completed -and $form -ne [IntPtr]::Zero -and [ModalRecordingWindow]::IsWindowVisible($form)){
                [uint32]$p=0;[void][ModalRecordingWindow]::GetWindowThreadProcessId($form,[ref]$p)
                if($p -eq $owner -and (Get-Process -Id $owner).StartTime.ToUniversalTime().Ticks -eq $created){[void][ModalRecordingWindow]::PostMessage($form,16,[IntPtr]::Zero,[IntPtr]::Zero)}
            }
        }
    }
}

function Test-AdminNavigationRecorded($Fixture,$Other) {
    . (Join-Path $PSScriptRoot 'Slice4beRecordingFixture.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beRecordingEvaluation.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beEvaluationContracts.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beGuideTestActions.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beActivityAssertions.ps1')
    function Hash([string]$path){(Get-FileHash -LiteralPath $path).Hash}
    function Retained($before,$after){foreach($path in $before.Keys){if(-not $after.ContainsKey($path) -or $after[$path] -cne $before[$path]){return $false}};return $true}
    function Act([string]$name,[object[]]$values=@()){Run 'invSys.Admin.xlam' ('TestD5Commands.'+$name) $values}
    $steps=@(@('Settings','ADMIN_SETTINGS_OPEN','OPENED'),@('Event Tracking','ADMIN_SETTINGS_PAGE_EVENT_TRACKING','SELECTED'),@('Event Detail','ADMIN_SETTINGS_PAGE_EVENT_DETAIL','SELECTED'),@('Action Paths','ADMIN_SETTINGS_PAGE_ACTION_PATHS','SELECTED'),@('Tracking','ADMIN_SETTINGS_PAGE_TRACKING','SELECTED'),@('General','ADMIN_SETTINGS_PAGE_GENERAL','SELECTED'),@('Close','ADMIN_SETTINGS_CLOSE','CLOSED'))
    $owned=[IO.Path]::GetFullPath($runRoot).TrimEnd('\')+'\'
    if(-not ([IO.Path]::GetFullPath($Fixture.Root)+'\').StartsWith($owned,[StringComparison]::OrdinalIgnoreCase)){throw 'Generated recording fixture required.'}
    $configBytes=[IO.File]::ReadAllBytes($Fixture.Config);$configHash=Hash $Fixture.Config
    $otherPins=BoundPins $Other.Root;$book=$null;$runs=@()
    try {
        SelectTarget $Fixture 'config-admin'
        if(-not [bool](Run 'invSys.Core.xlam' 'TestSettingsOwnerState.SetupRecordingPolicy')){throw 'Actual default recording policy unavailable.'}
        $authority=@{};foreach($file in Get-ChildItem $Fixture.Root -File -Filter '*.xlsb'){$authority[$file.FullName]=Hash $file.FullName}
        $book=$excel.Workbooks.Add();$book.Worksheets.Item(1).Cells.Item(1,1).Value2='Navigation recording custom canary'
        # Public launcher may prepare the operator surface; pin saved host after preparation.
        [void](Run 'invSys.Admin.xlam' 'TestD5Commands.OpenSettings');[void](Act 'CloseSettings')
        $hostPath=Join-Path $runRoot 'navigation-recording-host.xlsb';$book.SaveAs($hostPath,50);$book.Close($false);$book=$null
        $hostHash=Hash $hostPath;$book=$excel.Workbooks.Open($hostPath,0,$false)
        foreach($kind in @('Source','Observed','Interrupted')){
            CloseRecordingViewer;SelectTarget $Fixture 'config-admin';OpenRecordingViewer
            $prior=BoundPins $journalRoot;$oldActivity=ActivityPins
            Check ('NavigationRecorded.'+$kind+'.ActualStart') ((RecordingControl 'Start Recording' 'Click') -ceq 'DELIVERED')
            $starts=@(Get-ChildItem $journalRoot -File -Filter '*.json'|Where-Object{-not $prior.ContainsKey($_.FullName)}|ForEach-Object{Get-Content $_.FullName -Raw|ConvertFrom-Json}|Where-Object RecordType -CEQ 'Start')
            if($starts.Count -ne 1){throw 'One durable Start required.'}
            $sequence=$starts[0].SequenceId;$before=@(Get-Slice4beActivityFiles $Fixture)
            $ready=Join-Path $reportRoot ($kind+'-modal.ready');$trace=Join-Path $reportRoot ($kind+'-input.jsonl')
            [void](Act 'ArmNavigationRecording' @($ready,($kind -ceq 'Interrupted')))
            $job=Start-SettingsModalInput $ready $trace ($kind -ceq 'Interrupted')
            try {
                [void](Run 'invSys.Admin.xlam' 'modAdmin.Open_Settings')
                [void](Wait-Job $job -Timeout 10)
                if($job.State -ne 'Completed'){Receive-Job $job -ErrorAction Stop;throw 'Native modal input job did not complete.'}
                $receipt=Receive-Job $job -ErrorAction Stop
                if(-not $receipt.Completed){throw 'Native modal journey unavailable.'}
            }finally{if($job.State -eq 'Running'){Stop-Job $job};Remove-Job $job;[void](Act 'ArmNavigationRecording' @('', $false))}
            $records=@(Get-Slice4beActivityFiles $Fixture|Where-Object{$_ -cnotin $before}|ForEach-Object{Get-Content $_ -Raw|ConvertFrom-Json})
            $count=if($kind -ceq 'Interrupted'){1}else{7}
            Check ('NavigationRecorded.'+$kind+'.ExactObservationCount') ($records.Count -eq 2*$count)
            for($ordinal=1;$ordinal -le $count;$ordinal++){
                $rows=@($records|Where-Object Ordinal -EQ $ordinal);$first=@($rows|Where-Object OutcomeCode -CEQ 'REQUESTED');$last=@($rows|Where-Object OutcomeCode -CEQ $steps[$ordinal-1][2])
                $valid=$rows.Count -eq 2 -and $first.Count -eq 1 -and $last.Count -eq 1
                if($valid){$valid=$first[0].ActivityId -ceq $last[0].ActivityId -and $first[0].RecordId -cne $last[0].RecordId}
                foreach($row in $rows){
                    $valid=$valid -and $row.ControlId -ceq $steps[$ordinal-1][1] -and $row.SequenceId -ceq $sequence -and $row.OwnerId -ceq 'ADMIN_SETTINGS_UI' -and $row.CatalogVersion -eq 33 -and $row.WarehouseId -ceq $Fixture.Warehouse -and $row.UserId -ceq 'config-admin' -and @($row.SourceEventRefs).Count -eq 0
                    $wire=ConvertTo-Json $row -Depth 20 -Compress
                    foreach($value in @($Fixture.Secret,(CredentialHash $Fixture.Secret),$Fixture.Root)){$valid=$valid -and -not $wire.Contains($value)}
                }
                Check ('NavigationRecorded.'+$kind+'.'+$ordinal+'.ExactPairAndContext') $valid
            }
            if($kind -cne 'Interrupted'){Check ('NavigationRecorded.'+$kind+'.ActualStop') ((RecordingControl 'Stop Recording' 'Click') -ceq 'DELIVERED')}
            $closed=@(RecordingJournal $sequence|Where-Object RecordType -CEQ 'Close')
            $valid=$closed.Count -eq 1
            if($valid){$valid=$closed[0].ActionCount -eq $count -and @($closed[0].Observations).Count -eq 2*$count -and $closed[0].Lifecycle -ceq $(if($kind -ceq 'Interrupted'){'Incomplete'}else{'Stopped'})}
            if($valid -and $kind -ceq 'Interrupted'){$valid=$closed[0].ReasonCode -ceq 'SESSION_CHANGED'}
            Check ('NavigationRecorded.'+$kind+'.ExactJournalLifecycle') $valid
            Check ('NavigationRecorded.'+$kind+'.ImmutableJournalChain') (JournalChain $sequence (2*$count+2))
            Check ('NavigationRecorded.'+$kind+'.PriorActivityImmutable') (PinsRetained $oldActivity)
            Check ('NavigationRecorded.'+$kind+'.AuthorityPreserved') (@($authority.Keys|Where-Object{(Hash $_) -cne $authority[$_]}).Count -eq 0)
            if($closed.Count -ne 1){throw 'Closed recording prerequisite unavailable.'}
            $runs+=@{Kind=$kind;Journal=$closed[0];Records=$records;Success=($kind -cne 'Interrupted');Incomplete=($kind -ceq 'Interrupted')}
        }
        CloseRecordingViewer;SelectTarget $Fixture 'config-admin'
        if(-not [bool](Run 'invSys.Admin.xlam' 'modAdminConsole.PublishReadFixtureForTest')){throw 'Actual publication prerequisite failed.'}
        $training=BoundPins $journalRoot
        . (Join-Path $PSScriptRoot 'Slice4beGeneralSettingsGuidePaths.ps1')
        Test-GeneralSettingsGuidePaths $Fixture $runs $steps -SettingsNavigation
        Check 'NavigationRecorded.TrainingEvidenceImmutable' (Retained $training (BoundPins $journalRoot))
        Check 'NavigationRecorded.OtherWarehousePreserved' (BoundSame $otherPins (BoundPins $Other.Root))
        Check 'NavigationRecorded.OperatorCanaryPreserved' ([string]$book.Worksheets.Item(1).Cells.Item(1,1).Value2 -ceq 'Navigation recording custom canary')
    }finally{
        [void](Act 'CleanupNavigationRecording');CloseRecordingViewer
        if($null -ne $book){$book.Close($false)}
        [IO.File]::WriteAllBytes($Fixture.Config,$configBytes);SelectTarget $Fixture 'config-admin'
    }
    Check 'NavigationRecorded.FinalConfigAndHostPreserved' ((Hash $Fixture.Config) -ceq $configHash -and (Hash $hostPath) -ceq $hostHash)
}
