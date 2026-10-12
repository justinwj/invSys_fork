# Observe the real modal launcher after its presentation hook, then perform internal cleanup.
function Install-AdminSettingsEntryFailureProbe($TestModule,$FormCode) {
    $anchor='If TestD5Commands.DriveDefaultClose(Me) Then Exit Sub'
    $hits=@(for($line=1;$line -le $FormCode.CountOfLines;$line++){if($FormCode.Lines($line,1).Trim() -ceq $anchor){$line}})
    if($hits.Count -ne 1){throw 'Unique modal observation boundary required; not behavioral RED.'}
    $FormCode.InsertLines($hits[0],'    If TestD5Commands.DriveEntryBoundary(Me) Then Exit Sub')
    $FormCode.AddFromString(@'
Public Function EntryBoundaryNoticeForTest() As String
    EntryBoundaryNoticeForTest = mLblStatus.Caption
End Function
'@)
    $TestModule.CodeModule.AddFromString(@'
Private mBoundaryArmed As Boolean
Private mBoundaryVisibleCount As Long
Private mBoundaryNotice As String
Public Sub EntryBoundaryArm(ByVal enabled As Boolean)
    mBoundaryArmed = enabled
    If enabled Then mBoundaryNotice = ""
End Sub
Public Function DriveEntryBoundary(ByVal instance As frmAdminSettings) As Boolean
    If Not mBoundaryArmed Then Exit Function
    mBoundaryArmed = False
    If instance.Visible Then mBoundaryVisibleCount = mBoundaryVisibleCount + 1
    mBoundaryNotice = instance.EntryBoundaryNoticeForTest()
    Unload instance
    DriveEntryBoundary = True
End Function
Public Function EntryBoundaryCount() As Long
    EntryBoundaryCount = mBoundaryVisibleCount
End Function
Public Function EntryBoundaryNotice() As String
    EntryBoundaryNotice = mBoundaryNotice
End Function
'@)
}

function Test-AdminSettingsEntryFailures($Fixture,$Other) {
    . (Join-Path $PSScriptRoot 'Slice4beActivityAssertions.ps1')
    function Act([string]$name,[object[]]$values=@()){Run 'invSys.Admin.xlam' ('TestD5Commands.'+$name) $values}
    function Probe([string]$name,[object[]]$values=@()){Run 'invSys.Core.xlam' ('TestAdminEntry.'+$name) $values}
    function Files {@(Get-Slice4beActivityFiles $Fixture)}
    function Policy([string]$mode='Enabled'){if(-not [bool](Probe 'EnablePolicy' @($mode))){throw 'Launcher policy fixture unavailable; not behavioral RED.'}}
    function Open([string]$label,[bool]$visible,[string]$outcome='',[string]$notice='') {
        $before=@(Files);$count=[int](Act 'EntryBoundaryCount');$context=[string](Run 'invSys.Core.xlam' 'modActivity.CaptureContext')
        [void](Act 'EntryBoundaryArm' @($true))
        try{[void](Run 'invSys.Admin.xlam' 'modAdmin.Open_Settings')}
        finally{[void](Act 'EntryBoundaryArm' @($false));[void](Act 'CleanupDefaultSettings')}
        Check ($label+'.ActualPresentation') (([int](Act 'EntryBoundaryCount') - $count) -eq [int]$visible)
        Check ($label+'.NoHiddenDefaultInstance') ([int](Act 'RemainingSettingsInstances') -eq 0)
        Check ($label+'.CapturedSessionPreserved') ($context -cne '' -and [string](Run 'invSys.Core.xlam' 'modActivity.CaptureContext') -ceq $context)
        $rows=@(Files|Where-Object {$_ -cnotin $before}|ForEach-Object{Get-Content $_ -Raw|ConvertFrom-Json})
        if($outcome){
            $first=@($rows|Where-Object OutcomeCode -CEQ 'REQUESTED');$last=@($rows|Where-Object OutcomeCode -CEQ $outcome)
            $valid=$rows.Count -eq 2 -and $first.Count -eq 1 -and $last.Count -eq 1
            if($valid){$valid=$first[0].ActivityId -ceq $last[0].ActivityId -and $first[0].RecordId -cne $last[0].RecordId}
            foreach($row in $rows){
                $terminal=$row.OutcomeCode -ceq $outcome;$success=$row.OutcomeCode -ceq 'OPENED'
                $severity=if($terminal -and $outcome -ceq 'DENIED'){'Blocked'}else{'Info'}
                $valid=$valid -and $row.ControlId -ceq 'ADMIN_SETTINGS_OPEN' -and $row.OwnerId -ceq 'ADMIN_SETTINGS_UI' -and $row.WarehouseId -ceq $Fixture.Warehouse -and $row.UserId -ceq 'config-admin' -and $row.StationId -ceq 'S1' -and $row.CatalogVersion -eq 33 -and $row.Severity -ceq $severity -and $row.DataEffect -ceq $(if($terminal){'Unchanged'}else{'Unknown'}) -and $row.EventCode -ceq ('ADMIN_SETTINGS_OPEN_'+$row.OutcomeCode) -and @($row.SourceEventRefs).Count -eq 0
                if($terminal){$message=if($success){'Settings form opened. No settings save or Domain application is asserted.'}else{'Settings could not be opened in the current context.'};$valid=$valid -and $row.UserMessage -ceq $message}
                $wire=ConvertTo-Json $row -Depth 20 -Compress
                $valid=$valid -and ([bool](Probe 'Concludes' @($wire)) -eq $success)
                foreach($value in @($Fixture.Secret,(CredentialHash $Fixture.Secret),$Fixture.Root,'PinHash')){$valid=$valid -and -not $wire.Contains($value)}
            }
            Check ($label+'.ExactOwnerOutcomeAndConclusion') $valid
        }else{Check ($label+'.NoOptionalActivity') ($rows.Count -eq 0)}
        if($notice){Check ($label+'.PresentedTrackingNotice') ([string](Act 'EntryBoundaryNotice') -ceq $notice)}
    }
    $pins=@{};foreach($file in Files){$pins[$file]=(Get-FileHash $file).Hash}
    if($pins.Count -lt 4){throw 'Public-launcher baseline history required.'}
    $otherPins=@{};foreach($file in Get-ChildItem $Other.Root -Recurse -File){$otherPins[$file.FullName]=(Get-FileHash $file.FullName).Hash}
    $authorityPins=@{};foreach($file in Get-ChildItem $Fixture.Root -File -Filter '*.xlsb'){if($file.FullName -cne $Fixture.Config){$authorityPins[$file.FullName]=(Get-FileHash $file.FullName).Hash}}
    if($authorityPins.Count -lt 3){throw 'Generated Domain/Auth authority fixtures required.'}
    $book=$null
    try {
        SelectTarget $Fixture;Policy
        $book=$excel.Workbooks.Add();$book.Worksheets.Item(1).Cells.Item(1,1).Value2='Launcher boundary host'
        $path=Join-Path $runRoot 'settings-entry-boundaries.xlsb';$book.SaveAs($path,50)
        # The existing launcher may prepare operator surfaces. Establish that baseline first.
        Open 'AdminEntryBoundary.InitialSurfacePreparation' $true 'OPENED'
        $book.Save();$book.Close($false);$book=$null;$hostHash=(Get-FileHash $path).Hash;$book=$excel.Workbooks.Open($path)
        foreach($mode in @('Defaults','Disabled','UserOff','OpenOff','Enabled')){
            Policy $mode;$hash=(Get-FileHash $Fixture.Config).Hash
            Open ('AdminEntryBoundary.'+$mode) $true $(if($mode -ceq 'Enabled'){'OPENED'}else{''})
            Check ('AdminEntryBoundary.'+$mode+'.PolicyPreserved') ((Get-FileHash $Fixture.Config).Hash -ceq $hash)
        }
        $configBytes=[IO.File]::ReadAllBytes($Fixture.Config);$configHash=(Get-FileHash $Fixture.Config).Hash
        $legacyIds=([string](Probe 'HistoricalPolicyIds')).Split("`n")
        try {
            $cfg=$excel.Workbooks.Open($Fixture.Config,0,$false)
            try {
                (Table $cfg 'tblEventTrackingPolicies').ListColumns.Item('CatalogVersion').DataBodyRange.Value2=32.0
                $controls=Table $cfg 'tblEventTrackingControls'
                for($i=$controls.ListRows.Count;$i -ge 1;$i--){if($controls.ListRows.Item($i).Range.Cells.Item(1,$controls.ListColumns.Item('ControlId').Index).Value2 -cnotin $legacyIds){$controls.ListRows.Item($i).Delete()}}
                $cfg.Save()
            }finally{$cfg.Close($false);$cfg=$null}
            if(-not [bool](Probe 'ExistingPolicyReadable')){throw 'Readable catalog32 policy required; not behavioral RED.'}
            Open 'AdminEntryBoundary.OlderPolicy' $true '' 'Tracking unavailable: the saved policy does not include this control.'
        }finally{[IO.File]::WriteAllBytes($Fixture.Config,$configBytes)}
        Check 'AdminEntryBoundary.OlderPolicy.ConfigRestored' ((Get-FileHash $Fixture.Config).Hash -ceq $configHash)
        $blocked=Join-Path (Join-Path $Fixture.Root 'Training\Activity') $Fixture.Warehouse;$held=$blocked+'-entry-held'
        foreach($item in @($blocked,$held)){if(-not [IO.Path]::GetFullPath($item).StartsWith([IO.Path]::GetFullPath($Fixture.Root).TrimEnd('\')+'\',[StringComparison]::OrdinalIgnoreCase)){throw 'Storage fault escaped generated fixture.'}}
        if(Test-Path -LiteralPath $held){throw 'Preserve existing held fixture.'}
        $moved=$false;$beforeFault=@(Files)
        try {
            if(Test-Path -LiteralPath $blocked){Move-Item -LiteralPath $blocked -Destination $held;$moved=$true}
            [IO.File]::WriteAllText($blocked,'Disposable launcher activity-store fault')
            Open 'AdminEntryBoundary.StorageUnavailable' $true '' 'Tracking unavailable: the training record could not be saved.'
        }finally{if(Test-Path -LiteralPath $blocked -PathType Leaf){Remove-Item -LiteralPath $blocked};if($moved){Move-Item -LiteralPath $held -Destination $blocked}}
        Check 'AdminEntryBoundary.Storage.NoInventedRecordsAfterRestore' (@(Files).Count -eq $beforeFault.Count)
        $authPath=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Auth.xlsb');$authBytes=[IO.File]::ReadAllBytes($authPath);$authHash=(Get-FileHash $authPath).Hash
        try {
            $context=[string](Run 'invSys.Core.xlam' 'modActivity.CaptureContext')
            $auth=$excel.Workbooks.Open($authPath,0,$false)
            try {
                $caps=Table $auth 'tblCapabilities';$changed=0
                foreach($row in $caps.ListRows){if($row.Range.Cells.Item(1,$caps.ListColumns.Item('UserId').Index).Value2 -ceq 'config-admin' -and $row.Range.Cells.Item(1,$caps.ListColumns.Item('Capability').Index).Value2 -ceq 'ADMIN_MAINT'){$row.Range.Cells.Item(1,$caps.ListColumns.Item('Status').Index).Value2='Inactive';$changed++}}
                $auth.Save()
            }finally{$auth.Close($false);$auth=$null}
            [void](Run 'invSys.Core.xlam' 'modAuth.ReloadAuth' @($Fixture.Warehouse))
            if($changed -eq 0 -or [bool](Run 'invSys.Core.xlam' 'modRoleUiAccess.CanCurrentUserPerformCapabilityCached' @('ADMIN_MAINT')) -or $context -ceq '' -or [string](Run 'invSys.Core.xlam' 'modActivity.CaptureContext') -cne $context){throw 'Same-session permission-loss fixture unavailable; not behavioral RED.'}
            $ready=Join-Path $runRoot 'entry-denial.ready';$stop=Join-Path $runRoot 'entry-denial.stop'
            $job=Start-Job -ArgumentList $repo,([long]$excel.Hwnd),$ready,$stop -ScriptBlock {
                param($repo,$handle,$ready,$stop)
                $ErrorActionPreference='Stop'
                . (Join-Path $repo 'tools/plan022-dialog-observer.ps1')
                Invoke-Plan022NativeDialogObservation -ProcessId 0 -TimeoutSeconds 0
                [uint32]$owner=0;[void][Plan022NativeDialogs]::GetWindowThreadProcessId([IntPtr]$handle,[ref]$owner)
                if(-not $owner){throw 'Owned denial process unavailable.'}
                $created=(Get-Process -Id $owner).StartTime.ToUniversalTime().Ticks
                [IO.File]::WriteAllText($ready,'Ready');$seen=$false;$until=[DateTime]::UtcNow.AddSeconds(60)
                while(-not (Test-Path -LiteralPath $stop) -and [DateTime]::UtcNow -lt $until){
                    if((Get-Process -Id $owner).StartTime.ToUniversalTime().Ticks -ne $created){throw 'Denial process identity changed.'}
                    $texts=@([Plan022NativeDialogs]::Poll($owner))
                    if(@($texts|Where-Object{$_ -like '*ControlType.Text*Current user lacks ADMIN_MAINT capability.*'}).Count){$seen=$true}
                    Start-Sleep -Milliseconds 100
                }
                [pscustomobject]@{DenialVisible=$seen}
            }
            try {
                $until=[DateTime]::UtcNow.AddSeconds(12)
                while(-not(Test-Path -LiteralPath $ready) -and $job.State -eq 'Running' -and [DateTime]::UtcNow -lt $until){Start-Sleep -Milliseconds 100}
                if(-not(Test-Path -LiteralPath $ready)){throw 'Denial observer unavailable; not behavioral RED.'}
                Open 'AdminEntryBoundary.PermissionDenied' $false 'DENIED'
                [IO.File]::WriteAllText($stop,'Stop');[void](Wait-Job $job -Timeout 5)
                if($job.State -ne 'Completed'){throw 'Denial observer did not finish.'}
                $result=Receive-Job $job -ErrorAction Stop
                Check 'AdminEntryBoundary.PermissionDenied.VisibleDenial' ([bool]$result.DenialVisible)
            }finally{[IO.File]::WriteAllText($stop,'Stop');if($job.State -eq 'Running'){Stop-Job $job};Remove-Job $job}
        }finally{[IO.File]::WriteAllBytes($authPath,$authBytes);[void](Run 'invSys.Core.xlam' 'modAuth.ReloadAuth' @($Fixture.Warehouse))}
        Check 'AdminEntryBoundary.AuthRestored' ((Get-FileHash $authPath).Hash -ceq $authHash)
        Check 'AdminEntryBoundary.ConfigPreserved' ((Get-FileHash $Fixture.Config).Hash -ceq $configHash)
        Check 'AdminEntryBoundary.OlderRecordsImmutable' (@($pins.Keys|Where-Object{(Get-FileHash $_).Hash -cne $pins[$_]}).Count -eq 0)
        Check 'AdminEntryBoundary.DomainAndAuthBytesPreserved' (@($authorityPins.Keys|Where-Object{(Get-FileHash $_).Hash -cne $authorityPins[$_]}).Count -eq 0)
        Check 'AdminEntryBoundary.OtherWarehousePreserved' (@($otherPins.Keys|Where-Object{(Get-FileHash $_).Hash -cne $otherPins[$_]}).Count -eq 0)
        $book.Close($false);$book=$null
        Check 'AdminEntryBoundary.SavedHostPreserved' ((Get-FileHash $path).Hash -ceq $hostHash)
    }finally{[void](Act 'EntryBoundaryArm' @($false));[void](Act 'CleanupDefaultSettings');if($null -ne $book){$book.Close($false)};SelectTarget $Fixture}
}
