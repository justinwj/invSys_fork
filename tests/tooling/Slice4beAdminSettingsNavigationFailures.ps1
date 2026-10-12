# Existing D18 navigation boundaries, through real MSForms input on the packaged form.
function Test-AdminSettingsNavigationFailures($Fixture,$Other) {
    . (Join-Path $PSScriptRoot 'Slice4beActivityAssertions.ps1')
    function Act([string]$Name,[object[]]$Values=@()){Run 'invSys.Admin.xlam' ('TestD5Commands.'+$Name) $Values}
    function Probe([string]$Name,[object[]]$Values=@()){Run 'invSys.Core.xlam' ('TestAdminEntry.'+$Name) $Values}
    function Files {@(Get-Slice4beActivityFiles $Fixture)}
    function Policy([string]$mode='Enabled') {
        if(-not [bool](Probe 'EnablePolicy' @($mode))){throw 'Navigation policy fixture unavailable; not behavioral RED.'}
    }
    function Open {[void](Act 'OpenSettings');[void](Act 'ShowSettings')}
    function Visit($case,[string]$label,[bool]$recorded,[string]$notice='') {
        $before=@(Files);$otherBefore=@(Get-Slice4beActivityFiles $Other)
        [void](Act 'EntryNavPrepare' @($case[1],[int]$case[2]))
        $h=[IntPtr][long][double](Act 'EntryNavHandle')
        $g=([string](Act 'EntryNavGeometry' @($case[1]))).Split('|')
        [SettingsNavigationCoordinates]::Click($h,[IntPtr]$excel.Hwnd,([double]$g[0]+$case[4]),([double]$g[1]+7),[double]$g[2],[double]$g[3])
        Start-Sleep -Milliseconds 250
        if([int](Act 'EntryNavState' @($case[1])) -ne $case[3] -or -not ([string](Act 'EntryNavTrace')).Contains('_MouseDown;')){throw 'Native failure-case selection unavailable; not behavioral RED.'}
        Check ($label+'.LocalSelectionStillWorks') $true
        $rows=@(Files|Where-Object {$_ -cnotin $before}|ForEach-Object {Get-Content $_ -Raw|ConvertFrom-Json})
        $id='ADMIN_SETTINGS_PAGE_'+$case[0]
        if($recorded){
            $first=@($rows|Where-Object OutcomeCode -CEQ 'REQUESTED');$last=@($rows|Where-Object OutcomeCode -CEQ 'SELECTED')
            $valid=$rows.Count -eq 2 -and $first.Count -eq 1 -and $last.Count -eq 1
            if($valid){$valid=$first[0].ActivityId -ceq $last[0].ActivityId -and $first[0].RecordId -cne $last[0].RecordId}
            foreach($row in $rows){
                $positive=$row.OutcomeCode -ceq 'SELECTED'
                $valid=$valid -and $row.ControlId -ceq $id -and $row.UserId -ceq 'config-admin' -and $row.WarehouseId -ceq $Fixture.Warehouse -and $row.OwnerId -ceq 'ADMIN_SETTINGS_UI' -and $row.CatalogVersion -eq 33 -and $row.Severity -ceq 'Info' -and $row.DataEffect -ceq $(if($positive){'Unchanged'}else{'Unknown'}) -and $row.EventCode -ceq ($id+'_'+$row.OutcomeCode) -and @($row.SourceEventRefs).Count -eq 0
                if($positive){$valid=$valid -and $row.UserMessage -ceq 'Settings page selected. No settings save or Domain application is asserted.'}
                $wire=ConvertTo-Json $row -Depth 20 -Compress
                $valid=$valid -and ([bool](Probe 'Concludes' @($wire)) -eq $positive)
                foreach($value in @($Fixture.Secret,(CredentialHash $Fixture.Secret),$Fixture.Root,'PinHash')){$valid=$valid -and -not $wire.Contains($value)}
            }
            Check ($label+'.ExactLocalPresentationPair') $valid
        }else{Check ($label+'.NoOptionalOrRetargetedRecords') ($rows.Count -eq 0 -and @(Get-Slice4beActivityFiles $Other).Count -eq $otherBefore.Count)}
        if($notice){
            $text=[string](Act 'EntryNavNotice')
            Check ($label+'.VisibleFixedNotice') ($text -ceq $notice -and -not $text.Contains($Fixture.Root) -and -not $text.Contains($Fixture.Secret))
        }
    }
    $cases=@(@('GENERAL',$false,1,0,22),@('EVENT_TRACKING',$false,0,1,88),@('TRACKING',$true,1,0,22),@('EVENT_DETAIL',$true,0,1,75),@('ACTION_PATHS',$true,1,2,140))
    $pins=@{};foreach($file in Files){$pins[$file]=(Get-FileHash $file).Hash}
    if($pins.Count -lt 20){throw 'Nonempty baseline history required; not behavioral RED.'}
    $otherPins=@{};foreach($file in Get-ChildItem $Other.Root -Recurse -File){$otherPins[$file.FullName]=(Get-FileHash $file.FullName).Hash}
    $authorityPins=@{};foreach($file in Get-ChildItem $Fixture.Root -File -Filter '*.xlsb'){if($file.FullName -cne $Fixture.Config){$authorityPins[$file.FullName]=(Get-FileHash $file.FullName).Hash}}
    if($authorityPins.Count -lt 3){throw 'Generated Domain/Auth authority fixtures required.'}
    $book=$null
    try {
        SelectTarget $Fixture
        $book=$excel.Workbooks.Add();$book.Worksheets.Item(1).Cells.Item(1,1).Value2='Navigation boundary host'
        $path=Join-Path $runRoot 'settings-navigation-boundaries.xlsb';$book.SaveAs($path,50);$book.Close($false);$book=$null
        $hostHash=(Get-FileHash $path).Hash;$book=$excel.Workbooks.Open($path)
        # Change policy while the same visible form remains open: no stale policy cache.
        Policy;Open
        foreach($mode in @('Defaults','Disabled','UserOff','ControlsOff','Enabled')){
            Policy $mode;$hash=(Get-FileHash $Fixture.Config).Hash
            foreach($case in $cases){Visit $case ('AdminNavigationBoundary.'+$mode+'.'+$case[0]) ($mode -ceq 'Enabled')}
            Check ('AdminNavigationBoundary.'+$mode+'.PolicyPreserved') ((Get-FileHash $Fixture.Config).Hash -ceq $hash)
        }
        $before=@(Files)
        [void](Act 'EntryNavReactivate');[void](Act 'EntryNavReactivate')
        Check 'AdminNavigationBoundary.RepeatedActivationExcluded' (@(Files).Count -eq $before.Count)
        [void](Act 'CloseSettings')
        foreach($mode in @('ChangedTarget','SignedOut')){
            SelectTarget $Fixture;Open
            if($mode -ceq 'ChangedTarget'){SelectTarget $Other}else{[void](Run 'invSys.Core.xlam' 'modAuth.SignOut')}
            foreach($case in $cases){Visit $case ('AdminNavigationBoundary.'+$mode+'.'+$case[0]) $false 'Session or warehouse changed. Reopen Settings.'}
            [void](Act 'CloseSettings')
        }
        SelectTarget $Fixture
        $authPath=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Auth.xlsb')
        $authBytes=[IO.File]::ReadAllBytes($authPath);$authHash=(Get-FileHash $authPath).Hash
        try {
            Open;$context=[string](Run 'invSys.Core.xlam' 'modActivity.CaptureContext')
            $auth=$excel.Workbooks.Open($authPath,0,$false)
            try {
                $caps=Table $auth 'tblCapabilities';$changed=0
                foreach($row in $caps.ListRows){if($row.Range.Cells.Item(1,$caps.ListColumns.Item('UserId').Index).Value2 -ceq 'config-admin' -and $row.Range.Cells.Item(1,$caps.ListColumns.Item('Capability').Index).Value2 -ceq 'ADMIN_MAINT'){$row.Range.Cells.Item(1,$caps.ListColumns.Item('Status').Index).Value2='Inactive';$changed++}}
                $auth.Save()
            }finally{$auth.Close($false);$auth=$null}
            [void](Run 'invSys.Core.xlam' 'modAuth.ReloadAuth' @($Fixture.Warehouse))
            if($changed -eq 0 -or [bool](Run 'invSys.Core.xlam' 'modRoleUiAccess.CanCurrentUserPerformCapabilityCached' @('ADMIN_MAINT')) -or $context -ceq '' -or [string](Run 'invSys.Core.xlam' 'modActivity.CaptureContext') -cne $context){throw 'Same-session permission-loss fixture unavailable; not behavioral RED.'}
            foreach($case in $cases){Visit $case ('AdminNavigationBoundary.PermissionLoss.'+$case[0]) $true}
        }finally{[void](Act 'CloseSettings');[IO.File]::WriteAllBytes($authPath,$authBytes);[void](Run 'invSys.Core.xlam' 'modAuth.ReloadAuth' @($Fixture.Warehouse))}
        Check 'AdminNavigationBoundary.AuthRestored' ((Get-FileHash $authPath).Hash -ceq $authHash)
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
            Open
            foreach($case in $cases){Visit $case ('AdminNavigationBoundary.OlderPolicy.'+$case[0]) $false 'Tracking unavailable: the saved policy does not include this control.'}
        }finally{[void](Act 'CloseSettings');[IO.File]::WriteAllBytes($Fixture.Config,$configBytes)}
        Check 'AdminNavigationBoundary.OlderPolicy.ConfigRestored' ((Get-FileHash $Fixture.Config).Hash -ceq $configHash)
        $blocked=Join-Path (Join-Path $Fixture.Root 'Training\Activity') $Fixture.Warehouse;$held=$blocked+'-navigation-held'
        foreach($item in @($blocked,$held)){if(-not [IO.Path]::GetFullPath($item).StartsWith([IO.Path]::GetFullPath($Fixture.Root).TrimEnd('\')+'\',[StringComparison]::OrdinalIgnoreCase)){throw 'Storage fault escaped generated fixture.'}}
        if(Test-Path -LiteralPath $held){throw 'Preserve existing held fixture.'}
        Open;$moved=$false;$beforeFault=@(Files)
        try {
            if(Test-Path -LiteralPath $blocked){Move-Item -LiteralPath $blocked -Destination $held;$moved=$true}
            [IO.File]::WriteAllText($blocked,'Disposable navigation activity-store fault')
            foreach($case in $cases){Visit $case ('AdminNavigationBoundary.StorageUnavailable.'+$case[0]) $false 'Tracking unavailable: the training record could not be saved.'}
            CaptureOwnedFormByCaptionEvidence 'invSys Settings' 'admin-navigation-storage-notice.png'
        }finally{
            [void](Act 'CloseSettings')
            if(Test-Path -LiteralPath $blocked -PathType Leaf){Remove-Item -LiteralPath $blocked}
            if($moved){Move-Item -LiteralPath $held -Destination $blocked}
        }
        Check 'AdminNavigationBoundary.Storage.ConfigPreserved' ((Get-FileHash $Fixture.Config).Hash -ceq $configHash)
        Check 'AdminNavigationBoundary.Storage.NoInventedRecordsAfterRestore' (@(Files).Count -eq $beforeFault.Count)
        Check 'AdminNavigationBoundary.OlderRecordsImmutable' (@($pins.Keys|Where-Object{(Get-FileHash $_).Hash -cne $pins[$_]}).Count -eq 0)
        Check 'AdminNavigationBoundary.DomainAndAuthBytesPreserved' (@($authorityPins.Keys|Where-Object{(Get-FileHash $_).Hash -cne $authorityPins[$_]}).Count -eq 0)
        Check 'AdminNavigationBoundary.OtherWarehousePreserved' (@($otherPins.Keys|Where-Object{(Get-FileHash $_).Hash -cne $otherPins[$_]}).Count -eq 0)
        $book.Close($false);$book=$null
        Check 'AdminNavigationBoundary.SavedHostPreserved' ((Get-FileHash $path).Hash -ceq $hostHash)
    }finally{[void](Act 'CloseSettings');if($null -ne $book){$book.Close($false)};SelectTarget $Fixture}
}
