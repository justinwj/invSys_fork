# D18 catalog32: actual Close button/native teardown, without replacing either owner.
function Install-AdminCloseActivityProbe($TestModule) {
    $TestModule.CodeModule.AddFromString(@'
Public Sub AdminCloseButtonForTest()
    mForm.PolicyTestCloseAction
End Sub
Public Function AdminCloseContextForTest() As String
    AdminCloseContextForTest = modActivity.CaptureContext()
End Function
'@)
    $probe=$packages['invSys.Core.xlam'].VBProject.VBComponents.Add(1)
    $probe.Name='TestAdminClose'
    $probe.CodeModule.AddFromString(@'
Option Explicit
Public Function EnableDefaultPolicy(Optional ByVal enabled As Boolean = True, Optional ByVal excludedUser As String = "") As Boolean
    Dim context As String, version As Long, request As String, report As String
    Dim model As Object, user As Object, users As Collection
    context = modActivity.CaptureContext()
    If Not modTrackingPolicySettings.ReadEditor(context, version, request, report) Then Exit Function
    Set model = modTrackingPolicyModel.Defaults(enabled)
    If excludedUser <> "" Then
        Set user = CreateObject("Scripting.Dictionary")
        user.Add "UserId", excludedUser: user.Add "Record", False
        Set users = model("Users"): users.Add user
    End If
    request = modTrainingJson.EncodeObject(model)
    EnableDefaultPolicy = modTrackingPolicySettings.SavePolicy(context, version, request, report)
End Function
Public Function Definition(ByVal version As Long) As String
    Dim value As Object
    Set value = modActivityCatalog.Control("ADMIN_SETTINGS_CLOSE", version)
    If Not value Is Nothing Then Definition = modTrainingJson.EncodeObject(value)
End Function
Public Function Concludes(ByVal wire As String) As Boolean
    Concludes = modEvaluationMatches.CommandCompleted(modTrainingJson.DecodeObject(wire))
End Function
Public Function ExistingPolicyReadable() As Boolean
    Dim version As Long, collect As Boolean, visible As Boolean, notice As String
    ExistingPolicyReadable = modActivityPolicy.ReadPolicy(modNasConnection.GetCurrentTarget(), "ADMIN_SETTINGS_SAVE_VALUE", version, collect, visible, notice)
    ExistingPolicyReadable = ExistingPolicyReadable And collect
End Function
'@)
}

function Close-AdminNativeForTest {
    if(-not ('AdminCloseNative' -as [type])){
        Add-Type @'
using System; using System.Runtime.InteropServices;
public static class AdminCloseNative {
 [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h,out uint id);
 [DllImport("user32.dll")] public static extern bool IsWindow(IntPtr h);
 [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
 [DllImport("user32.dll",SetLastError=true)] public static extern bool PostMessage(IntPtr h,uint message,IntPtr w,IntPtr l);
}
'@
    }
    $window=[InvSysSettingsCapture]::OwnedVisibleForm('invSys Settings',[IntPtr]$excel.Hwnd)
    [uint32]$owner=0;[uint32]$actual=0
    [void][AdminCloseNative]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd,[ref]$owner)
    [void][AdminCloseNative]::GetWindowThreadProcessId($window,[ref]$actual)
    if(-not $owner -or $owner -ne $actual -or -not [AdminCloseNative]::IsWindowVisible($window)){throw 'Owned visible Settings fixture unavailable; not behavioral RED.'}
    $posted=[AdminCloseNative]::PostMessage($window,0x10,[IntPtr]::Zero,[IntPtr]::Zero)
    $errorCode=if($posted){0}else{[Runtime.InteropServices.Marshal]::GetLastWin32Error()}
    if(-not $posted){throw ('Native Close delivery failed; Win32 error '+$errorCode)}
    $until=[DateTime]::UtcNow.AddSeconds(10)
    while([AdminCloseNative]::IsWindow($window) -and [DateTime]::UtcNow -lt $until){Start-Sleep -Milliseconds 100}
    return -not [AdminCloseNative]::IsWindow($window)
}

function Test-AdminCloseActivity($Fixture,$Other) {
    function Files {@(Get-Slice4beActivityFiles $Fixture)}
    function Act([string]$name){Run 'invSys.Admin.xlam' ('TestD5Commands.'+$name)}
    function Open {
        [void](Act 'OpenSettings');[void](Act 'ShowSettings')
        if([int](Act 'RemainingSettingsInstances') -ne 1){throw 'One Settings form required; not behavioral RED.'}
    }
    function Dismissed {([int](Act 'RemainingSettingsInstances') -eq 0 -and [InvSysSettingsCapture]::OwnedVisibleForm('invSys Settings',[IntPtr]$excel.Hwnd) -eq [IntPtr]::Zero)}
    function Pair([string[]]$before,[string]$label) {
        $raw=@(Files|Where-Object{$_ -cnotin $before}|ForEach-Object{[IO.File]::ReadAllText($_)})
        $rows=@($raw|ForEach-Object{$_|ConvertFrom-Json})
        $attempt=@($rows|Where-Object OutcomeCode -ceq 'REQUESTED');$closed=@($rows|Where-Object OutcomeCode -ceq 'CLOSED')
        $pair=$rows.Count -eq 2 -and $attempt.Count -eq 1 -and $closed.Count -eq 1
        $facts=$pair;$safe=$pair;$integrity=$pair;$linked=$false;$terminal=$false
        foreach($row in $rows){$facts=$facts -and $row.ControlId -ceq 'ADMIN_SETTINGS_CLOSE' -and $row.OwnerId -ceq 'ADMIN_SETTINGS_UI' -and $row.UserId -ceq 'config-admin' -and $row.WarehouseId -ceq $Fixture.Warehouse -and $row.CatalogVersion -eq 32 -and @($row.SourceEventRefs).Count -eq 0}
        foreach($text in $raw){
            foreach($value in @($Fixture.Secret,(CredentialHash $Fixture.Secret),$Fixture.Root,$canary,'PinHash','mBtn')){if($text.Contains($value)){$safe=$false}}
            $match=[regex]::Match($text,',"ContentSha256":"([a-f0-9]{64})"\}$')
            if(-not $match.Success){$integrity=$false;continue}
            $sha=[Security.Cryptography.SHA256]::Create()
            try{$hash=[BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($text.Substring(0,$match.Index)+'}'))).Replace('-','').ToLowerInvariant()}finally{$sha.Dispose()}
            $integrity=$integrity -and $hash -ceq $match.Groups[1].Value
        }
        if($pair){
            $linked=$attempt[0].ActivityId -cne '' -and $attempt[0].ActivityId -ceq $closed[0].ActivityId -and $attempt[0].RecordId -cne $closed[0].RecordId
            $facts=$facts -and $attempt[0].Severity -ceq 'Info' -and $attempt[0].DataEffect -ceq 'Unknown' -and $attempt[0].EventCode -ceq 'ADMIN_SETTINGS_CLOSE_REQUESTED' -and $closed[0].Severity -ceq 'Info' -and $closed[0].DataEffect -ceq 'Unchanged' -and $closed[0].EventCode -ceq 'ADMIN_SETTINGS_CLOSE_CLOSED' -and $closed[0].UserMessage -ceq 'Settings form dismissed; unsaved form edits discarded. No settings save or Domain application is asserted.'
            $terminal=-not [bool](Run 'invSys.Core.xlam' 'TestAdminClose.Concludes' @($raw|Where-Object{($_|ConvertFrom-Json).OutcomeCode -ceq 'REQUESTED'})) -and [bool](Run 'invSys.Core.xlam' 'TestAdminClose.Concludes' @($raw|Where-Object{($_|ConvertFrom-Json).OutcomeCode -ceq 'CLOSED'}))
        }
        Check ($label+'.OneAttemptAndClosed') $pair
        Check ($label+'.DistinctCorrelatedRecords') $linked
        Check ($label+'.ExactContextAndDismissalFacts') $facts
        Check ($label+'.Redacted') $safe
        Check ($label+'.ContentIntegrity') $integrity
        Check ($label+'.OnlyDismissalConcludes') $terminal
    }
    $book=$null;$canary='CLOSE'+[guid]::NewGuid().ToString('N')
    try {
        $excel.Visible=$true;$book=$excel.Workbooks.Add();$book.Worksheets.Item(1).Cells.Item(1,1).Value2=$canary
        $book.SaveAs((Join-Path $runRoot 'admin-close-operator.xlsm'),52);$book.Activate()
        SelectTarget $Fixture
        if(-not [bool](Run 'invSys.Core.xlam' 'TestAdminClose.EnableDefaultPolicy')){throw 'Close policy fixture unavailable; not behavioral RED.'}
        $configHash=(Get-FileHash $Fixture.Config).Hash;$otherHash=(Get-FileHash $Other.Config).Hash
        $old=[string](Run 'invSys.Core.xlam' 'TestAdminClose.Definition' @(31));$definition=[string](Run 'invSys.Core.xlam' 'TestAdminClose.Definition' @(32))
        Check 'AdminCloseActivity.Catalog.OlderExcludes' ($old -ceq '')
        $d=if($definition){$definition|ConvertFrom-Json}else{$null}
        Check 'AdminCloseActivity.Catalog.DiscoveredControl' ($null -ne $d -and $d.ControlId -ceq 'ADMIN_SETTINGS_CLOSE' -and $d.OwnerId -ceq 'ADMIN_SETTINGS_UI' -and $d.Class -ceq 'Command' -and $d.Role -ceq 'Admin' -and $d.Capability -ceq 'ADMIN_MAINT')
        # Seed prior history through an already-registered real Settings handler,
        # so preservation cannot pass merely because there are no older records.
        [void](Act 'OpenSettings')
        if(-not [bool](Run 'invSys.Admin.xlam' 'TestD5Commands.SaveSettings' @('BatchSize','500'))){throw 'Prior-activity Settings fixture unavailable; not behavioral RED.'}
        [void](Act 'CloseSettings')
        $configHash=(Get-FileHash $Fixture.Config).Hash
        $pins=@{};foreach($file in Files){$pins[$file]=(Get-FileHash $file).Hash}
        if($pins.Count -lt 2){throw 'Nonempty prior activity required; not behavioral RED.'}
        foreach($mode in @('Button','Native')){
            $before=@(Files);Open
            Check ('AdminCloseActivity.'+$mode+'.OpeningIsNotClose') (@(Files).Count -eq $before.Count)
            if($CaptureEvidence){CaptureOwnedFormByCaptionEvidence 'invSys Settings' ('admin-close-'+$mode.ToLowerInvariant()+'.png')}
            if($mode -ceq 'Button'){[void](Act 'AdminCloseButtonForTest')}elseif(-not (Close-AdminNativeForTest)){throw 'Native Close did not dismiss the form; inspect owner before observation RED.'}
            if(-not (Dismissed)){throw 'Existing dismissal failed; not observation RED.'}
            Check ('AdminCloseActivity.'+$mode+'.ActualDismissal') $true
            Pair $before ('AdminCloseActivity.'+$mode)
            [void](Act 'CloseSettings')
        }
        foreach($mode in @('InternalUnload','ChangedTarget','SignedOut')){
            SelectTarget $Fixture;Open;$before=@(Files);$otherBefore=@(Get-Slice4beActivityFiles $Other)
            if($mode -ceq 'InternalUnload'){[void](Act 'CloseSettings')}
            else{
                if($mode -ceq 'ChangedTarget'){SelectTarget $Other}else{[void](Run 'invSys.Core.xlam' 'modAuth.SignOut')}
                [void](Act 'AdminCloseButtonForTest')
            }
            Check ('AdminCloseActivity.'+$mode+'.StillDismisses') (Dismissed)
            Check ('AdminCloseActivity.'+$mode+'.NoInventedOrRetargetedRecords') (@(Files).Count -eq $before.Count -and @(Get-Slice4beActivityFiles $Other).Count -eq $otherBefore.Count)
            [void](Act 'CloseSettings')
        }
        Check 'AdminCloseActivity.SavedConfigPreserved' ($configHash -ceq (Get-FileHash $Fixture.Config).Hash -and $otherHash -ceq (Get-FileHash $Other.Config).Hash)
        Check 'AdminCloseActivity.OperatorValuesPreserved' ([string]$book.Worksheets.Item(1).Cells.Item(1,1).Value2 -ceq $canary -and $book.Saved)
        Check 'AdminCloseActivity.OlderRecordsImmutable' (@($pins.Keys|Where-Object{(Get-FileHash $_).Hash -cne $pins[$_]}).Count -eq 0)
        SelectTarget $Fixture
        if(-not [bool](Run 'invSys.Core.xlam' 'TestAdminClose.EnableDefaultPolicy' @($false))){throw 'Disabled policy fixture unavailable.'}
        $disabledHash=(Get-FileHash $Fixture.Config).Hash
        foreach($mode in @('Button','Native')){
            Open;$before=@(Files)
            if($mode -ceq 'Button'){[void](Act 'AdminCloseButtonForTest')}elseif(-not (Close-AdminNativeForTest)){throw 'Native disabled dismissal failed.'}
            Check ('AdminCloseActivity.Disabled.'+$mode+'.StillDismisses') (Dismissed)
            Check ('AdminCloseActivity.Disabled.'+$mode+'.NoOptionalRecords') (@(Files).Count -eq $before.Count)
            [void](Act 'CloseSettings')
        }
        Check 'AdminCloseActivity.Disabled.PolicyPreserved' ($disabledHash -ceq (Get-FileHash $Fixture.Config).Hash)
        if(-not [bool](Run 'invSys.Core.xlam' 'TestAdminClose.EnableDefaultPolicy')){throw 'Restored policy fixture unavailable.'}
        $authPath=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Auth.xlsb')
        $authBytes=[IO.File]::ReadAllBytes($authPath);$authHash=(Get-FileHash $authPath).Hash
        try {
            Open;$context=[string](Run 'invSys.Core.xlam' 'modActivity.CaptureContext')
            $auth=$excel.Workbooks.Open($authPath,0,$false)
            try {
                $caps=Table $auth 'tblCapabilities';$changed=0
                foreach($row in $caps.ListRows){if($row.Range.Cells.Item(1,$caps.ListColumns.Item('UserId').Index).Value2 -ceq 'config-admin' -and $row.Range.Cells.Item(1,$caps.ListColumns.Item('Capability').Index).Value2 -ceq 'ADMIN_MAINT'){$row.Range.Cells.Item(1,$caps.ListColumns.Item('Status').Index).Value2='Inactive';$changed++}}
                $auth.Save()
            } finally {$auth.Close($false);$auth=$null}
            [void](Run 'invSys.Core.xlam' 'modAuth.ReloadAuth' @($Fixture.Warehouse))
            $lost=$changed -gt 0 -and -not [bool](Run 'invSys.Core.xlam' 'modRoleUiAccess.CanCurrentUserPerformCapabilityCached' @('ADMIN_MAINT')) -and $context -cne '' -and [string](Run 'invSys.Core.xlam' 'modActivity.CaptureContext') -ceq $context
            if(-not $lost){throw 'Same-session permission-loss fixture unavailable; not behavioral RED.'}
            Check 'AdminCloseActivity.PermissionLoss.SameCapturedSession' $true
            $before=@(Files);[void](Act 'AdminCloseButtonForTest')
            Check 'AdminCloseActivity.PermissionLoss.StillDismisses' (Dismissed)
            Pair $before 'AdminCloseActivity.PermissionLoss'
            [void](Act 'CloseSettings')
        } finally {
            [void](Act 'CloseSettings');[IO.File]::WriteAllBytes($authPath,$authBytes)
            [void](Run 'invSys.Core.xlam' 'modAuth.ReloadAuth' @($Fixture.Warehouse))
        }
        Check 'AdminCloseActivity.PermissionLoss.AuthRestored' ($authHash -ceq (Get-FileHash $authPath).Hash)
        if(-not [bool](Run 'invSys.Core.xlam' 'TestAdminClose.EnableDefaultPolicy' @($true,'config-admin'))){throw 'User-off policy fixture unavailable.'}
        $userPolicyHash=(Get-FileHash $Fixture.Config).Hash
        foreach($mode in @('Button','Native')){
            Open;$before=@(Files)
            if($mode -ceq 'Button'){[void](Act 'AdminCloseButtonForTest')}elseif(-not (Close-AdminNativeForTest)){throw 'Native user-off dismissal failed.'}
            Check ('AdminCloseActivity.UserOff.'+$mode+'.StillDismisses') (Dismissed)
            Check ('AdminCloseActivity.UserOff.'+$mode+'.NoOptionalRecords') (@(Files).Count -eq $before.Count)
            [void](Act 'CloseSettings')
        }
        Check 'AdminCloseActivity.UserOff.PolicyPreserved' ((Get-FileHash $Fixture.Config).Hash -ceq $userPolicyHash)
        if(-not [bool](Run 'invSys.Core.xlam' 'TestAdminClose.EnableDefaultPolicy')){throw 'Restored enabled policy fixture unavailable.'}
        $closeAction={param($mode)
            if($mode -ceq 'Native'){return (Close-AdminNativeForTest)}
            [void](Act 'AdminCloseButtonForTest');return (Dismissed)
        }
        $configBytes=[IO.File]::ReadAllBytes($Fixture.Config);$configHash=(Get-FileHash $Fixture.Config).Hash
        try {
            $cfg=$excel.Workbooks.Open($Fixture.Config,0,$false)
            try {
                (Table $cfg 'tblEventTrackingPolicies').ListColumns.Item('CatalogVersion').DataBodyRange.Value2=31.0
                $controls=Table $cfg 'tblEventTrackingControls'
                for($i=$controls.ListRows.Count;$i -ge 1;$i--){if($controls.ListRows.Item($i).Range.Cells.Item(1,$controls.ListColumns.Item('ControlId').Index).Value2 -ceq 'ADMIN_SETTINGS_CLOSE'){$controls.ListRows.Item($i).Delete()}}
                $cfg.Save()
            } finally {$cfg.Close($false);$cfg=$null}
            Check 'AdminCloseActivity.OlderPolicy.ExistingControlReadable' ([bool](Run 'invSys.Core.xlam' 'TestAdminClose.ExistingPolicyReadable'))
            Open;$before=@(Files)
            $notice=Invoke-ProductionCloseWithNotice -Mode Button -CloseForTest $closeAction
            Check 'AdminCloseActivity.OlderPolicy.DismissalAndNotice' ($notice.Dismissed -and (Dismissed) -and $notice.TrackingNoticeVisible)
            Check 'AdminCloseActivity.OlderPolicy.NoImplicitEnablement' (@(Files).Count -eq $before.Count)
            [void](Act 'CloseSettings')
        } finally {[IO.File]::WriteAllBytes($Fixture.Config,$configBytes)}
        Check 'AdminCloseActivity.OlderPolicy.ConfigRestored' ((Get-FileHash $Fixture.Config).Hash -ceq $configHash)
        $blocked=Join-Path (Join-Path $Fixture.Root 'Training\Activity') $Fixture.Warehouse;$held=$blocked+'-admin-close-held'
        foreach($path in @($blocked,$held)){if(-not [IO.Path]::GetFullPath($path).StartsWith([IO.Path]::GetFullPath($Fixture.Root).TrimEnd('\')+'\',[StringComparison]::OrdinalIgnoreCase)){throw 'Tracking fault escaped generated fixture.'}}
        if(Test-Path -LiteralPath $held){throw 'Preserve existing held fixture.'}
        foreach($mode in @('Button','Native')){
            Open;$before=@(Files);$moved=$false
            try {
                if(Test-Path -LiteralPath $blocked){Move-Item -LiteralPath $blocked -Destination $held;$moved=$true}
                [IO.File]::WriteAllText($blocked,'Blocked disposable Admin Close activity path')
                $notice=Invoke-ProductionCloseWithNotice -Mode $mode -CloseForTest $closeAction
                Check ('AdminCloseActivity.Unavailable.'+$mode+'.StillDismisses') ($notice.Dismissed -and (Dismissed))
                Check ('AdminCloseActivity.Unavailable.'+$mode+'.ActualTrackingNotice') $notice.TrackingNoticeVisible
            } finally {
                if(Test-Path -LiteralPath $blocked -PathType Leaf){Remove-Item -LiteralPath $blocked}
                if($moved){Move-Item -LiteralPath $held -Destination $blocked}
            }
            Check ('AdminCloseActivity.Unavailable.'+$mode+'.NoFallbackActivity') (@(Files).Count -eq $before.Count)
            [void](Act 'CloseSettings')
        }
        Check 'AdminCloseActivity.Unavailable.ConfigPreserved' ((Get-FileHash $Fixture.Config).Hash -ceq $configHash)
        Check 'AdminCloseActivity.Unavailable.OlderRecordsImmutable' (@($pins.Keys|Where-Object{(Get-FileHash $_).Hash -cne $pins[$_]}).Count -eq 0)
    } finally {
        [void](Act 'CloseSettings');if($null -ne $book){$book.Close($false)};SelectTarget $Fixture
    }
}
