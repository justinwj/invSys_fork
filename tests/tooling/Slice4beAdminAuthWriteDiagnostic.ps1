# Observes existing form write handlers after authorization-context changes.
# Counts/booleans only; no credentials, account values or paths in observations.
# Diagnostic completion is not authorization or recording acceptance. The opt-in
# D8-USERS-01 assertions require the nine observed unauthorized writes to stop.
function Test-AdminAuthWriteDiagnostic($Fixture,$Other) {
    $owned=[IO.Path]::GetFullPath($runRoot).TrimEnd('\')+'\'
    $authPath=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Auth.xlsb')
    $otherPath=Join-Path $Other.Root ($Other.Warehouse+'.invSys.Auth.xlsb')
    foreach($path in @($authPath,$otherPath)){
        if(-not [IO.Path]::GetFullPath($path).StartsWith($owned,[StringComparison]::OrdinalIgnoreCase)){throw 'Write diagnostic requires generated Auth fixtures.'}
    }
    $original=[IO.File]::ReadAllBytes($authPath)
    $originalHash=Get-AuthReadFixtureHash $authPath
    $otherHash=Get-AuthReadFixtureHash $otherPath
    $observations=[Collections.Generic.List[object]]::new()
    $auth=$null
    try {
        foreach($state in @('Admin','SignedOut','Reader','OtherWarehouseAdmin')){
            foreach($action in @('Save','Deactivate','Delete')){
                [IO.File]::WriteAllBytes($authPath,$original)
                SelectTarget $Fixture
                $initialAdmin=[bool](Run 'invSys.Core.xlam' 'modAuth.CanPerform' @('ADMIN_MAINT','config-admin',$Fixture.Warehouse,'S1'))
                if(-not $initialAdmin){throw 'Initial Admin context missing; diagnostic invalid.'}
                $loaded=[bool](Run 'invSys.Admin.xlam' 'TestAdminAuthRefresh.LoadForTest' @($authPath))
                if(-not $loaded){throw 'Original generated Auth selection missing.'}
                switch($state){
                    'SignedOut' {[void](Run 'invSys.Core.xlam' 'modAuth.SignOut')}
                    'Reader' {SelectTarget $Fixture 'config-reader'}
                    'OtherWarehouseAdmin' {SelectTarget $Other}
                }
                $signedIn=[bool](Run 'invSys.Core.xlam' 'modAuth.IsSignedIn')
                $currentUser=[string](Run 'invSys.Core.xlam' 'modAuth.GetCurrentUserId')
                $currentWarehouse=[string](Run 'invSys.Core.xlam' 'modNasConnection.GetCurrentTargetWarehouseId')
                $originalTarget=$currentWarehouse -ceq $Fixture.Warehouse
                $canAdminCurrent=[bool](Run 'invSys.Core.xlam' 'modAuth.CanPerform' @('ADMIN_MAINT',$currentUser,$currentWarehouse,'S1'))
                $canAdminOriginal=[bool](Run 'invSys.Core.xlam' 'modAuth.CanPerform' @('ADMIN_MAINT',$currentUser,$Fixture.Warehouse,'S1'))
                $selectedOriginal=([string](Run 'invSys.Admin.xlam' 'TestAdminAuthRefresh.SelectedPathForTest')) -ieq $authPath
                if(-not $selectedOriginal){throw 'Retained form selection changed before write.'}
                $stateValid=switch($state){
                    'Admin' {$signedIn -and $originalTarget -and $canAdminCurrent}
                    'SignedOut' {-not $signedIn -and -not $canAdminCurrent}
                    'Reader' {$signedIn -and $originalTarget -and -not $canAdminCurrent}
                    'OtherWarehouseAdmin' {$signedIn -and -not $originalTarget -and $canAdminCurrent}
                }
                Check ('AdminAuthWriteDiagnostic.'+$state+'.'+$action+'.PreconditionsVerified') $stateValid
                if(-not $stateValid){throw 'Requested authorization state not established.'}
                [void](Run 'invSys.Admin.xlam' 'TestAdminAuthRefresh.WriteReaderForTest' @($action))
                $changed=(Get-AuthReadFixtureHash $authPath) -cne $originalHash
                $auth=$excel.Workbooks.Open($authPath,0,$true)
                $users=Table $auth 'tblUsers';$caps=Table $auth 'tblCapabilities'
                $rows=@($users.ListRows|Where-Object{$_.Range.Cells.Item(1,$users.ListColumns.Item('UserId').Index).Value2 -ceq 'config-reader'})
                $persisted=switch($action){
                    'Save' {$rows.Count -eq 1 -and $rows[0].Range.Cells.Item(1,$users.ListColumns.Item('DisplayName').Index).Value2 -ceq 'Updated fixture reader'}
                    'Deactivate' {$rows.Count -eq 1 -and $rows[0].Range.Cells.Item(1,$users.ListColumns.Item('Status').Index).Value2 -ceq 'INACTIVE'}
                    'Delete' {$rows.Count -eq 0 -and @($caps.ListRows|Where-Object{$_.Range.Cells.Item(1,$caps.ListColumns.Item('UserId').Index).Value2 -ceq 'config-reader'}).Count -eq 0}
                }
                $auth.Close($false);$auth=$null
                $otherPreserved=(Get-AuthReadFixtureHash $otherPath) -ceq $otherHash
                Check ('AdminAuthWriteDiagnostic.'+$state+'.'+$action+'.OtherAuthorityPreserved') $otherPreserved
                if($state -ceq 'Admin'){Check ('AdminAuthWriteDiagnostic.Admin.'+$action+'.BaselinePersists') ($changed -and $persisted)}
                if($RequireAdminAuthWriteAuthority -and $state -cne 'Admin'){
                    Check ('AdminAuthWriteAuthority.'+$state+'.'+$action+'.DeniedWithoutMutation') (-not $changed -and -not $persisted)
                }
                $observations.Add([pscustomobject]@{Case=$state;Action=$action;OpenedAsAdmin=$initialAdmin;SignedInBeforeWrite=$signedIn;CurrentTargetMatchesOriginal=$originalTarget;CurrentTargetAdmin=$canAdminCurrent;OriginalTargetAdmin=$canAdminOriginal;FormRetainedOriginalTarget=$selectedOriginal;AuthBytesChanged=$changed;RequestedMutationPersisted=[bool]$persisted;OtherAuthorityPreserved=$otherPreserved})
                [void](Run 'invSys.Admin.xlam' 'TestAdminAuthRefresh.CloseForTest')
            }
        }
    } finally {
        if($null -ne $auth){$auth.Close($false)}
        [void](Run 'invSys.Admin.xlam' 'TestAdminAuthRefresh.CloseForTest')
        [IO.File]::WriteAllBytes($authPath,$original)
    }
    SelectTarget $Fixture
    Check 'AdminAuthWriteDiagnostic.FixturesRestored' ((Get-AuthReadFixtureHash $authPath) -ceq $originalHash -and (Get-AuthReadFixtureHash $otherPath) -ceq $otherHash)
    [pscustomobject]@{DiagnosticOnly=(-not $RequireAdminAuthWriteAuthority);AuthorizationAssertionsEnabled=[bool]$RequireAdminAuthWriteAuthority;AuthorizationAccepted=$false;Cases=$observations.ToArray()}|ConvertTo-Json -Depth 4|Set-Content (Join-Path $reportRoot 'admin-auth-write-diagnostic.json')
}
