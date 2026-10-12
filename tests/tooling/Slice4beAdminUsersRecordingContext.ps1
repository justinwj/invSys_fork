# Context changes occur only at optional recording entry. Business handlers and
# Core authority remain real; generated Auth/activity contents never enter reports.
function Install-AdminUsersRecordingContextProbe {
    $core=$packages['invSys.Core.xlam'].VBProject
    $probe=$core.VBComponents.Add(1);$probe.Name='TestUsersRecordingContext'
    $probe.CodeModule.AddFromString(@'
Option Explicit
Private mControl As String, mMode As String, mRoot As String, mSecret As String
Private mOriginal As String, mWarehouse As String, mSeen As Boolean, mObserved As Boolean
Public Function Arm(ByVal controlId As String, ByVal mode As String, ByVal root As String, ByVal secret As String) As Boolean
    Dim target As WarehouseTarget
    mControl = controlId: mMode = mode: mRoot = root: mSecret = secret
    mSeen = False: mObserved = False: mOriginal = "": mWarehouse = ""
    If controlId = "" Then Exit Function
    mOriginal = modActivity.CaptureContext()
    Set target = modNasConnection.GetCurrentTarget()
    If target Is Nothing Then Exit Function
    mWarehouse = target.WarehouseId
    Arm = (mOriginal <> "")
End Function
Public Sub AtEntry(ByVal controlId As String)
    Dim status As String, target As WarehouseTarget
    If mControl = "" Or controlId <> mControl Or mSeen Then Exit Sub
    mSeen = True
    Select Case mMode
        Case "SignOut"
            modAuth.SignOut
            mObserved = (Not modAuth.IsSignedIn() And modActivity.CaptureContext() = "")
        Case "SwitchActor"
            status = modAuth.SignInCurrentTargetForAutomation("config-reader", mSecret, "")
            mObserved = (Left$(status, 3) = "OK|" And modAuth.GetCurrentUserId() = "config-reader" And modActivity.CaptureContext() <> mOriginal)
        Case "SwitchWarehouse"
            modRuntimeWorkbooks.SetCoreDataRootOverride mRoot
            status = modNasConnection.SelectWarehouseTargetForAutomation(mRoot, mRoot, "S1", False)
            If Left$(status, 3) <> "OK|" Then Exit Sub
            modNasConnection.SetCurrentTargetPathsForTest "\\fixture-host\config-command", mRoot
            status = modAuth.SignInCurrentTargetForAutomation("config-admin", mSecret, "")
            Set target = modNasConnection.GetCurrentTarget()
            If target Is Nothing Then Exit Sub
            mObserved = (Left$(status, 3) = "OK|" And target.WarehouseId <> mWarehouse And modActivity.CaptureContext() <> mOriginal)
    End Select
End Sub
Public Function Observed() As Boolean
    Observed = mSeen And mObserved
End Function
'@)
    $activity=$core.VBComponents.Item('modActivity').CodeModule
    $start=$activity.ProcStartLine('BeginAction',0);$count=$activity.ProcCountLines('BeginAction',0)
    $hits=@(for($line=$start;$line -lt $start+$count;$line++){if($activity.Lines($line,1).Trim() -ieq 'notice = ""'){$line}})
    if($hits.Count -ne 1){throw 'Users context observation boundary unavailable; not product RED.'}
    $activity.InsertLines($hits[0]+1,'    TestUsersRecordingContext.AtEntry controlId')
}

function Test-AdminUsersRecordingContext($Fixture,$Other) {
    $authPath=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Auth.xlsb')
    $otherPath=Join-Path $Other.Root ($Other.Warehouse+'.invSys.Auth.xlsb')
    $inventory=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Data.Inventory.xlsb')
    $owned=[IO.Path]::GetFullPath($runRoot).TrimEnd('\')+'\'
    foreach($path in @($authPath,$otherPath)){if(-not [IO.Path]::GetFullPath($path).StartsWith($owned,[StringComparison]::OrdinalIgnoreCase)){throw 'Generated authority fixtures required.'}}
    $originalHash=Get-AuthReadFixtureHash $authPath;$otherHash=Get-AuthReadFixtureHash $otherPath;$inventoryHash=Get-AuthReadFixtureHash $inventory
    try{
        foreach($mode in @('SignOut','SwitchActor','SwitchWarehouse')){
            foreach($action in @('Save','Deactivate','Delete')){
                SelectTarget $Fixture
                if(-not [bool](Run 'invSys.Core.xlam' 'TestUsersRecording.Policy')){throw 'Enabled policy unavailable.'}
                if(-not [bool](Run 'invSys.Admin.xlam' 'TestAdminAuthRefresh.LoadForTest' @($authPath))){throw 'Captured form unavailable.'}
                $before=@{};foreach($path in @((Get-Slice4beActivityFiles $Fixture),(Get-Slice4beActivityFiles $Other))|ForEach-Object{$_}){if($null -ne $path){$before[$path]=(Get-FileHash -LiteralPath $path).Hash}}
                $id='ADMIN_USER_'+$action.ToUpperInvariant();$target=if($mode -ceq 'SwitchWarehouse'){$Other}else{$Fixture}
                [void](Run 'invSys.Core.xlam' 'TestUsersRecordingStorage.Arm' @($id,''))
                if(-not [bool](Run 'invSys.Core.xlam' 'TestUsersRecordingContext.Arm' @($id,$mode,$target.Root,$target.Secret))){throw 'Original signed-in recording context unavailable.'}
                try{
                    [void](Run 'invSys.Admin.xlam' 'TestAdminAuthRefresh.WriteReaderForTest' @($action))
                    $changed=[bool](Run 'invSys.Core.xlam' 'TestUsersRecordingContext.Observed')
                    $counts=([string](Run 'invSys.Core.xlam' 'TestUsersRecordingStorage.Observations')) -split '\|'
                }finally{
                    [void](Run 'invSys.Core.xlam' 'TestUsersRecordingContext.Arm' @('','','',''))
                    [void](Run 'invSys.Core.xlam' 'TestUsersRecordingStorage.Arm' @('',''))
                }
                $prefix='UsersContext.'+$mode+'.'+$action
                Check ($prefix+'.ContextChangeReached') $changed
                if(-not $changed){throw 'Context change fixture failed; not product RED.'}
                Check ($prefix+'.OneCoreCallNoPersistence') ($counts.Count -eq 5 -and $counts[3] -ceq '1' -and $counts[4] -ceq '0')
                Check ($prefix+'.OriginalAuthorityRefused') ([bool](Run 'invSys.Admin.xlam' 'TestAdminAuthRefresh.AuthorityRefusalForTest'))
                Check ($prefix+'.OneFixedTrackingNotice') ([int](Run 'invSys.Admin.xlam' 'TestAdminAuthRefresh.TrackingNotices') -eq 1)
                Check ($prefix+'.AuthoritiesAndInventoryUnchanged') ((Get-AuthReadFixtureHash $authPath) -ceq $originalHash -and (Get-AuthReadFixtureHash $otherPath) -ceq $otherHash -and (Get-AuthReadFixtureHash $inventory) -ceq $inventoryHash)
                $after=@((Get-Slice4beActivityFiles $Fixture),(Get-Slice4beActivityFiles $Other))|ForEach-Object{$_}|Where-Object{$null -ne $_}
                Check ($prefix+'.NoStaleOrLaterAttribution') (@($after).Count -eq $before.Count -and @($after|Where-Object{-not $before.ContainsKey($_)}).Count -eq 0)
                $immutable=$true;foreach($path in $before.Keys){$immutable=$immutable -and (Test-Path -LiteralPath $path) -and (Get-FileHash -LiteralPath $path).Hash -ceq $before[$path]}
                Check ($prefix+'.PriorActivityImmutable') $immutable
                Check ($prefix+'.NoOwnedAuthLeaked') (@($excel.Workbooks|Where-Object{$_.FullName -ieq $authPath -or $_.FullName -ieq $otherPath}).Count -eq 0)
                [void](Run 'invSys.Admin.xlam' 'TestAdminAuthRefresh.CloseForTest')
            }
        }
    }finally{
        [void](Run 'invSys.Core.xlam' 'TestUsersRecordingContext.Arm' @('','','',''))
        [void](Run 'invSys.Core.xlam' 'TestUsersRecordingStorage.Arm' @('',''))
        [void](Run 'invSys.Admin.xlam' 'TestAdminAuthRefresh.CloseForTest')
        SelectTarget $Fixture
    }
    Check 'UsersContext.AuthoritiesPreservedAtEnd' ((Get-AuthReadFixtureHash $authPath) -ceq $originalHash -and (Get-AuthReadFixtureHash $otherPath) -ceq $otherHash)
}
