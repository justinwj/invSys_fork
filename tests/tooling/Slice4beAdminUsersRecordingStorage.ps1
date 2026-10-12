# Generated Auth fixtures only. Real activity Open failures; observe actual Core
# entry/persistence without replacing the command, save, or error handler.
function Install-AdminUsersRecordingStorageProbe {
    $core=$packages['invSys.Core.xlam'].VBProject
    $probe=$core.VBComponents.Add(1);$probe.Name='TestUsersRecordingStorage'
    $probe.CodeModule.AddFromString(@'
Option Explicit
Private mControl As String, mOutcome As String
Private mInjected As Long, mFailures As Long, mError As Long
Private mCommands As Long, mSaves As Long
Public Sub Arm(ByVal controlId As String, ByVal outcome As String)
    mControl = controlId: mOutcome = outcome
    mInjected = 0: mFailures = 0: mError = 0: mCommands = 0: mSaves = 0
End Sub
Public Sub AtWrite(ByVal body As String, ByRef pending As String)
    Dim record As Object
    If mControl = "" Then Exit Sub
    Set record = modTrainingJson.DecodeObject(body)
    If record("ControlId") <> mControl Or record("OutcomeCode") <> mOutcome Then Exit Sub
    mInjected = mInjected + 1
    pending = pending & "\unavailable.pending"
End Sub
Public Sub AtFailure(ByVal number As Long)
    If mControl = "" Or mInjected = 0 Then Exit Sub
    mFailures = mFailures + 1: mError = number
End Sub
Public Sub AtCommand(ByVal command As String)
    If mControl = "ADMIN_USER_" & command Then mCommands = mCommands + 1
End Sub
Public Sub AtSave()
    If mControl <> "" Then mSaves = mSaves + 1
End Sub
Public Function Observations() As String
    Observations = CStr(mInjected) & "|" & CStr(mFailures) & "|" & CStr(mError <> 0) & "|" & CStr(mCommands) & "|" & CStr(mSaves)
End Function
'@)
    foreach($case in @(
        @('modActivityStore','Append','    Open pending For Binary Access Write Lock Read Write As #file','    TestUsersRecordingStorage.AtWrite body, pending',$true),
        @('modActivityStore','Append','Failed:','    TestUsersRecordingStorage.AtFailure Err.Number',$false),
        @('modAuthUserCommands','Execute','    outcome = "REJECTED"','    TestUsersRecordingStorage.AtCommand command',$false),
        @('modAuthUserCommands','Execute','    wb.Save','    TestUsersRecordingStorage.AtSave',$true)
    )){
        $module=$core.VBComponents.Item($case[0]).CodeModule
        $start=$module.ProcStartLine($case[1],0);$count=$module.ProcCountLines($case[1],0)
        $hits=@(for($line=$start;$line -lt $start+$count;$line++){if($module.Lines($line,1).Trim() -ieq $case[2].Trim()){$line}})
        if($hits.Count -ne 1){throw ('Users storage observation boundary unavailable: '+$case[0]+'.'+$case[1]+'; not product RED.')}
        $position=$hits[0];if(-not $case[4]){$position++}
        $module.InsertLines($position,$case[3])
    }
}

function Test-AdminUsersRecordingStorage($Fixture,$Other) {
    $authPath=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Auth.xlsb')
    $otherPath=Join-Path $Other.Root ($Other.Warehouse+'.invSys.Auth.xlsb')
    $inventory=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Data.Inventory.xlsb')
    $owned=[IO.Path]::GetFullPath($runRoot).TrimEnd('\')+'\'
    if(-not [IO.Path]::GetFullPath($authPath).StartsWith($owned,[StringComparison]::OrdinalIgnoreCase)){throw 'Disposable generated Auth required.'}
    $original=[IO.File]::ReadAllBytes($authPath);$originalHash=Get-AuthReadFixtureHash $authPath
    $otherHash=Get-AuthReadFixtureHash $otherPath;$inventoryHash=Get-AuthReadFixtureHash $inventory
    $auth=$null
    try{
        foreach($mode in @('Attempt','Outcome')){
            foreach($action in @('Save','Deactivate','Delete')){
                [IO.File]::WriteAllBytes($authPath,$original);SelectTarget $Fixture
                if(-not [bool](Run 'invSys.Core.xlam' 'TestUsersRecording.Policy')){throw 'Enabled recording policy unavailable.'}
                if(-not [bool](Run 'invSys.Admin.xlam' 'TestAdminAuthRefresh.LoadForTest' @($authPath))){throw 'Captured Users form unavailable.'}
                $before=@{};foreach($path in @(Get-Slice4beActivityFiles $Fixture)){$before[$path]=(Get-FileHash -LiteralPath $path).Hash}
                $id='ADMIN_USER_'+$action.ToUpperInvariant();$wanted=if($mode -ceq 'Attempt'){'REQUESTED'}else{'SAVED'}
                [void](Run 'invSys.Core.xlam' 'TestUsersRecordingStorage.Arm' @($id,$wanted))
                try{
                    [void](Run 'invSys.Admin.xlam' 'TestAdminAuthRefresh.WriteReaderForTest' @($action))
                    $observations=([string](Run 'invSys.Core.xlam' 'TestUsersRecordingStorage.Observations')) -split '\|'
                }finally{[void](Run 'invSys.Core.xlam' 'TestUsersRecordingStorage.Arm' @('',''))}
                $prefix='UsersStorage.'+$mode+'.'+$action
                $fault=$observations.Count -eq 5 -and $observations[0] -ceq '1' -and $observations[1] -ceq '1' -and $observations[2] -ceq 'True'
                Check ($prefix+'.RealOpenFailureReached') $fault
                if(-not $fault){throw 'Real storage failure fixture was not reached; not behavioral RED.'}
                Check ($prefix+'.CoreCommandOnce') ($observations[3] -ceq '1')
                Check ($prefix+'.PersistenceOnce') ($observations[4] -ceq '1')
                Check ($prefix+'.OneFixedTrackingNotice') ([int](Run 'invSys.Admin.xlam' 'TestAdminAuthRefresh.TrackingNotices') -eq 1)
                Check ($prefix+'.OwnedAuthClosed') (@($excel.Workbooks|Where-Object {$_.FullName -ieq $authPath}).Count -eq 0)
                $auth=$excel.Workbooks.Open($authPath,0,$true)
                $users=Table $auth 'tblUsers';$caps=Table $auth 'tblCapabilities'
                $rows=@($users.ListRows|Where-Object {$_.Range.Cells.Item(1,$users.ListColumns.Item('UserId').Index).Value2 -ceq 'config-reader'})
                $state=$auth.ReadOnly -and $auth.Saved -and (Get-AuthReadFixtureHash $authPath) -cne $originalHash
                if($action -ceq 'Delete'){
                    $state=$state -and $rows.Count -eq 0 -and @($caps.ListRows|Where-Object {$_.Range.Cells.Item(1,$caps.ListColumns.Item('UserId').Index).Value2 -ceq 'config-reader'}).Count -eq 0
                }else{
                    $status=if($action -ceq 'Save'){'ACTIVE'}else{'INACTIVE'}
                    $state=$state -and $rows.Count -eq 1 -and $rows[0].Range.Cells.Item(1,$users.ListColumns.Item('Status').Index).Value2 -ceq $status
                    if($action -ceq 'Save'){$state=$state -and $rows[0].Range.Cells.Item(1,$users.ListColumns.Item('DisplayName').Index).Value2 -ceq 'Updated fixture reader'}
                }
                Check ($prefix+'.FreshSavedBusinessState') $state
                $auth.Close($false);$auth=$null
                $records=@(Get-Slice4beActivityFiles $Fixture|Where-Object {-not $before.ContainsKey($_)}|ForEach-Object{Get-Content -LiteralPath $_ -Raw|ConvertFrom-Json})
                $honest=if($mode -ceq 'Attempt'){$records.Count -eq 0}else{$records.Count -eq 1 -and $records[0].ControlId -ceq $id -and $records[0].OutcomeCode -ceq 'REQUESTED'}
                Check ($prefix+'.HonestPartialEvidence') $honest
                $safe=$true
                foreach($record in $records){
                    $wire=$record|ConvertTo-Json -Depth 20 -Compress
                    foreach($value in @($Fixture.Root,$Fixture.Secret,(CredentialHash $Fixture.Secret),'config-reader','Updated fixture reader','PinHash','unavailable.pending')){$safe=$safe -and -not $wire.Contains($value)}
                    $safe=$safe -and $record.UserId -ceq 'config-admin' -and $record.WarehouseId -ceq $Fixture.Warehouse -and @($record.SourceEventRefs).Count -eq 0 -and -not [bool](Run 'invSys.Core.xlam' 'TestUsersRecording.Concludes' @($wire))
                }
                Check ($prefix+'.NoPrivatePayloadOrFalseConclusion') $safe
                $immutable=$true;foreach($path in $before.Keys){$immutable=$immutable -and (Test-Path -LiteralPath $path) -and (Get-FileHash -LiteralPath $path).Hash -ceq $before[$path]}
                Check ($prefix+'.PriorActivityImmutable') $immutable
                Check ($prefix+'.OtherAuthorityAndInventoryPreserved') ((Get-AuthReadFixtureHash $otherPath) -ceq $otherHash -and (Get-AuthReadFixtureHash $inventory) -ceq $inventoryHash)
                $activityRoot=Join-Path $Fixture.Root 'Training/Activity'
                Check ($prefix+'.NoPendingArtifacts') (@(Get-ChildItem -LiteralPath $activityRoot -Recurse -File -Filter '*.pending').Count -eq 0)
                [void](Run 'invSys.Admin.xlam' 'TestAdminAuthRefresh.CloseForTest')
            }
        }
    }finally{
        [void](Run 'invSys.Core.xlam' 'TestUsersRecordingStorage.Arm' @('',''))
        [void](Run 'invSys.Admin.xlam' 'TestAdminAuthRefresh.CloseForTest')
        if($null -ne $auth){$auth.Close($false)}
        [IO.File]::WriteAllBytes($authPath,$original);SelectTarget $Fixture
    }
    Check 'UsersStorage.AuthFixtureRestored' ((Get-AuthReadFixtureHash $authPath) -ceq $originalHash)
}
