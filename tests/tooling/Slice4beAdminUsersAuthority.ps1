# Direct Core checks supplement the actual packaged-handler authority RED/GREEN.
# Only disposable generated Auth files are changed; never emit account payloads.
function Install-AdminUsersAuthorityProbe {
    $probe=$packages['invSys.Core.xlam'].VBProject.VBComponents.Add(1);$probe.Name='TestUsersAuthority'
    $probe.CodeModule.AddFromString(@'
Option Explicit
Private mContext As String
Private mReport As String
Private mSaveGate As TestUsersSaveGate
Public Function Capture() As Boolean
    mContext = modAuthUserCommands.CaptureContext()
    Capture = (Len(mContext) > 0)
End Function
Public Function Invoke(ByVal action As String, ByVal path As String, ByVal warehouse As String, Optional ByVal user As String = "config-reader", Optional ByVal credential As String = "") As Boolean
    mReport = ""
    Select Case action
        Case "Save": Invoke = modAuthUserCommands.SaveUser(mContext, path, user, "Updated fixture reader", credential, warehouse, "S1", 2, mReport)
        Case "Deactivate": Invoke = modAuthUserCommands.DeactivateUser(mContext, path, user, mReport)
        Case "Delete": Invoke = modAuthUserCommands.DeleteUser(mContext, path, user, mReport)
        Case Else: Err.Raise 5, , "Unknown fixture command"
    End Select
End Function
Public Function HasRefusal() As Boolean
    HasRefusal = (Len(mReport) > 0)
End Function
Public Function PersistenceRefusal() As Boolean
    PersistenceRefusal = (mReport = "Users & Roles could not persist the command. Inspect the current Auth state before retrying.")
End Function
Public Sub CancelSave(ByVal path As String)
    Set mSaveGate = New TestUsersSaveGate
    mSaveGate.Bind Application, path
End Sub
Public Sub ReleaseSaveGate()
    Set mSaveGate = Nothing
End Sub
Public Function CancelSeen() As Boolean
    If Not mSaveGate Is Nothing Then CancelSeen = mSaveGate.Seen
End Function
'@)
    $gate=$packages['invSys.Core.xlam'].VBProject.VBComponents.Add(2);$gate.Name='TestUsersSaveGate'
    $gate.CodeModule.AddFromString(@'
Option Explicit
Private WithEvents mExcel As Application
Private mPath As String
Public Seen As Boolean
Public Sub Bind(ByVal host As Application, ByVal path As String)
    Set mExcel = host: mPath = path
End Sub
Private Sub mExcel_WorkbookBeforeSave(ByVal Wb As Workbook, ByVal SaveAsUI As Boolean, Cancel As Boolean)
    If StrComp(Wb.FullName, mPath, vbTextCompare) = 0 Then
        Seen = True: Cancel = True
    End If
End Sub
'@)
    $form=$packages['invSys.Admin.xlam'].VBProject.VBComponents.Item('frmCreateDeleteUser').CodeModule
    $start=$form.ProcStartLine('MsgBox',0);$count=$form.ProcCountLines('MsgBox',0)
    $source=$form.Lines($start,$count)
    $anchor='    MsgBox = vbYes'
    if([regex]::Matches($source,[regex]::Escape($anchor)).Count -ne 1){throw 'Confirmation seam missing.'}
    $source=$source.Replace($anchor,"    TestAdminAuthRefresh.OnConfirmation`r`n"+$anchor)
    $form.DeleteLines($start,$count);$form.InsertLines($start,$source)
    $start=$form.ProcStartLine('ShowStatusForm',0);$count=$form.ProcCountLines('ShowStatusForm',0)
    $source=$form.Lines($start,$count)
    $anchor='(?im)^[ \t]*mLblStatus\.Caption[ \t]*=[ \t]*messageText[ \t]*\r?$'
    if([regex]::Matches($source,$anchor).Count -ne 1){throw 'Status observation anchor missing.'}
    $source=[regex]::Replace($source,$anchor,('$0'+"`r`n    TestAdminAuthRefresh.ObserveAuthorityStatus messageText"))
    $form.DeleteLines($start,$count);$form.InsertLines($start,$source)
    $admin=$packages['invSys.Admin.xlam'].VBProject.VBComponents.Item('TestAdminAuthRefresh').CodeModule
    $admin.InsertLines($admin.CountOfDeclarationLines+1,"Private mSignOutOnConfirmation As Boolean`r`nPrivate mAuthorityRefusal As Boolean")
    $admin.AddFromString(@'
Public Sub ArmConfirmation()
    mSignOutOnConfirmation = True
End Sub
Public Sub OnConfirmation()
    If mSignOutOnConfirmation Then
        mSignOutOnConfirmation = False
        modAuth.SignOut
    End If
End Sub
Public Function AuthorityRefusalForTest() As Boolean
    AuthorityRefusalForTest = mAuthorityRefusal
End Function
Public Sub ObserveAuthorityStatus(ByVal message As String)
    mAuthorityRefusal = (message = "The invSys session or warehouse changed. Reopen Users & Roles.")
End Sub
'@)
}

function Test-AdminUsersAuthority($Fixture,$Other) {
    $authPath=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Auth.xlsb')
    $otherPath=Join-Path $Other.Root ($Other.Warehouse+'.invSys.Auth.xlsb')
    $owned=[IO.Path]::GetFullPath($runRoot).TrimEnd('\')+'\'
    foreach($path in @($authPath,$otherPath)){if(-not [IO.Path]::GetFullPath($path).StartsWith($owned,[StringComparison]::OrdinalIgnoreCase)){throw 'Authority gate requires generated fixtures.'}}
    $original=[IO.File]::ReadAllBytes($authPath);$originalHash=Get-AuthReadFixtureHash $authPath
    $otherHash=Get-AuthReadFixtureHash $otherPath;$auth=$null
    try {
        foreach($case in @('SignedOut','Reader','OtherWarehouse','Reauthenticated','OtherPath','LostCapability','DirtyBorrowed','ReadOnlyBorrowed','Missing','Invalid')){
            foreach($action in @('Save','Deactivate','Delete')){
                [IO.File]::WriteAllBytes($authPath,$original)
                SelectTarget $Fixture
                if(-not [bool](Run 'invSys.Core.xlam' 'TestUsersAuthority.Capture')){throw 'Initial captured context missing.'}
                $selectedPath=$authPath
                switch($case){
                    'SignedOut' {[void](Run 'invSys.Core.xlam' 'modAuth.SignOut')}
                    'Reader' {SelectTarget $Fixture 'config-reader'}
                    'OtherWarehouse' {SelectTarget $Other}
                    'Reauthenticated' {SelectTarget $Fixture}
                    'OtherPath' {$selectedPath=$otherPath}
                    'LostCapability' {
                        $auth=$excel.Workbooks.Open($authPath,0,$false);$caps=Table $auth 'tblCapabilities'
                        foreach($row in $caps.ListRows){if($row.Range.Cells.Item(1,$caps.ListColumns.Item('UserId').Index).Value2 -ceq 'config-admin' -and $row.Range.Cells.Item(1,$caps.ListColumns.Item('Capability').Index).Value2 -ceq 'ADMIN_MAINT'){$row.Range.Cells.Item(1,$caps.ListColumns.Item('Status').Index).Value2='INACTIVE'}}
                        $auth.Save();$auth.Close($false);$auth=$null
                        Check ('UsersCore.'+$case+'.'+$action+'.CachedGrantStillPresent') ([bool](Run 'invSys.Core.xlam' 'modAuth.CanPerform' @('ADMIN_MAINT','config-admin',$Fixture.Warehouse,'S1')))
                    }
                    'DirtyBorrowed' {
                        $auth=$excel.Workbooks.Open($authPath,0,$false)
                        if(-not [bool](Run 'invSys.Admin.xlam' 'TestAdminAuthRefresh.BorrowedCanaryForTest' @($auth.Name,$true))){throw 'Dirty canary missing.'}
                    }
                    'ReadOnlyBorrowed' {$auth=$excel.Workbooks.Open($authPath,0,$true)}
                    'Missing' {[IO.File]::Delete($authPath)}
                    'Invalid' {$auth=$excel.Workbooks.Open($authPath,0,$false);(Table $auth 'tblUsers').ListColumns.Item('Status').Delete();$auth.Save();$auth.Close($false);$auth=$null}
                }
                $before=if(Test-Path -LiteralPath $authPath){Get-AuthReadFixtureHash $authPath}else{''}
                [void](Run 'invSys.Core.xlam' 'modAuth.ResetAdminRefreshEnsureForTest')
                $ok=[bool](Run 'invSys.Core.xlam' 'TestUsersAuthority.Invoke' @($action,$selectedPath,$Fixture.Warehouse))
                Check ('UsersCore.'+$case+'.'+$action+'.DeniedWithReason') (-not $ok -and [bool](Run 'invSys.Core.xlam' 'TestUsersAuthority.HasRefusal'))
                $after=if(Test-Path -LiteralPath $authPath){Get-AuthReadFixtureHash $authPath}else{''}
                Check ('UsersCore.'+$case+'.'+$action+'.AuthoritiesPreserved') ($before -ceq $after -and (Get-AuthReadFixtureHash $otherPath) -ceq $otherHash -and [long](Run 'invSys.Core.xlam' 'modAuth.AdminRefreshEnsureCountForTest') -eq 0)
                if($null -ne $auth){
                    $retained=$auth.FullName -ieq $authPath
                    if($case -ceq 'DirtyBorrowed'){$retained=$retained -and [bool](Run 'invSys.Admin.xlam' 'TestAdminAuthRefresh.BorrowedCanaryForTest' @($auth.Name,$false))}
                    else{$retained=$retained -and $auth.ReadOnly -and $auth.Saved}
                    Check ('UsersCore.'+$case+'.'+$action+'.BorrowedStatePreserved') $retained
                    $auth.Close($false);$auth=$null
                }else{Check ('UsersCore.'+$case+'.'+$action+'.NoOwnedWorkbookLeaked') (@($excel.Workbooks|Where-Object {$_.FullName -ieq $authPath}).Count -eq 0)}
            }
        }
        [IO.File]::WriteAllBytes($authPath,$original);SelectTarget $Fixture
        if(-not [bool](Run 'invSys.Admin.xlam' 'TestAdminAuthRefresh.LoadForTest' @($authPath))){throw 'Delete confirmation fixture missing.'}
        [void](Run 'invSys.Admin.xlam' 'TestAdminAuthRefresh.ArmConfirmation')
        [void](Run 'invSys.Admin.xlam' 'TestAdminAuthRefresh.WriteReaderForTest' @('Delete'))
        Check 'UsersCore.Delete.ConfirmationSignOutRefusedByActualHandler' (-not [bool](Run 'invSys.Core.xlam' 'modAuth.IsSignedIn') -and [bool](Run 'invSys.Admin.xlam' 'TestAdminAuthRefresh.AuthorityRefusalForTest') -and (Get-AuthReadFixtureHash $authPath) -ceq $originalHash)
        if($CaptureEvidence){
            [void](Run 'invSys.Admin.xlam' 'TestAdminAuthRefresh.SetVisibleForTest' @($true))
            CaptureOwnedFormByCaptionEvidence 'Users & Roles' 'admin-users-context-refusal.png'
            [void](Run 'invSys.Admin.xlam' 'TestAdminAuthRefresh.SetVisibleForTest' @($false))
        }
        [void](Run 'invSys.Admin.xlam' 'TestAdminAuthRefresh.CloseForTest')
        foreach($action in @('Save','Deactivate','Delete')){
            [IO.File]::WriteAllBytes($authPath,$original);SelectTarget $Fixture
            [void](Run 'invSys.Core.xlam' 'TestUsersAuthority.Capture')
            $auth=$excel.Workbooks.Open($authPath,0,$false)
            [void](Run 'invSys.Core.xlam' 'TestUsersAuthority.CancelSave' @($authPath))
            $wasEnabled=$excel.EnableEvents;$excel.EnableEvents=$true
            try{$ok=[bool](Run 'invSys.Core.xlam' 'TestUsersAuthority.Invoke' @($action,$authPath,$Fixture.Warehouse))}finally{$excel.EnableEvents=$wasEnabled}
            Check ('UsersCore.CancelledSave.'+$action+'.TruthfulFailure') (-not $ok -and [bool](Run 'invSys.Core.xlam' 'TestUsersAuthority.CancelSeen') -and [bool](Run 'invSys.Core.xlam' 'TestUsersAuthority.PersistenceRefusal'))
            Check ('UsersCore.CancelledSave.'+$action+'.BorrowedDirtyAndDiskUnchanged') (-not $auth.Saved -and $auth.FullName -ieq $authPath -and (Get-AuthReadFixtureHash $authPath) -ceq $originalHash)
            [void](Run 'invSys.Core.xlam' 'TestUsersAuthority.ReleaseSaveGate')
            $auth.Close($false);$auth=$null
        }
        [IO.File]::WriteAllBytes($authPath,$original);SelectTarget $Fixture
        [void](Run 'invSys.Core.xlam' 'TestUsersAuthority.Capture')
        $ok=[bool](Run 'invSys.Core.xlam' 'TestUsersAuthority.Invoke' @('Save',$authPath,$Fixture.Warehouse,'new-fixture-user'))
        Check 'UsersCore.NewUser.MissingCredentialNoOrphanRow' (-not $ok -and (Get-AuthReadFixtureHash $authPath) -ceq $originalHash)
    } finally {
        [void](Run 'invSys.Core.xlam' 'TestUsersAuthority.ReleaseSaveGate')
        [void](Run 'invSys.Admin.xlam' 'TestAdminAuthRefresh.CloseForTest')
        if($null -ne $auth){$auth.Close($false)}
        [IO.File]::WriteAllBytes($authPath,$original)
    }
    SelectTarget $Fixture
    Check 'UsersCore.FixturesRestored' ((Get-AuthReadFixtureHash $authPath) -ceq $originalHash -and (Get-AuthReadFixtureHash $otherPath) -ceq $otherHash)
    Test-AdminUsersDataPreservation $Fixture
}

function Test-AdminUsersDataPreservation($Fixture) {
    $authPath=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Auth.xlsb')
    $owned=[IO.Path]::GetFullPath($runRoot).TrimEnd('\')+'\'
    if(-not [IO.Path]::GetFullPath($authPath).StartsWith($owned,[StringComparison]::OrdinalIgnoreCase)){throw 'Data test requires generated authority.'}
    $original=[IO.File]::ReadAllBytes($authPath);$originalHash=Get-AuthReadFixtureHash $authPath;$auth=$null
    try {
        SelectTarget $Fixture
        if(-not [bool](Run 'invSys.Core.xlam' 'TestUsersAuthority.Capture')){throw 'Captured authority missing.'}
        $auth=$excel.Workbooks.Open($authPath,0,$false)
        $users=Table $auth 'tblUsers';$caps=Table $auth 'tblCapabilities'
        $userExtra=$users.ListColumns.Add(1);$userExtra.Name='Operator_User_Canary';$userExtra.DataBodyRange.Value2='generated user annotation'
        $capExtra=$caps.ListColumns.Add(1);$capExtra.Name='Operator_Cap_Canary';$capExtra.DataBodyRange.Value2='generated capability annotation'
        $oldCapRows=$caps.ListRows.Count
        $reader=@($users.ListRows|Where-Object{$_.Range.Cells.Item(1,$users.ListColumns.Item('UserId').Index).Value2 -ceq 'config-reader'})
        if($reader.Count -ne 1){throw 'Unique generated reader required.'}
        $row=$reader[0].Range
        foreach($date in @(@('ValidFrom','2000-01-01'),@('ValidTo','2099-12-31'))){$cell=$row.Cells.Item(1,$users.ListColumns.Item($date[0]).Index);$cell.NumberFormat='@';$cell.Value2=$date[1]}
        $priorCredential=$row.Cells.Item(1,$users.ListColumns.Item('PinHash').Index).Value2
        $auth.Save()
        $ok=[bool](Run 'invSys.Core.xlam' 'TestUsersAuthority.Invoke' @('Save',$authPath,$Fixture.Warehouse))
        Check 'UsersCore.Columns.UpdateUsesNamesAndRetainsBorrowedBook' ($ok -and $auth.Saved -and $auth.FullName -ieq $authPath -and $users.ListColumns.Item('UserId').Index -gt 1 -and $row.Cells.Item(1,$users.ListColumns.Item('DisplayName').Index).Value2 -ceq 'Updated fixture reader')
        Check 'UsersCore.Columns.UnknownUserDatesAndBlankCredentialPreserved' ($row.Cells.Item(1,1).Value2 -ceq 'generated user annotation' -and $row.Cells.Item(1,$users.ListColumns.Item('ValidFrom').Index).Value2 -ceq '2000-01-01' -and $row.Cells.Item(1,$users.ListColumns.Item('ValidTo').Index).Value2 -ceq '2099-12-31' -and $row.Cells.Item(1,$users.ListColumns.Item('PinHash').Index).Value2 -ceq $priorCredential)
        $canaries=$true;for($i=1;$i -le $oldCapRows;$i++){$canaries=$canaries -and $caps.DataBodyRange.Cells.Item($i,1).Value2 -ceq 'generated capability annotation'}
        Check 'UsersCore.Columns.ExistingCapabilityAnnotationsPreserved' $canaries
        $ok=[bool](Run 'invSys.Core.xlam' 'TestUsersAuthority.Invoke' @('Deactivate',$authPath,$Fixture.Warehouse))
        Check 'UsersCore.Columns.DeactivateOnlyManagedStatus' ($ok -and $auth.Saved -and $row.Cells.Item(1,$users.ListColumns.Item('Status').Index).Value2 -ceq 'INACTIVE' -and $row.Cells.Item(1,1).Value2 -ceq 'generated user annotation')
        $ok=[bool](Run 'invSys.Core.xlam' 'TestUsersAuthority.Invoke' @('Delete',$authPath,$Fixture.Warehouse))
        $left=@($users.ListRows|Where-Object{$_.Range.Cells.Item(1,$users.ListColumns.Item('UserId').Index).Value2 -ceq 'config-reader'})
        $leftCaps=@($caps.ListRows|Where-Object{$_.Range.Cells.Item(1,$caps.ListColumns.Item('UserId').Index).Value2 -ceq 'config-reader'})
        Check 'UsersCore.Columns.DeleteExactUserAndCapabilities' ($ok -and $auth.Saved -and $left.Count -eq 0 -and $leftCaps.Count -eq 0)
        $canaries=$true;foreach($item in $users.ListRows){$canaries=$canaries -and $item.Range.Cells.Item(1,1).Value2 -ceq 'generated user annotation'}
        foreach($item in $caps.ListRows){$canaries=$canaries -and $item.Range.Cells.Item(1,1).Value2 -ceq 'generated capability annotation'}
        Check 'UsersCore.Columns.OtherUsersAndCapabilitiesRetainAnnotations' $canaries
        $auth.Close($false);$auth=$null
        $credential=[guid]::NewGuid().ToString('N')
        $ok=[bool](Run 'invSys.Core.xlam' 'TestUsersAuthority.Invoke' @('Save',$authPath,$Fixture.Warehouse,'new-fixture-user',$credential))
        Check 'UsersCore.NewUser.OwnedAuthoritySavedAndClosed' ($ok -and @($excel.Workbooks|Where-Object{$_.FullName -ieq $authPath}).Count -eq 0)
        $auth=$excel.Workbooks.Open($authPath,0,$true);$users=Table $auth 'tblUsers';$caps=Table $auth 'tblCapabilities'
        $created=@($users.ListRows|Where-Object{$_.Range.Cells.Item(1,$users.ListColumns.Item('UserId').Index).Value2 -ceq 'new-fixture-user'})
        $grants=@($caps.ListRows|Where-Object{$_.Range.Cells.Item(1,$caps.ListColumns.Item('UserId').Index).Value2 -ceq 'new-fixture-user'})
        Check 'UsersCore.NewUser.FreshSavedUserAndExactCapabilityScope' ($created.Count -eq 1 -and $grants.Count -eq 1 -and $grants[0].Range.Cells.Item(1,$caps.ListColumns.Item('Capability').Index).Value2 -ceq 'RECEIVE_POST' -and $grants[0].Range.Cells.Item(1,$caps.ListColumns.Item('WarehouseId').Index).Value2 -ceq $Fixture.Warehouse -and $grants[0].Range.Cells.Item(1,$caps.ListColumns.Item('StationId').Index).Value2 -ceq 'S1')
    } finally {
        if($null -ne $auth){$auth.Close($false)}
        [IO.File]::WriteAllBytes($authPath,$original)
    }
    SelectTarget $Fixture
    Check 'UsersCore.Columns.FixtureRestored' ((Get-AuthReadFixtureHash $authPath) -ceq $originalHash)
}
