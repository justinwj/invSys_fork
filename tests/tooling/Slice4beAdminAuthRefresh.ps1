# Approved D8-A through the real form handler. All authority fixtures are generated;
# observations contain counts/booleans only, never users, credentials or cell values.
function Install-AdminAuthRefreshProbe {
    $core=$packages['invSys.Core.xlam'].VBProject.VBComponents.Item('modAuth').CodeModule
    $core.InsertLines($core.CountOfDeclarationLines+1,'Private mAdminRefreshEnsureCalls As Long')
    $start=$core.ProcStartLine('EnsureAuthSchema',0);$count=$core.ProcCountLines('EnsureAuthSchema',0)
    $source=$core.Lines($start,$count)
    $anchor='    On Error GoTo FailEnsure'
    if([regex]::Matches($source,[regex]::Escape($anchor)).Count -ne 1){throw 'Auth schema observation anchor unavailable.'}
    $source=$source.Replace($anchor,$anchor+"`r`n    mAdminRefreshEnsureCalls = mAdminRefreshEnsureCalls + 1")
    $core.DeleteLines($start,$count);$core.InsertLines($start,$source)
    $core.AddFromString(@'
Public Sub ResetAdminRefreshEnsureForTest()
    mAdminRefreshEnsureCalls = 0
End Sub
Public Function AdminRefreshEnsureCountForTest() As Long
    AdminRefreshEnsureCountForTest = mAdminRefreshEnsureCalls
End Function
'@)
    $form=$packages['invSys.Admin.xlam'].VBProject.VBComponents.Item('frmCreateDeleteUser').CodeModule
    foreach($procedure in @('RefreshUsersForm','LoadSelectedUserForm')){
    $start=$form.ProcStartLine($procedure,0);$count=$form.ProcCountLines($procedure,0)
    $source=$form.Lines($start,$count)
    $anchor='    Set wb = OpenSelectedAuthWorkbookForm(openedTransient, report)'
    if([regex]::Matches($source,[regex]::Escape($anchor)).Count -ne 1){throw 'Users Refresh observation anchor unavailable.'}
    $source=$source.Replace($anchor,$anchor+"`r`n    TestAdminAuthRefresh.Observe wb")
    $form.DeleteLines($start,$count);$form.InsertLines($start,$source)
    }
    $form.AddFromString(@'
Public Function RefreshUsersForTest() As Long
    mBtnRefreshUsers_Click
    RefreshUsersForTest = mLstUsers.ListCount
End Function
Public Function SelectedPathForTest() As String
    SelectedPathForTest = CStr(mTxtAuthPath.Value)
End Function
Public Sub SetVisibleForTest(ByVal shown As Boolean)
    If shown Then
        mTxtNasUser.Value = "": mTxtNasPassword.Value = "": mTxtPin.Value = ""
        Me.Show vbModeless
    Else
        Me.Hide
    End If
End Sub
Public Function SelectReaderForTest() As Long
    Dim i As Long
    For i = 0 To mLstUsers.ListCount - 1
        If CStr(mLstUsers.List(i, 0)) = "config-reader" Then
            mLstUsers.ListIndex = i
            mLstUsers_Click
            SelectReaderForTest = Abs(CLng(CStr(mTxtUserId.Value) = "config-reader" And CStr(mTxtPin.Value) = ""))
            Exit Function
        End If
    Next i
End Function
Public Sub WriteReaderForTest(ByVal action As String)
    mTxtUserId.Value = "config-reader"
    mTxtDisplayName.Value = "Updated fixture reader"
    mTxtPin.Value = ""
    mChkRoamingStation.Value = False
    mTxtStationId.Value = "S1"
    mChkReceivePost.Value = True
    Select Case action
        Case "Save": mBtnSave_Click
        Case "Deactivate": mBtnDeactivate_Click
        Case "Delete": mBtnDelete_Click
        Case Else: Err.Raise 5, , "Unknown fixture write action"
    End Select
End Sub
Private Function MsgBox(ByVal prompt As String, Optional ByVal buttons As VbMsgBoxStyle = vbOKOnly, Optional ByVal title As String = "") As VbMsgBoxResult
    If prompt <> "Delete user 'config-reader' and all assigned capabilities?" Or title <> "Users & Roles" Then Err.Raise 5, , "Unexpected fixture confirmation"
    MsgBox = vbYes
End Function
'@)
    $probe=$packages['invSys.Admin.xlam'].VBProject.VBComponents.Add(1);$probe.Name='TestAdminAuthRefresh'
    $probe.CodeModule.AddFromString(@'
Option Explicit
Private mSeen As Boolean
Private mReadOnly As Boolean
Private mSaved As Boolean
Public Function LoadForTest(ByVal expectedPath As String) As Boolean
    Load frmCreateDeleteUser
    LoadForTest = (StrComp(frmCreateDeleteUser.SelectedPathForTest(), expectedPath, vbTextCompare) = 0)
End Function
Public Sub Observe(ByVal wb As Workbook)
    If wb Is Nothing Then Exit Sub
    mSeen = True: mReadOnly = wb.ReadOnly: mSaved = wb.Saved
End Sub
Public Function RefreshForTest(Optional ByVal selection As Boolean = False) As String
    Dim rows As Long
    mSeen = False: mReadOnly = False: mSaved = False
    modAuth.ResetAdminRefreshEnsureForTest
    If selection Then
        rows = frmCreateDeleteUser.SelectReaderForTest()
    Else
        rows = frmCreateDeleteUser.RefreshUsersForTest()
    End If
    RefreshForTest = CStr(rows) & "|" & CStr(mSeen) & "|" & CStr(mReadOnly) & "|" & CStr(mSaved) & "|" & CStr(modAuth.AdminRefreshEnsureCountForTest())
End Function
Public Sub CloseForTest()
    Unload frmCreateDeleteUser
End Sub
Public Sub SetVisibleForTest(ByVal shown As Boolean)
    frmCreateDeleteUser.SetVisibleForTest shown
End Sub
Public Sub WriteReaderForTest(ByVal action As String)
    frmCreateDeleteUser.WriteReaderForTest action
End Sub
Public Function BorrowedCanaryForTest(ByVal bookName As String, ByVal stage As Boolean) As Boolean
    Dim wb As Workbook, ws As Worksheet, lo As ListObject, extra As ListColumn
    Set wb = Application.Workbooks(bookName)
    For Each ws In wb.Worksheets
        For Each lo In ws.ListObjects
            If lo.Name = "tblUsers" Then
                If stage Then
                    Set extra = lo.ListColumns.Add
                    extra.Name = "OperatorAuthReadCanary"
                    extra.DataBodyRange.Value2 = "unsaved fixture annotation"
                Else
                    Set extra = lo.ListColumns("OperatorAuthReadCanary")
                End If
                BorrowedCanaryForTest = (Not wb.Saved And CStr(extra.DataBodyRange.Cells(1, 1).Value2) = "unsaved fixture annotation")
                Exit Function
            End If
        Next lo
    Next ws
End Function
'@)
}

function Test-AdminAuthRefresh($Fixture,$Other) {
    SelectTarget $Fixture
    $authPath=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Auth.xlsb')
    $otherPath=Join-Path $Other.Root ($Other.Warehouse+'.invSys.Auth.xlsb')
    $owned=[IO.Path]::GetFullPath($runRoot).TrimEnd('\')+'\'
    foreach($path in @($authPath,$otherPath)){
        if(-not [IO.Path]::GetFullPath($path).StartsWith($owned,[StringComparison]::OrdinalIgnoreCase)){throw 'Admin Auth fixture escaped generated root.'}
    }
    $original=[IO.File]::ReadAllBytes($authPath)
    $originalHash=Get-AuthReadFixtureHash $authPath
    $otherHash=Get-AuthReadFixtureHash $otherPath
    $auth=$null;$otherBook=$null
    function InvokeRefresh([bool]$Selection=$false) {
        $fields=([string](Run 'invSys.Admin.xlam' 'TestAdminAuthRefresh.RefreshForTest' @($Selection))) -split '\|'
        if($fields.Count -ne 5){throw 'Invalid typed Refresh observation.'}
        [pscustomobject]@{Rows=[int]$fields[0];Seen=[bool]::Parse($fields[1]);ReadOnly=[bool]::Parse($fields[2]);Saved=[bool]::Parse($fields[3]);EnsureCalls=[int]$fields[4]}
    }
    function AssertClosed([string]$Case) {
        $open=@($excel.Workbooks|Where-Object{$_.FullName -ieq $authPath})
        Check ('AdminAuthRefresh.'+$Case+'.OwnedWorkbookClosed') ($open.Count -eq 0)
    }
    try {
        $auth=$excel.Workbooks.Open($authPath,0,$false)
        $users=Table $auth 'tblUsers';$expectedRows=$users.ListRows.Count
        $auth.Close($false);$auth=$null
        $loaded=[bool](Run 'invSys.Admin.xlam' 'TestAdminAuthRefresh.LoadForTest' @($authPath))
        Check 'AdminAuthRefresh.Fixture.ExactGeneratedSelection' $loaded
        if(-not $loaded){throw 'Users form did not select the generated fixture.'}
        foreach($case in @('Healthy','MissingStatus','MissingCapabilities')){
            [IO.File]::WriteAllBytes($authPath,$original)
            if($case -ne 'Healthy'){
                $auth=$excel.Workbooks.Open($authPath,0,$false)
                if($case -eq 'MissingStatus'){(Table $auth 'tblUsers').ListColumns.Item('Status').Delete()}
                else{(Table $auth 'tblCapabilities').Delete()}
                $auth.Save();$auth.Close($false);$auth=$null
            }
            $before=Get-AuthReadFixtureHash $authPath
            $r=InvokeRefresh
            if($case -eq 'Healthy'){
                Check 'AdminAuthRefresh.Healthy.ReadsExistingUsers' ($r.Seen -and $r.Rows -eq $expectedRows -and $expectedRows -gt 0)
                Check 'AdminAuthRefresh.Healthy.OpensReadOnly' ($r.Seen -and $r.ReadOnly)
                Check 'AdminAuthRefresh.Healthy.DoesNotDirtyAuthority' ($r.Seen -and $r.Saved)
            }else{
                Check ('AdminAuthRefresh.'+$case+'.FailsClosed') ($r.Rows -eq 0 -and -not $r.Seen)
            }
            Check ('AdminAuthRefresh.'+$case+'.NoSchemaMaintenance') ($r.EnsureCalls -eq 0)
            Check ('AdminAuthRefresh.'+$case+'.FileBytesPreserved') ((Get-AuthReadFixtureHash $authPath) -ceq $before)
            AssertClosed $case
            if($CaptureEvidence -and $case -in @('Healthy','MissingStatus')){
                # Only generated fixture data is visible; remembered NAS fields
                # and credential controls are blanked in the disposable form.
                [void](Run 'invSys.Admin.xlam' 'TestAdminAuthRefresh.SetVisibleForTest' @($true))
                CaptureOwnedFormByCaptionEvidence 'Users & Roles' ('admin-auth-'+$case.ToLower()+'.png')
                [void](Run 'invSys.Admin.xlam' 'TestAdminAuthRefresh.SetVisibleForTest' @($false))
            }
            if($case -eq 'Healthy'){
                $r=InvokeRefresh $true
                Check 'AdminAuthRefresh.Selection.ReadsUserWithoutCredentialDisclosure' ($r.Rows -eq 1 -and $r.Seen)
                Check 'AdminAuthRefresh.Selection.ReadOnlyWithoutSchemaMaintenance' ($r.ReadOnly -and $r.Saved -and $r.EnsureCalls -eq 0)
                Check 'AdminAuthRefresh.Selection.FileBytesPreserved' ((Get-AuthReadFixtureHash $authPath) -ceq $before)
                AssertClosed 'Selection'
            }
        }
        [IO.File]::WriteAllBytes($authPath,$original)
        $auth=$excel.Workbooks.Open($authPath,0,$false)
        $staged=[bool](Run 'invSys.Admin.xlam' 'TestAdminAuthRefresh.BorrowedCanaryForTest' @($auth.Name,$true))
        if(-not $staged){throw 'Borrowed Auth canary staging failed before Refresh.'}
        $before=Get-AuthReadFixtureHash $authPath
        $dirty=-not [bool]$auth.Saved
        $r=InvokeRefresh
        Check 'AdminAuthRefresh.Borrowed.ReadsWithoutSchemaMaintenance' ($r.Rows -eq $expectedRows -and $r.EnsureCalls -eq 0)
        Check 'AdminAuthRefresh.Borrowed.DirtyWorkbookAndCustomValuePreserved' ($dirty -and -not $auth.Saved -and $auth.FullName -ieq $authPath -and [bool](Run 'invSys.Admin.xlam' 'TestAdminAuthRefresh.BorrowedCanaryForTest' @($auth.Name,$false)))
        Check 'AdminAuthRefresh.Borrowed.FileBytesPreserved' ((Get-AuthReadFixtureHash $authPath) -ceq $before)
        $auth.Close($false);$auth=$null
        [IO.File]::Delete($authPath)
        $otherBook=$excel.Workbooks.Open($otherPath,0,$true)
        $r=InvokeRefresh
        Check 'AdminAuthRefresh.Missing.FailsClosedWithOtherWarehouseOpen' ($r.Rows -eq 0 -and -not $r.Seen -and $r.EnsureCalls -eq 0)
        Check 'AdminAuthRefresh.Missing.DoesNotCreateAuthority' (-not(Test-Path -LiteralPath $authPath))
        Check 'AdminAuthRefresh.OtherWarehouse.PreservedAndStillOpen' ($otherBook.FullName -ieq $otherPath -and $otherBook.Saved -and (Get-AuthReadFixtureHash $otherPath) -ceq $otherHash)
        $otherBook.Close($false);$otherBook=$null
        [IO.File]::WriteAllBytes($authPath,$original)
        foreach($action in @('Save','Deactivate','Delete')){
            [void](Run 'invSys.Admin.xlam' 'TestAdminAuthRefresh.WriteReaderForTest' @($action))
            $auth=$excel.Workbooks.Open($authPath,0,$true)
            $users=Table $auth 'tblUsers';$caps=Table $auth 'tblCapabilities'
            $rows=@($users.ListRows|Where-Object{$_.Range.Cells.Item(1,$users.ListColumns.Item('UserId').Index).Value2 -ceq 'config-reader'})
            $ok=$false
            if($action -eq 'Save'){$ok=$rows.Count -eq 1 -and $rows[0].Range.Cells.Item(1,$users.ListColumns.Item('DisplayName').Index).Value2 -ceq 'Updated fixture reader'}
            elseif($action -eq 'Deactivate'){$ok=$rows.Count -eq 1 -and $rows[0].Range.Cells.Item(1,$users.ListColumns.Item('Status').Index).Value2 -ceq 'INACTIVE'}
            else{$ok=$rows.Count -eq 0 -and @($caps.ListRows|Where-Object{$_.Range.Cells.Item(1,$caps.ListColumns.Item('UserId').Index).Value2 -ceq 'config-reader'}).Count -eq 0}
            Check ('AdminAuthRefresh.ExplicitWrite.'+$action+'.PersistedThroughRealHandler') $ok
            $auth.Close($false);$auth=$null
        }
    } finally {
        [void](Run 'invSys.Admin.xlam' 'TestAdminAuthRefresh.CloseForTest')
        if($null -ne $auth){$auth.Close($false)}
        if($null -ne $otherBook){$otherBook.Close($false)}
        [IO.File]::WriteAllBytes($authPath,$original)
    }
    SelectTarget $Fixture
    Check 'AdminAuthRefresh.Fixture.RestoredForCommandRegressions' ((Get-AuthReadFixtureHash $authPath) -ceq $originalHash)
}
