# Test-only fault at the actual local preference write; no operational registry faults.
. (Join-Path $PSScriptRoot 'Slice4beAdminRootFeedbackNative.ps1')
. (Join-Path $PSScriptRoot 'Slice4beAdminRootRecording.ps1')
function Install-AdminWarehouseRootWriteProbe {
    $project=$packages['invSys.Admin.xlam'].VBProject
    $probe=$project.VBComponents.Item('TestDirectoryColumns').CodeModule
    $probe.InsertLines($probe.CountOfDeclarationLines+1,@'
Public RootAction As Boolean
Public RootNative As Boolean
Private mRootInput As String, mRootFault As Boolean, mRootSuccess As Long
Private mRootWrites As Long, mRootFailures As Long
Private mRootReadMode As String, mRootReads As Long, mRootDirectoryFault As Boolean
Private mRootFeedback As String, mRootStyle As Long, mRootMessages As Long
Private mRootTracking As Long
'@)
    $probe.AddFromString(@'
Public Sub SetRootPreference(ByVal value As String)
    VBA.SaveSetting "invSys", "Admin", "WarehouseScanRoots", value
End Sub
Public Function RootPreferenceEquals(ByVal expected As String) As Boolean
    RootPreferenceEquals = (StrComp(VBA.GetSetting("invSys", "Admin", "WarehouseScanRoots", ""), expected, vbBinaryCompare) = 0)
End Function
Public Sub NativeRootFeedback(ByVal enabled As Boolean)
    RootNative = enabled
End Sub
Public Function RootTrackingNotices() As Long
    RootTrackingNotices = mRootTracking
End Function
Public Function RootReadback(ByVal actual As String) As String
    RootReadback = actual
    If Not RootAction Or mRootWrites = 0 Then Exit Function
    mRootReads = mRootReads + 1
    If mRootReadMode = "ReadFailure" Then Err.Raise 70, , "Generated preference-read refusal"
    If mRootReadMode = "ReadMismatch" Then RootReadback = actual & "|generated-mismatch"
End Function
Public Function RootDirectoryFails(ByRef report As String) As Boolean
    If RootAction And mRootDirectoryFault Then
        report = "Generated directory refresh failure."
        RootDirectoryFails = True
    End If
End Function
Public Function RootWriteAllowed() As Boolean
    RootWriteAllowed = True
    If Not RootAction Then Exit Function
    mRootWrites = mRootWrites + 1
    If mRootFault Then
        mRootFailures = mRootFailures + 1
        RootWriteAllowed = False
    End If
End Function
Public Function InvokeRoot(ByVal enteredRoot As String, ByVal failWrite As Boolean, ByVal readMode As String, ByVal failDirectory As Boolean) As String
    Dim control As New TestDirectoryRibbon, before As String, after As String
    Dim item As Variant, remembered As Boolean, selected As Boolean
    Dim number As Long, description As String
    On Error GoTo Failed
    mHost.Activate: mHost.Worksheets("Personal_notes").Activate
    If Not (modAdmin.ResolveInteractiveAdminWorkbook() Is mHost) Then Err.Raise 5, , "Root host not bound"
    mRootInput = enteredRoot: mRootFault = failWrite: mRootSuccess = 0
    mRootWrites = 0: mRootFailures = 0: mCalls = 0: mOpened = False
    mRootReadMode = readMode: mRootReads = 0: mRootDirectoryFault = failDirectory
    mRootFeedback = "": mRootStyle = 0: mRootMessages = 0
    mRootTracking = 0
    mSuccess = 0: mOther = 0: mPrompts = 0
    before = VBA.GetSetting("invSys", "Admin", "WarehouseScanRoots", "")
    RootAction = True
    RibbonOnActionAdmin control
    RootAction = False
    after = VBA.GetSetting("invSys", "Admin", "WarehouseScanRoots", "")
    If enteredRoot <> "" Then
        For Each item In Split(after, "|")
            If StrComp(CStr(item), enteredRoot, vbTextCompare) = 0 Then remembered = True
        Next item
    End If
    selected = (Application.ActiveWorkbook Is mHost)
    InvokeRoot = CStr(mCalls) & "|" & CStr(mOpened) & "|" & CStr(mRootSuccess) & "|" & CStr(mPrompts) & "|" & CStr(mRootWrites) & "|" & CStr(mRootFailures) & "|" & CStr(remembered) & "|" & CStr(before = after) & "|" & CStr(selected) & "|" & CStr(mOther)
    InvokeRoot = InvokeRoot & "|" & CStr(mRootReads) & "|" & CStr(mRootMessages) & "|" & CStr(mRootStyle) & "|" & mRootFeedback
    Exit Function
Failed:
    number = Err.Number: description = Err.Description: RootAction = False
    Err.Raise number, , description
End Function
'@)
    $start=$probe.ProcStartLine('ObserveMessage',0);$count=$probe.ProcCountLines('ObserveMessage',0)
    $source=$probe.Lines($start,$count)
    $changed=$source.Replace('ByVal text As String)', 'ByVal text As String, Optional ByVal style As Long = 0)').Replace('    If text = "Warehouse directory refreshed." Then',@'
    If RootAction And Left$(text, 21) = "Tracking unavailable:" Then
        mRootTracking = mRootTracking + 1: ObserveMessage = vbOK: Exit Function
    End If
    If RootAction Then
        mRootMessages = mRootMessages + 1: mRootStyle = style
        mRootFeedback = Replace(Replace(text, vbCrLf, " "), vbLf, " ")
        If RootNative Then ObserveMessage = VBA.MsgBox(text, style, "invSys Admin"): Exit Function
    End If
    If RootAction And text = "Warehouse root remembered and directory refreshed." Then
        mRootSuccess = mRootSuccess + 1
        ObserveMessage = vbOK
        Exit Function
    End If
    If text = "Warehouse directory refreshed." Then
'@.TrimEnd())
    if($changed -ceq $source){throw 'Root feedback observation seam changed'}
    $probe.DeleteLines($start,$count);$probe.InsertLines($start,$changed)
    $start=$probe.ProcStartLine('ObservePrompt',0)
    $probe.InsertLines($start+1,'    If RootAction Then ObservePrompt = mRootInput')
    $ribbon=$project.VBComponents.Item('TestDirectoryRibbon').CodeModule
    $lines=@(for($line=1;$line -le $ribbon.CountOfLines;$line++){if($ribbon.Lines($line,1).Trim() -ceq 'Private Property Get IRibbonControl_Id() As String'){$line}})
    if($lines.Count -ne 1){throw 'Root Ribbon callback selector unavailable'}
    $ribbon.InsertLines($lines[0]+1,'    If TestDirectoryColumns.RootAction Then IRibbonControl_Id = "btnAdminWarehouseRoot": Exit Property')
    $admin=$project.VBComponents.Item('modAdmin').CodeModule
    $start=$admin.ProcStartLine('MsgBox',0);$count=$admin.ProcCountLines('MsgBox',0)
    $source=$admin.Lines($start,$count).Replace('ObserveMessage(Prompt)','ObserveMessage(Prompt, Buttons)')
    $admin.DeleteLines($start,$count);$admin.InsertLines($start,$source)
    $rootOwner=@($project.VBComponents|Where-Object Name -CEQ 'modAdminRootActivity')
    if($rootOwner.Count){$rootOwner[0].CodeModule.AddFromString(@'
Private Function MsgBox(ByVal Prompt As String, Optional ByVal Buttons As VbMsgBoxStyle = vbOKOnly, Optional ByVal Title As String = "") As VbMsgBoxResult
    MsgBox = TestDirectoryColumns.ObserveMessage(Prompt, Buttons)
End Function
Private Function InputBox(ByVal Prompt As String, Optional ByVal Title As String = "", Optional ByVal Default As String = "") As String
    InputBox = TestDirectoryColumns.ObservePrompt()
End Function
'@)}
    $owner=$project.VBComponents.Item('modAdminDirectoryActivity').CodeModule
    $start=$owner.ProcStartLine('OpenView',0);$count=$owner.ProcCountLines('OpenView',0)
    $source=$owner.Lines($start,$count)
    $changed=$source.Replace('    TestDirectoryColumns.OwnerEntered',"    TestDirectoryColumns.OwnerEntered`r`n    If TestDirectoryColumns.RootDirectoryFails(report) Then Exit Function")
    if($source -ceq $changed){throw 'Root directory failure seam unavailable'}
    $owner.DeleteLines($start,$count);$owner.InsertLines($start,$changed)
    $project.VBComponents.Item('modAdminConsole').CodeModule.AddFromString(@'
Private Function GetSetting(ByVal app As String, ByVal section As String, ByVal key As String, Optional ByVal default As String = "") As String
    GetSetting = VBA.GetSetting(app, section, key, default)
    If key = "WarehouseScanRoots" Then GetSetting = TestDirectoryColumns.RootReadback(GetSetting)
End Function
Private Sub SaveSetting(ByVal app As String, ByVal section As String, ByVal key As String, ByVal setting As String)
    If key = "WarehouseScanRoots" Then
        If Not TestDirectoryColumns.RootWriteAllowed() Then Err.Raise 70, , "Generated preference-write refusal"
    End If
    VBA.SaveSetting app, section, key, setting
End Sub
'@)
    Install-AdminRootRecordingProbe
}

function Test-AdminWarehouseRootWrite($Fixture,$Other,$HostBook,$OtherBook) {
    # Keep the strict original file-set proof independent of optional activity writes.
    SelectTarget $Fixture
    if(-not [bool](Run 'invSys.Core.xlam' 'TestRootActivity.SetPolicy' @('Disabled'))){throw 'Owner-only recording suppression prerequisite failed'}
    $before=Pins $Fixture.Root;$otherBefore=Pins $Other.Root
    $faultRoot=Join-Path $runRoot 'unremembered-root'
    New-Item -ItemType Directory -Path $faultRoot|Out-Null
    foreach($case in @('Success','Cancel','WriteFailure','ReadFailure','ReadMismatch','DirectoryFailure','WriteAndDirectoryFailure','ReadAndDirectoryFailure','MismatchAndDirectoryFailure')){
        SelectTarget $Fixture
        $cancel=$case -ceq 'Cancel';$fault=$case -in @('WriteFailure','WriteAndDirectoryFailure')
        $directoryFault=$case -like '*DirectoryFailure'
        $readMode=if($case -in @('ReadFailure','ReadAndDirectoryFailure')){'ReadFailure'}elseif($case -in @('ReadMismatch','MismatchAndDirectoryFailure')){'ReadMismatch'}else{''}
        $entered=if($cancel){''}else{Join-Path $faultRoot $case}
        $facts=([string](Invoke 'InvokeRoot' @($entered,$fault,$readMode,$directoryFault))) -split '\|'
        if($facts.Count -ne 14){throw 'Root callback facts unavailable; not product RED'}
        $prefix='WarehouseRoot.'+$case+'.';$confirmed=-not ($cancel -or $fault -or $readMode)
        $opened=-not ($cancel -or $directoryFault)
        Check ($prefix+'ActualPrompt') ($facts[3] -ceq '1')
        Check ($prefix+'ActualWriteBoundary') ($facts[4] -ceq $(if($cancel){'0'}else{'1'}) -and $facts[5] -ceq $(if($fault){'1'}else{'0'}))
        Check ($prefix+'ExistingDirectoryBehavior') ($facts[0] -ceq $(if($cancel){'0'}else{'1'}) -and $facts[1] -ceq [string]$opened -and $facts[8] -ceq 'True')
        Check ($prefix+'ExactSavedRootResult') ($facts[6] -ceq [string](-not ($cancel -or $fault)) -and $facts[7] -ceq [string]($cancel -or $fault))
        Check ($prefix+'NoFalseRememberedMessage') ($facts[2] -ceq $(if($confirmed -and $opened){'1'}else{'0'}))
        $message=if($cancel){''}elseif($confirmed -and $opened){'Warehouse root remembered and directory refreshed.'}else{
            $result=if($opened){'Warehouse directory refreshed.'}else{'Generated directory refresh failure.'}
            if($confirmed){$result}else{'Warehouse root preference could not be confirmed. '+$result}
        }
        Check ($prefix+'NoUnexpectedFeedback') ($facts[13] -ceq $message -and $facts[11] -ceq $(if($cancel){'0'}else{'1'}) -and $facts[12] -ceq $(if($cancel){'0'}elseif($confirmed -and $opened){'64'}else{'48'}) -and [int](Invoke 'RootTrackingNotices') -eq 0)
        Check ($prefix+'ActualReadbackBoundary') ($facts[10] -ceq $(if($cancel -or $fault){'0'}else{'1'}))
        $preserved=$true;foreach($part in @('Managed','Headers','Text','Formula','Tail','Notes')){$preserved=$preserved -and [bool](Invoke 'Preserved' @($part))}
        Check ($prefix+'CapturedWorkbookAndCustomColumnsPreserved') ($preserved -and [string]$OtherBook.Worksheets.Item(1).Cells.Item(1,1).Value2 -ceq 'Other workbook canary')
        Check ($prefix+'SavedAuthorityPreserved') ((Same $before (Pins $Fixture.Root)) -and (Same $otherBefore (Pins $Other.Root)))
    }
    $ordered=@(1..10|ForEach-Object{Join-Path $faultRoot ('slot'+$_)})
    [void](Invoke 'SetRootPreference' @(($ordered -join '|')))
    $selected=$ordered[2].ToUpperInvariant()
    $facts=([string](Invoke 'InvokeRoot' @(('  '+$selected+'\  '),$false,'',$false))) -split '\|'
    $expected=@($selected)+@($ordered[0],$ordered[1])+$ordered[3..7]
    Check 'WarehouseRoot.NormalizationDeduplicationOrderAndEightRootLimit' ([bool](Invoke 'RootPreferenceEquals' @(($expected -join '|'))) -and $facts[2] -ceq '1')
    $facts=([string](Invoke 'InvokeRoot' @('   ',$false,'',$false))) -split '\|'
    Check 'WarehouseRoot.WhitespaceCancellation' ($facts[0] -ceq '0' -and $facts[4] -ceq '0' -and $facts[7] -ceq 'True' -and $facts[11] -ceq '0')
    if($Phase -ceq 'GREEN'){
        foreach($mode in @('RefreshSucceeded','RefreshFailed')){
            Invoke-AdminRootFeedbackNative $mode (Join-Path $faultRoot $mode)
        }
    }
    if($AdminWarehouseRootRecording){Test-AdminRootRecording $Fixture $Other $HostBook $OtherBook}
}
