# Test-only fault at the actual local preference write; no operational registry faults.
function Install-AdminWarehouseRootWriteProbe {
    $project=$packages['invSys.Admin.xlam'].VBProject
    $probe=$project.VBComponents.Item('TestDirectoryColumns').CodeModule
    $probe.InsertLines($probe.CountOfDeclarationLines+1,@'
Public RootAction As Boolean
Private mRootInput As String, mRootFault As Boolean, mRootSuccess As Long
Private mRootWrites As Long, mRootFailures As Long
'@)
    $probe.AddFromString(@'
Public Function RootWriteAllowed() As Boolean
    RootWriteAllowed = True
    If Not RootAction Then Exit Function
    mRootWrites = mRootWrites + 1
    If mRootFault Then
        mRootFailures = mRootFailures + 1
        RootWriteAllowed = False
    End If
End Function
Public Function InvokeRoot(ByVal enteredRoot As String, ByVal failWrite As Boolean) As String
    Dim control As New TestDirectoryRibbon, before As String, after As String
    Dim item As Variant, remembered As Boolean, selected As Boolean
    Dim number As Long, description As String
    On Error GoTo Failed
    mHost.Activate: mHost.Worksheets("Personal_notes").Activate
    If Not (modAdmin.ResolveInteractiveAdminWorkbook() Is mHost) Then Err.Raise 5, , "Root host not bound"
    mRootInput = enteredRoot: mRootFault = failWrite: mRootSuccess = 0
    mRootWrites = 0: mRootFailures = 0: mCalls = 0: mOpened = False
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
    Exit Function
Failed:
    number = Err.Number: description = Err.Description: RootAction = False
    Err.Raise number, , description
End Function
'@)
    $start=$probe.ProcStartLine('ObserveMessage',0);$count=$probe.ProcCountLines('ObserveMessage',0)
    $source=$probe.Lines($start,$count)
    $changed=$source.Replace('    If text = "Warehouse directory refreshed." Then',@'
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
    $project.VBComponents.Item('modAdminConsole').CodeModule.AddFromString(@'
Private Sub SaveSetting(ByVal app As String, ByVal section As String, ByVal key As String, ByVal setting As String)
    If key = "WarehouseScanRoots" Then
        If Not TestDirectoryColumns.RootWriteAllowed() Then Err.Raise 70, , "Generated preference-write refusal"
    End If
    VBA.SaveSetting app, section, key, setting
End Sub
'@)
}

function Test-AdminWarehouseRootWrite($Fixture,$Other,$HostBook,$OtherBook) {
    $before=Pins $Fixture.Root;$otherBefore=Pins $Other.Root
    $faultRoot=Join-Path $runRoot 'unremembered-root'
    New-Item -ItemType Directory -Path $faultRoot|Out-Null
    foreach($case in @('Success','Cancel','WriteFailure')){
        SelectTarget $Fixture
        $entered=if($case -ceq 'Success'){$Other.Root}elseif($case -ceq 'WriteFailure'){$faultRoot}else{''}
        $facts=([string](Invoke 'InvokeRoot' @($entered,($case -ceq 'WriteFailure')))) -split '\|'
        if($facts.Count -ne 10){throw 'Root callback facts unavailable; not product RED'}
        $prefix='WarehouseRoot.'+$case+'.';$cancel=$case -ceq 'Cancel';$fault=$case -ceq 'WriteFailure'
        Check ($prefix+'ActualPrompt') ($facts[3] -ceq '1')
        Check ($prefix+'ActualWriteBoundary') ($facts[4] -ceq $(if($cancel){'0'}else{'1'}) -and $facts[5] -ceq $(if($fault){'1'}else{'0'}))
        Check ($prefix+'ExistingDirectoryBehavior') ($facts[0] -ceq $(if($cancel){'0'}else{'1'}) -and $facts[1] -ceq [string](-not $cancel) -and $facts[8] -ceq 'True')
        Check ($prefix+'ExactSavedRootResult') ($facts[6] -ceq [string]($case -ceq 'Success') -and $facts[7] -ceq [string]($case -cne 'Success'))
        Check ($prefix+'NoFalseRememberedMessage') ($facts[2] -ceq $(if($case -ceq 'Success'){'1'}else{'0'}))
        Check ($prefix+'NoUnexpectedFeedback') ($facts[9] -ceq '0')
        $preserved=$true;foreach($part in @('Managed','Headers','Text','Formula','Tail','Notes')){$preserved=$preserved -and [bool](Invoke 'Preserved' @($part))}
        Check ($prefix+'CapturedWorkbookAndCustomColumnsPreserved') ($preserved -and [string]$OtherBook.Worksheets.Item(1).Cells.Item(1,1).Value2 -ceq 'Other workbook canary')
        Check ($prefix+'SavedAuthorityPreserved') ((Same $before (Pins $Fixture.Root)) -and (Same $otherBefore (Pins $Other.Root)))
    }
}
