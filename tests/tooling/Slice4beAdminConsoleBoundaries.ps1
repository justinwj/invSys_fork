# Inject only at optional observation boundaries; original reads and writes run.
function Install-AdminConsoleBoundaries($CoreModule,$AdminProbe,$AdminProject) {
    $CoreModule.InsertLines($CoreModule.CountOfDeclarationLines+1,@'
Private mMode As String, mOtherRoot As String, mOtherSecret As String, mOtherBook As String
Private mEntryObserved As Boolean, mWriteFailures As Long
'@)
    $CoreModule.AddFromString(@'
Public Sub ArmBoundary(ByVal mode As String, ByVal root As String, ByVal secret As String, ByVal book As String)
    mMode = mode: mOtherRoot = root: mOtherSecret = secret: mOtherBook = book
    mEntryObserved = False: mWriteFailures = 0
End Sub
Public Sub AtEntry(ByVal controlId As String)
    Dim status As String
    If controlId <> "ADMIN_CONSOLE_OPEN" Or mEntryObserved Then Exit Sub
    Select Case mMode
        Case "SwitchWorkbook"
            Application.Workbooks(mOtherBook).Activate
            mEntryObserved = True
        Case "SignOut"
            modAuth.SignOut
            mEntryObserved = Not modAuth.IsSignedIn()
        Case "SwitchTarget", "EmptyScopeThenSwitchTarget"
            modRuntimeWorkbooks.SetCoreDataRootOverride mOtherRoot
            status = modNasConnection.SelectWarehouseTargetForAutomation(mOtherRoot, mOtherRoot, "S1", False)
            If Left$(status, 3) <> "OK|" Then Exit Sub
            status = modAuth.SignInCurrentTargetForAutomation("config-admin", mOtherSecret, "")
            mEntryObserved = (Left$(status, 3) = "OK|")
    End Select
End Sub
Public Sub AtWrite(ByVal body As String, ByRef pending As String)
    Dim record As Object, wanted As String
    If mMode <> "AttemptWriteFailure" And mMode <> "OutcomeWriteFailure" Then Exit Sub
    Set record = modTrainingJson.DecodeObject(body)
    If record("ControlId") <> "ADMIN_CONSOLE_OPEN" Then Exit Sub
    wanted = "REQUESTED": If mMode = "OutcomeWriteFailure" Then wanted = "OPENED"
    If record("OutcomeCode") <> wanted Then Exit Sub
    mWriteFailures = mWriteFailures + 1
    pending = pending & "\unavailable.pending"
End Sub
Public Function EntryObserved() As Boolean
    EntryObserved = mEntryObserved
End Function
Public Function WriteFailures() As Long
    WriteFailures = mWriteFailures
End Function
'@)
    $activity=$packages['invSys.Core.xlam'].VBProject.VBComponents.Item('modActivity').CodeModule
    $start=$activity.ProcStartLine('BeginAction',0);$end=$start+$activity.ProcCountLines('BeginAction',0)
    $lines=@(for($line=$start;$line -lt $end;$line++){if($activity.Lines($line,1).Trim() -ceq 'notice = ""'){$line}})
    if($lines.Count -ne 1){throw 'Console observation entry boundary changed.'}
    $activity.InsertLines($lines[0],'    TestConsoleActivity.AtEntry controlId')
    $store=$packages['invSys.Core.xlam'].VBProject.VBComponents.Item('modActivityStore').CodeModule
    $lines=@(for($line=1;$line -le $store.CountOfLines;$line++){if($store.Lines($line,1).Trim() -ceq 'Open pending For Binary Access Write Lock Read Write As #file'){$line}})
    if($lines.Count -ne 1){throw 'Real activity-store write boundary changed.'}
    $store.InsertLines($lines[0],'    TestConsoleActivity.AtWrite body, pending')
    $AdminProbe.InsertLines($AdminProbe.CountOfDeclarationLines+1,@'
Private mExpectedConfig As String, mReadMatches As Boolean, mReadCount As Long
Private mConsoleScanRoot As String
'@)
    $AdminProbe.AddFromString(@'
Public Sub SetScanRoot(ByVal root As String)
    mConsoleScanRoot = root
End Sub
Public Function ScanRoot() As String
    ScanRoot = mConsoleScanRoot
End Function
Public Sub PrepareReadScope()
    mExpectedConfig = modAdminConsole.CaptureConsoleConfigPath()
    mReadMatches = False: mReadCount = 0
End Sub
Public Sub ObserveReadScope(ByVal configPath As String)
    mReadMatches = (StrComp(configPath, mExpectedConfig, vbTextCompare) = 0)
    mReadCount = mReadCount + 1
End Sub
Public Function ReadScopePreserved() As Boolean
    ReadScopePreserved = (mReadCount = 1 And mReadMatches)
End Function
Public Function EmptyScopePreserved(ByVal name As String) As Boolean
    EmptyScopePreserved = (mExpectedConfig = "" And ReadScopePreserved() And _
        Application.Workbooks(name).Worksheets("AdminConsole").Range("B3").Value2 = "<none>")
End Function
'@)
    $start=$AdminProbe.ProcStartLine('Invoke',0);$count=$AdminProbe.ProcCountLines('Invoke',0)
    $source=$AdminProbe.Lines($start,$count);$changed=$source.Replace('    RibbonOnActionAdmin control',"    PrepareReadScope`r`n    RibbonOnActionAdmin control")
    if($source -ceq $changed){throw 'Console callback observation boundary changed.'}
    $AdminProbe.DeleteLines($start,$count);$AdminProbe.InsertLines($start,$changed)
    $owner=$AdminProject.VBComponents.Item('modAdminConsole').CodeModule
    $start=$owner.ProcStartLine('TryResolveExistingAdminRuntimeContext',0);$count=$owner.ProcCountLines('TryResolveExistingAdminRuntimeContext',0)
    $source=$owner.Lines($start,$count);$changed=$source.Replace('    If configPath = "" Then Exit Function',"    TestAdminConsole.ObserveReadScope configPath`r`n    If configPath = `"`" Then Exit Function")
    if($source -ceq $changed){throw 'Actual console config read boundary changed.'}
    $owner.DeleteLines($start,$count);$owner.InsertLines($start,$changed)
    $start=$owner.ProcStartLine('GetAdminRuntimeScanRoot',0);$count=$owner.ProcCountLines('GetAdminRuntimeScanRoot',0)
    $source=$owner.Lines($start,$count)
    $changed=$source.Replace('    rootPath = Trim$(modRuntimeWorkbooks.GetCoreDataRootOverride())',@'
    If TestAdminConsole.ScanRoot() <> "" Then
        GetAdminRuntimeScanRoot = TestAdminConsole.ScanRoot()
        Exit Function
    End If
    rootPath = Trim$(modRuntimeWorkbooks.GetCoreDataRootOverride())
'@)
    if($source -ceq $changed){throw 'Console-only scan-root fixture boundary changed.'}
    $owner.DeleteLines($start,$count);$owner.InsertLines($start,$changed)
}

function Test-AdminConsoleBoundaries($Fixture,$Other,$HostBook,$OtherBook) {
    $empty=Join-Path $runRoot 'empty-console-scope';New-Item -ItemType Directory -Path $empty|Out-Null
    $otherCount=$OtherBook.Worksheets.Count
    foreach($case in @('SwitchWorkbook','SignOut','SwitchTarget','EmptyScopeThenSwitchTarget','AttemptWriteFailure','OutcomeWriteFailure')){
        SelectTarget $Fixture
        if(-not [bool](Probe 'SetPolicy' @('Enabled'))){throw 'Console boundary policy prerequisite failed.'}
        $before=@(Get-Slice4beActivityFiles $Fixture);$otherBefore=@(Get-Slice4beActivityFiles $Other)
        $authority=Pins $Fixture.Root;$otherAuthority=Pins $Other.Root
        if($case -ceq 'EmptyScopeThenSwitchTarget'){[void](Run 'invSys.Admin.xlam' 'TestAdminConsole.SetScanRoot' @($empty))}
        [void](Probe 'ArmBoundary' @($case,$Other.Root,$Other.Secret,$OtherBook.Name))
        try{
            $values=([string](Run 'invSys.Admin.xlam' 'TestAdminConsole.Invoke' @($HostBook.Name))) -split '\|'
            if($values.Count -ne 7){throw 'Console boundary observation missing.'}
            $reached=if($case -like '*WriteFailure'){[int](Probe 'WriteFailures') -eq 1}else{[bool](Probe 'EntryObserved')}
            $prefix='AdminConsole.Boundary.'+$case+'.'
            Check ($prefix+'ActualFaultReached') $reached
            if(-not $reached){throw 'Console boundary was not reached; not product RED.'}
            Check ($prefix+'OwnerOpensOnceAndRetainsSurface') ($values[0] -ceq '1' -and $values[1] -ceq 'True' -and $values[2] -ceq 'True' -and $values[5] -ceq 'True')
            $count=if($case -ceq 'SwitchWorkbook'){'0'}else{'1'}
            Check ($prefix+'OneTruthfulTrackingNotice') ($values[3] -ceq $count -and $values[4] -ceq '0')
            Check ($prefix+'OriginalReadScopeUsed') ([bool](Run 'invSys.Admin.xlam' 'TestAdminConsole.ReadScopePreserved'))
            Check ($prefix+'OriginalAndOtherWorkbooksPreserved') ([bool](Run 'invSys.Admin.xlam' 'TestAdminConsole.HostPreserved' @($HostBook.Name)) -and $OtherBook.Worksheets.Count -eq $otherCount -and [string]$OtherBook.Worksheets.Item(1).Cells.Item(1,1).Value2 -ceq 'Other workbook canary')
            Check ($prefix+'SavedAuthorityAndPriorEvidencePreserved') ((Retained $authority) -and (Retained $otherAuthority))
            $records=@(Get-Slice4beActivityFiles $Fixture|Where-Object {$_ -cnotin $before}|ForEach-Object {Get-Content $_ -Raw|ConvertFrom-Json}|Where-Object ControlId -CEQ 'ADMIN_CONSOLE_OPEN')
            $honest=if($case -ceq 'SwitchWorkbook'){$records.Count -eq 2 -and @($records|Where-Object OutcomeCode -CEQ 'OPENED').Count -eq 1}elseif($case -ceq 'OutcomeWriteFailure'){$records.Count -eq 1 -and $records[0].OutcomeCode -ceq 'REQUESTED'}else{$records.Count -eq 0}
            Check ($prefix+'HonestOriginalEvidence') $honest
            Check ($prefix+'NoLaterWarehouseAttribution') (@(Get-Slice4beActivityFiles $Other).Count -eq $otherBefore.Count)
            if($case -ceq 'EmptyScopeThenSwitchTarget'){Check ($prefix+'EmptyScopeCannotDiscoverLaterTarget') ([bool](Run 'invSys.Admin.xlam' 'TestAdminConsole.EmptyScopePreserved' @($HostBook.Name)) -and @(Get-ChildItem -LiteralPath $empty -Recurse -File).Count -eq 0)}
        }finally{[void](Probe 'ArmBoundary' @('','','',''));[void](Run 'invSys.Admin.xlam' 'TestAdminConsole.SetScanRoot' @(''));SelectTarget $Fixture}
    }
}
