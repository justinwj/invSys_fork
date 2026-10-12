# Faults and context changes enter only at the optional observation boundary.
function Install-AdminVerifyBoundaries($Module) {
    $Module.InsertLines($Module.CountOfDeclarationLines+1,@'
Private mMode As String, mOtherRoot As String, mOtherSecret As String, mOtherBook As String
Private mEntryObserved As Boolean, mWriteFailures As Long
'@)
    $Module.AddFromString(@'
Public Sub ArmBoundary(ByVal mode As String, ByVal root As String, ByVal secret As String, ByVal book As String)
    mMode = mode: mOtherRoot = root: mOtherSecret = secret: mOtherBook = book
    mEntryObserved = False: mWriteFailures = 0
End Sub
Public Sub AtEntry(ByVal controlId As String)
    Dim status As String
    If controlId <> "ADMIN_VERIFY_ADDINS" Or mEntryObserved Then Exit Sub
    Select Case mMode
        Case "SwitchWorkbook"
            Application.Workbooks(mOtherBook).Activate
            mEntryObserved = True
        Case "SignOut"
            modAuth.SignOut
            mEntryObserved = Not modAuth.IsSignedIn()
        Case "SwitchTarget"
            modRuntimeWorkbooks.SetCoreDataRootOverride mOtherRoot
            status = modNasConnection.SelectWarehouseTargetForAutomation(mOtherRoot, mOtherRoot, "S1", False)
            If Left$(status, 3) <> "OK|" Then Exit Sub
            status = modAuth.SignInCurrentTargetForAutomation("config-admin", mOtherSecret, "")
            If Left$(status, 3) <> "OK|" Then Exit Sub
            ' A test-only cross-XLAM call removes the fixture override so the real
            ' owner must use its captured root rather than the later config.
            Application.Run "'invSys.Admin.xlam'!modAddinsPublish.ClearAddinsPublishSharePointRootOverride"
            mEntryObserved = True
    End Select
End Sub
Public Sub AtWrite(ByVal body As String, ByRef pending As String)
    Dim record As Object, wanted As String
    If mMode <> "AttemptWriteFailure" And mMode <> "OutcomeWriteFailure" Then Exit Sub
    Set record = modTrainingJson.DecodeObject(body)
    If record("ControlId") <> "ADMIN_VERIFY_ADDINS" Then Exit Sub
    wanted = "REQUESTED": If mMode = "OutcomeWriteFailure" Then wanted = "VERIFIED"
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
    $core=$packages['invSys.Core.xlam'].VBProject
    $activity=$core.VBComponents.Item('modActivity').CodeModule
    $start=$activity.ProcStartLine('BeginAction',0);$end=$start+$activity.ProcCountLines('BeginAction',0)
    $entries=@(for($line=$start;$line -lt $end;$line++){if($activity.Lines($line,1).Trim() -ceq 'notice = ""'){$line}})
    if($entries.Count -ne 1){throw 'Activity entry statement boundary changed; not product RED.'}
    $activity.InsertLines($entries[0],'    TestVerifyActivity.AtEntry controlId')
    $store=$core.VBComponents.Item('modActivityStore').CodeModule
    $hits=@(for($line=1;$line -le $store.CountOfLines;$line++){if($store.Lines($line,1).Trim() -ceq 'Open pending For Binary Access Write Lock Read Write As #file'){$line}})
    if($hits.Count -ne 1){throw 'Activity store Open boundary changed; not product RED.'}
    $store.InsertLines($hits[0],'    TestVerifyActivity.AtWrite body, pending')
}

function Test-AdminVerifyBoundaries($Fixture,$Other,$HostBook,$OtherBook) {
    $root=Join-Path $runRoot 'verify-boundaries';$addins=Join-Path $root 'Addins'
    New-Item -ItemType Directory -Path $addins -Force|Out-Null
    foreach($name in @('Core','Inventory.Domain','Designs.Domain','Operations','Admin')){[IO.File]::WriteAllText((Join-Path $addins ('invSys.'+$name+'.xlam')),'Non-executable file-presence fixture')}
    $files=Pins $root;$otherSheets=$OtherBook.Worksheets.Count
    foreach($case in @('SwitchWorkbook','SignOut','SwitchTarget','AttemptWriteFailure','OutcomeWriteFailure')){
        SelectTarget $Fixture
        if(-not [bool](Probe 'SetPolicy' @('Enabled'))){throw 'Boundary policy prerequisite unavailable.'}
        $before=@(Get-Slice4beActivityFiles $Fixture);$otherBefore=@(Get-Slice4beActivityFiles $Other)
        $authority=Pins $Other.Root;$configHash=FileHash $Fixture.Config
        $inventory=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Data.Inventory.xlsb');$inventoryHash=FileHash $inventory
        [void](Probe 'ArmBoundary' @($case,$Other.Root,$Other.Secret,$OtherBook.Name))
        try{
            $values=([string](Run 'invSys.Admin.xlam' 'TestAdminVerify.Invoke' @($root,$HostBook.Name,$OtherBook.Name))) -split '\|'
            if($values.Count -ne 7){throw 'Boundary observation unavailable.'}
            $observed=if($case -like '*WriteFailure'){[int](Probe 'WriteFailures') -eq 1}else{[bool](Probe 'EntryObserved')}
            Check ('AdminVerify.Boundary.'+$case+'.FaultReached') $observed
            if(-not $observed){throw 'Boundary mutation did not occur; not product RED.'}
            $prefix='AdminVerify.Boundary.'+$case+'.'
            Check ($prefix+'OwnerInvokedExactlyOnce') ($values[0] -ceq '1')
            Check ($prefix+'OriginalVerificationResult') ($values[1] -ceq 'VERIFIED')
            Check ($prefix+'OriginalRootUsed') ($values[6] -ceq 'True')
            $notices=if($case -ceq 'SwitchWorkbook'){'0'}else{'1'}
            Check ($prefix+'ExactlyExpectedTrackingNotice') ($values[2] -ceq $notices)
            Check ($prefix+'WorkbookBindingPreserved') ([bool](Run 'invSys.Admin.xlam' 'TestAdminVerify.HostPreserved' @($HostBook.Name)) -and $OtherBook.Worksheets.Count -eq $otherSheets -and [string]$OtherBook.Worksheets.Item(1).Cells.Item(1,1).Value2 -ceq 'Other workbook canary')
            Check ($prefix+'ReadOnlyFilesAndAuthority') ((SamePins $files (Pins $root)) -and (SamePins $authority (Pins $Other.Root)) -and (FileHash $Fixture.Config) -ceq $configHash -and (FileHash $inventory) -ceq $inventoryHash)
            $records=@(Get-Slice4beActivityFiles $Fixture|Where-Object {$_ -cnotin $before}|ForEach-Object {Get-Content $_ -Raw|ConvertFrom-Json}|Where-Object ControlId -CEQ 'ADMIN_VERIFY_ADDINS')
            $honest=if($case -ceq 'SwitchWorkbook'){$records.Count -eq 2 -and @($records|Where-Object OutcomeCode -CEQ 'VERIFIED').Count -eq 1}elseif($case -ceq 'OutcomeWriteFailure'){$records.Count -eq 1 -and $records[0].OutcomeCode -ceq 'REQUESTED'}else{$records.Count -eq 0}
            Check ($prefix+'HonestOriginalEvidence') $honest
            Check ($prefix+'NoLaterWarehouseAttribution') (@(Get-Slice4beActivityFiles $Other).Count -eq $otherBefore.Count)
        }finally{[void](Probe 'ArmBoundary' @('','','',''));SelectTarget $Fixture}
    }
}
