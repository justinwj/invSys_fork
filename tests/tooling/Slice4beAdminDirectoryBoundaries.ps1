# Faults occur at optional logging entry or the actual activity-store Open.
function Install-AdminDirectoryBoundaries($CoreModule,$AdminProbe,[string]$ControlId='ADMIN_DIRECTORY_OPEN',[string]$Outcome='OPENED',[string]$ProbeName='TestDirectoryActivity',[switch]$RootAction) {
    $CoreModule.InsertLines($CoreModule.CountOfDeclarationLines+1,@'
Private mMode As String, mOtherRoot As String, mOtherSecret As String, mOtherBook As String
Private mEntryObserved As Boolean, mWriteFailures As Long
'@)
    $boundarySource=@'
Public Sub ArmBoundary(ByVal mode As String, ByVal root As String, ByVal secret As String, ByVal book As String)
    mMode = mode: mOtherRoot = root: mOtherSecret = secret: mOtherBook = book
    mEntryObserved = False: mWriteFailures = 0
End Sub
Public Sub AtEntry(ByVal controlId As String)
    Dim status As String
    If controlId <> "ADMIN_DIRECTORY_OPEN" Or mEntryObserved Then Exit Sub
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
            mEntryObserved = (Left$(status, 3) = "OK|")
    End Select
End Sub
Public Sub AtWrite(ByVal body As String, ByRef pending As String)
    Dim record As Object, wanted As String
    If mMode <> "AttemptWriteFailure" And mMode <> "OutcomeWriteFailure" Then Exit Sub
    Set record = modTrainingJson.DecodeObject(body)
    If record("ControlId") <> "ADMIN_DIRECTORY_OPEN" Then Exit Sub
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
'@
    $CoreModule.AddFromString($boundarySource.Replace('ADMIN_DIRECTORY_OPEN',$ControlId).Replace('"OPENED"',('"'+$Outcome+'"')))
    $project=$packages['invSys.Core.xlam'].VBProject
    $activity=$project.VBComponents.Item('modActivity').CodeModule
    $start=$activity.ProcStartLine('BeginAction',0);$end=$start+$activity.ProcCountLines('BeginAction',0)
    $lines=@(for($line=$start;$line -lt $end;$line++){if($activity.Lines($line,1).Trim() -ceq 'notice = ""'){$line}})
    if($lines.Count -ne 1){throw 'Directory observation entry boundary changed.'}
    $activity.InsertLines($lines[0],('    '+$ProbeName+'.AtEntry controlId'))
    $store=$project.VBComponents.Item('modActivityStore').CodeModule
    $lines=@(for($line=1;$line -le $store.CountOfLines;$line++){if($store.Lines($line,1).Trim() -ceq 'Open pending For Binary Access Write Lock Read Write As #file'){$line}})
    if($lines.Count -ne 1){throw 'Directory actual store-write boundary unavailable.'}
    $store.InsertLines($lines[0],('    '+$ProbeName+'.AtWrite body, pending'))
    if($RootAction){return}
    $AdminProbe.InsertLines($AdminProbe.CountOfDeclarationLines+1,'Private mTrackingNotices As Long')
    $AdminProbe.AddFromString(@'
Public Function TrackingNotices() As Long
    TrackingNotices = mTrackingNotices
End Function
'@)
    $start=$AdminProbe.ProcStartLine('ObserveMessage',0);$count=$AdminProbe.ProcCountLines('ObserveMessage',0)
    $source=$AdminProbe.Lines($start,$count)
    $changed=$source.Replace('    If text = "Warehouse directory refreshed." Then',@'
    If Left$(text, 21) = "Tracking unavailable:" Then
        mTrackingNotices = mTrackingNotices + 1
        ObserveMessage = vbOK
        Exit Function
    End If
    If text = "Warehouse directory refreshed." Then
'@.TrimEnd())
    if($source -ceq $changed){throw 'Directory message observation seam changed.'}
    $AdminProbe.DeleteLines($start,$count);$AdminProbe.InsertLines($start,$changed)
    $start=$AdminProbe.ProcStartLine('Invoke',0);$count=$AdminProbe.ProcCountLines('Invoke',0)
    $source=$AdminProbe.Lines($start,$count);$changed=$source.Replace('    RibbonOnActionAdmin control',"    mTrackingNotices = 0`r`n    RibbonOnActionAdmin control")
    if($source -ceq $changed){throw 'Directory invocation reset seam changed.'}
    $AdminProbe.DeleteLines($start,$count);$AdminProbe.InsertLines($start,$changed)
}

function Test-AdminDirectoryBoundaries($Fixture,$Other,$HostBook,$OtherBook) {
    foreach($case in @('SwitchWorkbook','SignOut','SwitchTarget','AttemptWriteFailure','OutcomeWriteFailure')){
        SelectTarget $Fixture
        if(-not [bool](Probe 'SetPolicy' @('Enabled'))){throw 'Directory boundary policy prerequisite failed.'}
        $before=@(Get-Slice4beActivityFiles $Fixture);$otherBefore=@(Get-Slice4beActivityFiles $Other)
        $authority=Pins $Fixture.Root;$otherAuthority=Pins $Other.Root
        [void](Probe 'ArmBoundary' @($case,$Other.Root,$Other.Secret,$OtherBook.Name))
        try{
            $values=([string](Invoke 'Invoke')) -split '\|'
            if($values.Count -ne 6){throw 'Directory boundary callback facts unavailable.'}
            $reached=if($case -like '*WriteFailure'){[int](Probe 'WriteFailures') -eq 1}else{[bool](Probe 'EntryObserved')}
            if(-not $reached){throw 'Directory real optional boundary was not reached.'}
            $prefix='Directory.Boundary.'+$case+'.'
            Check ($prefix+'ActualFaultReached') $reached
            Check ($prefix+'OwnerOnceWithOriginalPresentation') ($values[0] -ceq '1' -and $values[1] -ceq 'True' -and $values[2] -ceq 'True' -and $values[3] -ceq '1' -and $values[4] -ceq '0')
            Check ($prefix+'OneTruthfulNoticeOrNone') ([int](Invoke 'TrackingNotices') -eq $(if($case -ceq 'SwitchWorkbook'){0}else{1}))
            Check ($prefix+'OriginalReadScopeRetained') ([bool](Invoke 'Preserved' @('Managed')))
            $preserved=$true;foreach($part in @('Headers','Text','Formula','Tail','Notes')){$preserved=$preserved -and [bool](Invoke 'Preserved' @($part))}
            Check ($prefix+'LocalAndOtherWorkbookPreserved') ($preserved -and [string]$OtherBook.Worksheets.Item(1).Cells.Item(1,1).Value2 -ceq 'Other directory workbook canary')
            Check ($prefix+'SavedAuthorityAndHistoryPreserved') ((Retained $authority) -and (Retained $otherAuthority))
            $records=@(Get-Slice4beActivityFiles $Fixture|Where-Object {$_ -cnotin $before}|ForEach-Object {Get-Content $_ -Raw|ConvertFrom-Json}|Where-Object ControlId -CEQ 'ADMIN_DIRECTORY_OPEN')
            $honest=if($case -ceq 'SwitchWorkbook'){$records.Count -eq 2 -and @($records|Where-Object OutcomeCode -CEQ 'OPENED').Count -eq 1}elseif($case -ceq 'OutcomeWriteFailure'){$records.Count -eq 1 -and $records[0].OutcomeCode -ceq 'REQUESTED'}else{$records.Count -eq 0}
            Check ($prefix+'HonestOriginalEvidence') $honest
            Check ($prefix+'NoEvidenceInOtherWarehouse') ((@(Get-Slice4beActivityFiles $Other) -join '|') -ceq ($otherBefore -join '|'))
        }finally{[void](Probe 'ArmBoundary' @('','','',''));SelectTarget $Fixture}
    }
}
