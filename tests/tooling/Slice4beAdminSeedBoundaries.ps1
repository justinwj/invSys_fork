# Real activity Open failures and session changes; business callback/owner stay real.
function Install-AdminSeedBoundariesProbe($AdminProbeModule) {
    $core=$packages['invSys.Core.xlam'].VBProject
    $probe=$core.VBComponents.Add(1);$probe.Name='TestSeedBoundary'
    $probe.CodeModule.AddFromString(@'
Option Explicit
Private mMode As String, mOutcome As String
Private mInjected As Long, mFailures As Long, mError As Long, mChanged As Boolean, mNestedQuiet As Boolean
Public Sub Arm(ByVal mode As String, ByVal outcome As String)
    mMode = mode: mOutcome = outcome
    mInjected = 0: mFailures = 0: mError = 0: mChanged = False: mNestedQuiet = False
End Sub
Public Sub AtEntry(ByVal controlId As String)
    If controlId <> "ADMIN_DEMO_SEED" Or mMode <> "EntrySignOut" Then Exit Sub
    If mInjected <> 0 Then Exit Sub
    mInjected = 1: modAuth.SignOut
    mChanged = (Not modAuth.IsSignedIn() And modActivity.CaptureContext() = "")
End Sub
Public Sub AtFinish(ByVal activityId As String)
    If mMode <> "FinishSignOut" Or activityId = "" Or mInjected <> 0 Then Exit Sub
    mInjected = 1: modAuth.SignOut
    mChanged = (Not modAuth.IsSignedIn() And modActivity.CaptureContext() = "")
End Sub
Public Sub AtWrite(ByVal body As String, ByRef pending As String)
    Dim record As Object, nested As String, notice As String
    If mMode <> "Storage" And mMode <> "Reenter" Then Exit Sub
    Set record = modTrainingJson.DecodeObject(body)
    If record("ControlId") <> "ADMIN_DEMO_SEED" Or record("OutcomeCode") <> mOutcome Then Exit Sub
    mInjected = mInjected + 1
    If mMode = "Storage" Then
        pending = pending & "\unavailable.pending"
    Else
        nested = modActivity.BeginAction("ADMIN_DEMO_SEED", modActivity.CaptureContext(), notice)
        mNestedQuiet = (nested = "" And notice = "")
    End If
End Sub
Public Sub AtFailure(ByVal number As Long)
    If mMode <> "Storage" Or mInjected = 0 Then Exit Sub
    mFailures = mFailures + 1: mError = number
End Sub
Public Function Observations() As String
    Observations = CStr(mInjected) & "|" & CStr(mFailures) & "|" & CStr(mError <> 0) & "|" & CStr(mChanged) & "|" & CStr(mNestedQuiet)
End Function
'@)
    foreach($case in @(
        @('modActivityStore','Append','Open pending For Binary Access Write Lock Read Write As #file','    TestSeedBoundary.AtWrite body, pending',$true),
        @('modActivityStore','Append','Failed:','    TestSeedBoundary.AtFailure Err.Number',$false),
        @('modActivity','BeginAction','notice = ""','    TestSeedBoundary.AtEntry controlId',$false),
        @('modActivity','FinishAction','On Error GoTo Failed','    TestSeedBoundary.AtFinish activityId',$false)
    )){
        $module=$core.VBComponents.Item($case[0]).CodeModule
        $start=$module.ProcStartLine($case[1],0);$count=$module.ProcCountLines($case[1],0)
        $hits=@(for($line=$start;$line -lt $start+$count;$line++){if($module.Lines($line,1).Trim() -ieq $case[2]){$line}})
        if($hits.Count -ne 1){throw ('Seed boundary unavailable: '+$case[0]+'.'+$case[1]+'; not product RED.')}
        $position=$hits[0];if(-not $case[4]){$position++}
        $module.InsertLines($position,$case[3])
    }
    $AdminProbeModule.AddFromString(@'
Private mBoundaryOwnerCalls As Long, mBoundaryNotices As Long, mBoundaryNoticeSafe As Boolean
Private mBoundaryInventoryPath As String
Public Sub ResetBoundary()
    mBoundaryOwnerCalls = 0: mBoundaryNotices = 0: mBoundaryNoticeSafe = True
End Sub
Public Sub BoundaryOwner()
    mBoundaryOwnerCalls = mBoundaryOwnerCalls + 1
End Sub
Public Sub BoundaryNotice(ByVal notice As String)
    mBoundaryNotices = mBoundaryNotices + 1
    Select Case notice
        Case "Tracking unavailable: the training record could not be saved.", _
             "Tracking unavailable: Seed could not be recorded for the captured warehouse and user.", _
             "Tracking unavailable: the action context is no longer current."
        Case Else: mBoundaryNoticeSafe = False
    End Select
End Sub
Public Function BoundaryCounts() As String
    BoundaryCounts = CStr(mBoundaryOwnerCalls) & "|" & CStr(mBoundaryNotices) & "|" & CStr(mBoundaryNoticeSafe)
End Function
Public Function CloseBoundaryProjection() As Boolean
    Dim wb As Workbook
    Set wb = modInventoryDomainBridge.ResolveInventoryWorkbookBridge(mWarehouse)
    ' Publication may dirty the live projection after canonical persistence.
    ' Discard only this generated fixture's unsaved projection, then inspect disk.
    mBoundaryInventoryPath = wb.FullName: wb.Close SaveChanges:=False
    CloseBoundaryProjection = True
End Function
Public Function ExactBoundaryEntity(ByVal sku As String, ByVal quantity As Double) As Boolean
    Dim wb As Workbook, lo As ListObject, row As Long, found As Long, receipts As Long
    If mBoundaryInventoryPath = "" Then Exit Function
    Set wb = Application.Workbooks.Open(Filename:=mBoundaryInventoryPath, UpdateLinks:=0, ReadOnly:=True, AddToMru:=False)
    Set lo = FindTable(wb, "tblInventoryEntities")
    For row = 1 To lo.ListRows.Count
        If StrComp(CStr(Value(lo, row, "SKU")), sku, vbBinaryCompare) = 0 Then
            If CDbl(Value(lo, row, "QtyOnHand")) = quantity And CStr(Value(lo, row, "System_Key")) <> "" Then found = found + 1
        End If
    Next row
    Set lo = FindTable(wb, "tblAppliedEvents")
    For row = 1 To lo.ListRows.Count
        If StrComp(CStr(Value(lo, row, "EventID")), mSeedEvent, vbBinaryCompare) = 0 Then
            If CStr(Value(lo, row, "Status")) = "APPLIED" Then receipts = receipts + 1
        End If
    Next row
    ExactBoundaryEntity = (found = 1 And receipts = 1 And wb.ReadOnly And wb.Saved)
    wb.Close SaveChanges:=False
End Function
'@)
    $admin=$packages['invSys.Admin.xlam'].VBProject
    $observer=$admin.VBComponents.Item('modAdminSeedObservation').CodeModule
    $hits=0
    for($line=1;$line -le $observer.CountOfLines;$line++){
        if($observer.Lines($line,1).Trim() -ieq 'If notice <> "" Then MsgBox notice, vbExclamation, "invSys Admin"'){
            $observer.ReplaceLine($line,'    If notice <> "" Then TestSeedOutcome.BoundaryNotice notice');$hits++
        }
    }
    if($hits -ne 2){throw 'Explicit Seed notice seam unavailable; not product RED.'}
    $owner=$admin.VBComponents.Item('modAdminInventorySeed').CodeModule
    foreach($procedure in @('SeedDemoInventoryForWarehouse','UploadDemoInventoryForWarehouse')){
        $start=$owner.ProcStartLine($procedure,0);$count=$owner.ProcCountLines($procedure,0)
        $statement=if($procedure -ceq 'SeedDemoInventoryForWarehouse'){'outcome = "FAILED": sourceEventId = ""'}else{'outcome = "REJECTED": sourceEventId = ""'}
        $hits=@(for($line=$start;$line -lt $start+$count;$line++){if($owner.Lines($line,1).Trim() -ieq $statement){$line}})
        if($hits.Count -ne 1){throw 'Seed owner-entry observation unavailable.'}
        $owner.InsertLines($hits[0]+1,'    TestSeedOutcome.BoundaryOwner')
    }
}

function Test-AdminSeedBoundaries($Fixture,$Other) {
    function SeedPins { $pins=@{};foreach($path in @(Get-Slice4beActivityFiles $Fixture)){$pins[$path]=(Get-FileHash -LiteralPath $path).Hash};return $pins }
    $auth=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Auth.xlsb')
    $otherAuth=Join-Path $Other.Root ($Other.Warehouse+'.invSys.Auth.xlsb')
    $authHash=(Get-FileHash $auth).Hash;$otherHash=(Get-FileHash $otherAuth).Hash
    try{
        foreach($case in @(
            @('AttemptUpload','Storage','REQUESTED',$true),
            @('OutcomeUpload','Storage','APPLIED',$true),
            @('AttemptNoOp','Storage','REQUESTED',$false),
            @('OutcomeNoOp','Storage','UNCHANGED',$false),
            @('EntrySignOut','EntrySignOut','',$false),
            @('FinishSignOut','FinishSignOut','',$false),
            @('Reenter','Reenter','REQUESTED',$true)
        )){
            SelectTarget $Fixture
            if(-not [bool](Run 'invSys.Core.xlam' 'TestSeedRecording.EnablePolicy')){throw 'Eligible Seed policy unavailable.'}
            if(-not [bool](Run 'invSys.Admin.xlam' 'TestSeedOutcome.Prepare' @($Fixture.Warehouse,$false))){throw 'Seed inbox fixture unavailable.'}
            $path='';$sku='DEMO-BOUNDARY-'+$case[0].ToUpperInvariant()
            if($case[3]){
                $path=Join-Path $Fixture.Root ('boundary-'+$case[0]+'.csv')
                [IO.File]::WriteAllText($path,("ITEM_CODE,ITEM,QTY,UOM,LOCATION`r`n"+$sku+",Boundary fixture,11,each,TRAINING`r`n"),[Text.UTF8Encoding]::new($false))
            }
            $before=SeedPins;$fingerprint=[string](Run 'invSys.Admin.xlam' 'TestSeedOutcome.WorkflowFingerprint')
            [void](Run 'invSys.Admin.xlam' 'TestSeedOutcome.ResetBoundary')
            [void](Run 'invSys.Core.xlam' 'TestSeedBoundary.Arm' @($case[1],$case[2]))
            try{
                $ack=[string](Run 'invSys.Admin.xlam' 'TestSeedOutcome.InvokeWorkflow' @('SEED',$path,$true))
                $observed=([string](Run 'invSys.Core.xlam' 'TestSeedBoundary.Observations')) -split '\|'
                $counts=([string](Run 'invSys.Admin.xlam' 'TestSeedOutcome.BoundaryCounts')) -split '\|'
            }finally{[void](Run 'invSys.Core.xlam' 'TestSeedBoundary.Arm' @('',''))}
            $prefix='SeedBoundary.'+$case[0]
            $reached=$observed[0] -ceq '1'
            if($case[1] -ceq 'Storage'){$reached=$reached -and $observed[1] -ceq '1' -and $observed[2] -ceq 'True'}
            elseif($case[1] -ceq 'Reenter'){$reached=$reached -and $observed[4] -ceq 'True'}
            else{$reached=$reached -and $observed[3] -ceq 'True'}
            Check ($prefix+'.ActualBoundaryReached') $reached
            if(-not $reached){throw 'Required Seed failure/context fixture not reached; not product RED.'}
            Check ($prefix+'.BusinessOwnerOnceAndSuccess') ($ack -ceq 'OK' -and $counts[0] -ceq '1')
            $notices=if($case[1] -ceq 'Reenter'){'0'}else{'1'}
            Check ($prefix+'.FixedNoticeOnceOrQuietReentry') ($counts[1] -ceq $notices -and $counts[2] -ceq 'True')
            $state=([string](Run 'invSys.Admin.xlam' 'TestSeedOutcome.State')) -split '\|'
            if($case[3]){
                $inventoryPath=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Data.Inventory.xlsb')
                if(-not [bool](Run 'invSys.Admin.xlam' 'TestSeedOutcome.CloseBoundaryProjection')){throw 'Generated projection did not close before fresh read.'}
                $savedHash=(Get-FileHash -LiteralPath $inventoryPath).Hash
                Check ($prefix+'.ExactSavedApplication') ($state[0] -ceq '1' -and $state[2] -ceq 'True' -and $state[6] -ceq 'True' -and [bool](Run 'invSys.Admin.xlam' 'TestSeedOutcome.ExactBoundaryEntity' @($sku,11)))
                Check ($prefix+'.FreshReadPreservesSavedBytes') ((Get-FileHash -LiteralPath $inventoryPath).Hash -ceq $savedHash)
            }else{
                Check ($prefix+'.NoOpPreservesInventoryAndInbox') ($state[0] -ceq '0' -and ([string](Run 'invSys.Admin.xlam' 'TestSeedOutcome.WorkflowFingerprint')) -ceq $fingerprint)
            }
            $records=@(Get-Slice4beActivityFiles $Fixture|Where-Object {-not $before.ContainsKey($_)}|ForEach-Object {Get-Content $_ -Raw|ConvertFrom-Json})
            $expected=if($case[1] -ceq 'Reenter'){2}elseif($case[0] -like 'Outcome*' -or $case[1] -ceq 'FinishSignOut'){1}else{0}
            $honest=$records.Count -eq $expected;$safe=$true
            foreach($record in $records){
                $terminal=$case[1] -ceq 'Reenter' -and $record.OutcomeCode -ceq 'APPLIED'
                $honest=$honest -and $record.ControlId -ceq 'ADMIN_DEMO_SEED' -and $record.UserId -ceq 'config-admin' -and $record.WarehouseId -ceq $Fixture.Warehouse -and ($record.OutcomeCode -ceq 'REQUESTED' -or $terminal)
                if(-not $terminal){$honest=$honest -and @($record.SourceEventRefs).Count -eq 0}
                $wire=$record|ConvertTo-Json -Depth 20 -Compress
                $honest=$honest -and ([bool](Run 'invSys.Core.xlam' 'TestSeedRecording.Concludes' @($wire)) -eq $terminal)
                foreach($value in @($Fixture.Root,$Fixture.Secret,(CredentialHash $Fixture.Secret),$sku,'unavailable.pending','PinHash')){$safe=$safe -and -not $wire.Contains($value)}
            }
            Check ($prefix+'.HonestRecordsWithoutFalseConclusion') $honest
            if($case[1] -ceq 'Reenter'){
                $requests=@($records|Where-Object OutcomeCode -CEQ 'REQUESTED');$outcomes=@($records|Where-Object OutcomeCode -CEQ 'APPLIED')
                $eventId=[string](Run 'invSys.Admin.xlam' 'TestSeedOutcome.ExactSeedEventForTest')
                $pair=$requests.Count -eq 1 -and $outcomes.Count -eq 1
                if($pair){$refs=@($outcomes[0].SourceEventRefs);$pair=$requests[0].ActivityId -ceq $outcomes[0].ActivityId -and $refs.Count -eq 1;if($pair){$pair=$refs[0].EventId -ceq $eventId -and $refs[0].WarehouseId -ceq $Fixture.Warehouse}}
                Check ($prefix+'.OuterPairAndExactSourceRetained') $pair
            }
            Check ($prefix+'.NoPrivatePayload') $safe
            $immutable=$true;foreach($file in $before.Keys){$immutable=$immutable -and (Get-FileHash -LiteralPath $file).Hash -ceq $before[$file]}
            Check ($prefix+'.PriorActivityImmutable') $immutable
            Check ($prefix+'.BothAuthoritiesPreserved') ((Get-FileHash $auth).Hash -ceq $authHash -and (Get-FileHash $otherAuth).Hash -ceq $otherHash)
            Check ($prefix+'.NoPendingArtifacts') (@(Get-ChildItem (Join-Path $Fixture.Root 'Training/Activity') -Recurse -File -Filter '*.pending').Count -eq 0)
            [void](Run 'invSys.Admin.xlam' 'TestSeedOutcome.Release')
        }
    }finally{
        [void](Run 'invSys.Core.xlam' 'TestSeedBoundary.Arm' @('',''))
        [void](Run 'invSys.Admin.xlam' 'TestSeedOutcome.Release')
        SelectTarget $Fixture
    }
}
