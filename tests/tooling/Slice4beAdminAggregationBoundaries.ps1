# Faults are injected only into disposable packages/owner-created fixtures.
# Results retain booleans and fixed case names, never source values or paths.
function Install-AggregationBoundaryProbe {
    $core=$packages['invSys.Core.xlam'].VBProject
    $probe=$core.VBComponents.Add(1);$probe.Name='TestAggregationBoundary'
    $probe.CodeModule.AddFromString(@'
Option Explicit
Private mFault As String, mPoint As String, mInjected As Boolean
Private mSource As String, mAlternate As String, mFeed As String
Public Sub Arm(ByVal fault As String, ByVal point As String, ByVal source As String, ByVal alternate As String, ByVal feed As String)
    mFault = fault: mPoint = point: mSource = source: mAlternate = alternate: mFeed = feed: mInjected = False
End Sub
Public Function Injected() As Boolean
    Injected = mInjected
End Function
Public Sub AtPoint(ByVal point As String)
    Dim fso As Object, target As WarehouseTarget, report As String
    Dim auth As Workbook, sheet As Worksheet, table As ListObject, row As ListRow
    If mInjected Or mFault = "" Or point <> mPoint Then Exit Sub
    Select Case mFault
        Case "SIGN_OUT": modAuth.SignOut: mInjected = True
        Case "SOURCE_CHANGED"
            Set fso = CreateObject("Scripting.FileSystemObject")
            fso.CopyFile mAlternate, mSource, True
            mInjected = True
        Case "SOURCE_MISSING"
            Set fso = CreateObject("Scripting.FileSystemObject")
            fso.DeleteFile mSource, False
            mInjected = True
        Case "FEED_CHANGED"
            Set target = modNasConnection.GetCurrentTarget()
            If Not modConfig.UpdateConfigValue("PathSharePointRoot", mFeed, report, target.WarehouseId, target.StationId) Then Err.Raise vbObjectError + 501
            mInjected = True
        Case "RIGHTS_REVOKED"
            Set target = modNasConnection.GetCurrentTarget()
            Set auth = Application.Workbooks.Open(target.AuthPath, UpdateLinks:=0, ReadOnly:=False)
            For Each sheet In auth.Worksheets
                For Each table In sheet.ListObjects
                    If table.Name = "tblCapabilities" Then
                        For Each row In table.ListRows
                            If CStr(row.Range.Cells(1, table.ListColumns("UserId").Index).Value2) = modAuth.GetCurrentUserId() And _
                               CStr(row.Range.Cells(1, table.ListColumns("Capability").Index).Value2) = "ADMIN_MAINT" Then _
                                row.Range.Cells(1, table.ListColumns("Status").Index).Value2 = "Disabled"
                        Next row
                    End If
                Next table
            Next sheet
            auth.Save: auth.Close False: mInjected = True
        Case "WRITE_FAILURE": mInjected = True: Err.Raise 1004
    End Select
End Sub
Public Function DirectInteractive(ByVal context As String, ByVal sources As String, ByVal outputPath As String) As Boolean
    Dim report As String
    ' Baseline comparison: this is the exact primitive Core writer dispatched by
    ' the current interactive handler. It has no captured-context argument/guard.
    ' Scheduled/fixture authorization is a separate preserved comparison.
    DirectInteractive = modHqAggregator.GenerateGlobalSnapshotFromFiles(sources, outputPath, report)
End Function
'@)
    $owner=$packages['invSys.Admin.xlam'].VBProject.VBComponents.Item('modAdminConsole').CodeModule
    $first=$owner.ProcStartLine('RunHQAggregationFromSourceSet',0);$count=$owner.ProcCountLines('RunHQAggregationFromSourceSet',0)
    $body=$owner.Lines($first,$count)
    $marker='    TestAggregation.Phase "OutputConfig"'
    if(-not $body.Contains($marker)){throw 'Post-inspection fault boundary unavailable.'}
    $body=$body.Replace($marker,'    TestAggregationBoundary.AtPoint "AFTER_INSPECTION"'+"`r`n"+$marker)
    $owner.DeleteLines($first,$count);$owner.InsertLines($first,$body)
    $writer=$core.VBComponents.Item('modHqAggregator').CodeModule
    $first=$writer.ProcStartLine('WriteGlobalSnapshotWorkbook',0);$count=$writer.ProcCountLines('WriteGlobalSnapshotWorkbook',0)
    $body=$writer.Lines($first,$count)
    $marker='    TestAggregationRoots.Stage "SaveAs"'
    if(-not $body.Contains($marker)){throw 'Writer fault boundary unavailable.'}
    $body=$body.Replace($marker,'    TestAggregationBoundary.AtPoint "BEFORE_SAVE"'+"`r`n"+$marker)
    $writer.DeleteLines($first,$count);$writer.InsertLines($first,$body)
}

function Test-AggregationBoundaries($Fixture,$Other,$HostBook,$OtherBook,[ref]$Observations) {
    function Call([string]$Name,[object[]]$Values=@()){Run 'invSys.Admin.xlam' ('TestAggregation.'+$Name) $Values}
    function Core([string]$Name,[object[]]$Values=@()){Run 'invSys.Core.xlam' ('TestAggregationBoundary.'+$Name) $Values}
    function FileHash([string]$Path){if(-not(Test-Path -LiteralPath $Path)){return ''};$stream=[IO.File]::Open($Path,'Open','Read',([IO.FileShare]::ReadWrite -bor [IO.FileShare]::Delete));try{(Get-FileHash -InputStream $stream).Hash}finally{$stream.Dispose()}}
    $owned=[IO.Path]::GetFullPath($runRoot).TrimEnd('\')+'\'
    $sourceBytes=[IO.File]::ReadAllBytes($Fixture.Snapshot)
    $configBytes=[IO.File]::ReadAllBytes($Fixture.Config)
    $authPath=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Auth.xlsb')
    $authBytes=[IO.File]::ReadAllBytes($authPath)
    $feed=Split-Path -Parent (Split-Path -Parent $Fixture.Output)
    $changedFeed=Join-Path $runRoot 'changed-aggregation-feed'
    if(-not [IO.Path]::GetFullPath($changedFeed).StartsWith($owned,[StringComparison]::OrdinalIgnoreCase)){throw 'Changed feed escapes disposable scope.'}
    try{
        foreach($case in @('SourceChangedAfterInspection','SourceMissingAfterInspection','FeedChangedAfterInspection','SignedOutAfterInspection','RightsRevokedAfterInspection','FailureBeforeSave','BorrowedCleanOutput','BorrowedDirtyOutput','ReadOnlyOutput','DirtyConfig','UnreadableConfig','DirtyAuth')){
            SelectTarget $Fixture
            [void](Core 'Arm' @('','','','',''))
            [void](Run 'invSys.Core.xlam' 'TestAggregationRoots.PathComparison' @($true,$Fixture.Output))
            if(-not [bool](Run 'invSys.Core.xlam' 'TestAggregationRoots.NormalizedControl' @($Fixture.Snapshot,$Fixture.Output))){throw 'Existing output fixture failed; not behavioral RED.'}
            $before=FileHash $Fixture.Output
            $borrowed=$null;$authorityBook=$null
            $outputAttributes=[IO.File]::GetAttributes($Fixture.Output)
            if($case -cin @('BorrowedCleanOutput','BorrowedDirtyOutput')){
                $borrowed=$excel.Workbooks.Open($Fixture.Output,0,$false)
                if($case -ceq 'BorrowedDirtyOutput'){$borrowed.Worksheets.Item(1).Cells.Item(1,12).Value2='Preserve disposable output edit'}
            }
            if(-not [bool](Call 'OpenForm' @($HostBook.Name)) -or [int](Call 'SelectSource' @($Fixture.Warehouse)) -ne 1){throw 'Actual aggregation entry unavailable.'}
            if($case -ceq 'ReadOnlyOutput'){[IO.File]::SetAttributes($Fixture.Output,($outputAttributes -bor [IO.FileAttributes]::ReadOnly))}
            if($case -cin @('DirtyConfig','DirtyAuth')){
                $authorityPath=if($case -ceq 'DirtyConfig'){$Fixture.Config}else{$authPath}
                $authorityBook=$excel.Workbooks.Open($authorityPath,0,$false)
                $authorityBook.Worksheets.Item(1).Cells.Item(1,30).Value2='Preserve disposable authority edit'
                Check ('Aggregation.Boundary.'+$case+'.DirtyPrecondition') (-not [bool]$authorityBook.Saved)
            }
            if($case -ceq 'UnreadableConfig'){[IO.File]::WriteAllBytes($Fixture.Config,[Text.Encoding]::ASCII.GetBytes('Invalid disposable configuration'))}
            $fault=switch($case){
                'SourceChangedAfterInspection'{'SOURCE_CHANGED'}
                'SourceMissingAfterInspection'{'SOURCE_MISSING'}
                'FeedChangedAfterInspection'{'FEED_CHANGED'}
                'SignedOutAfterInspection'{'SIGN_OUT'}
                'RightsRevokedAfterInspection'{'RIGHTS_REVOKED'}
                'FailureBeforeSave'{'WRITE_FAILURE'}
                default{''}
            }
            $point=if($case -ceq 'FailureBeforeSave'){'BEFORE_SAVE'}else{'AFTER_INSPECTION'}
            [void](Core 'Arm' @($fault,$point,$Fixture.Snapshot,$Other.Snapshot,$changedFeed))
            [void](Call 'Aggregate')
            $success=[bool](Call 'OwnerSuccess');$calls=[int](Run 'invSys.Core.xlam' 'TestAggregationRoots.Calls')
            $preserved=(FileHash $Fixture.Output) -ceq $before
            if($fault){Check ('Aggregation.Boundary.'+$case+'.FaultInjected') ([bool](Core 'Injected'))}
            Check ('Aggregation.Boundary.'+$case+'.NoFalseCompletion') (-not $success)
            Check ('Aggregation.Boundary.'+$case+'.ExistingOutputPreserved') $preserved
            $redacted=[bool](Run 'invSys.Admin.xlam' 'frmAggregationSources.FeedbackRedactedForTest' @($Fixture.Root,$feed))
            Check ('Aggregation.Boundary.'+$case+'.FeedbackExcludesEnteredPaths') $redacted
            if($case -cne 'FailureBeforeSave'){
                Check ('Aggregation.Boundary.'+$case+'.RefusedBeforeWriter') ($calls -eq 0)
            }else{
                $retry=@(Get-ChildItem -LiteralPath (Split-Path -Parent $Fixture.Output) -Recurse -File -Filter '*aggregation*retry*.json' -ErrorAction SilentlyContinue)
                Check 'Aggregation.Boundary.FailureBeforeSave.DurableRetryStateRetained' ($retry.Count -gt 0)
                [void](Core 'Arm' @('','','','',''))
                [void](Call 'Aggregate')
                Check 'Aggregation.Boundary.FailureBeforeSave.SameFormRetryRequiresRecapture' (-not [bool](Call 'OwnerSuccess') -and [int](Run 'invSys.Core.xlam' 'TestAggregationRoots.Calls') -eq 0)
            }
            if($null -ne $borrowed){
                $stillOpen=$false;$statePreserved=$false
                foreach($book in $excel.Workbooks){if([string]$book.FullName -ceq $Fixture.Output){$stillOpen=$true;$statePreserved=([bool]$book.Saved -eq ($case -ceq 'BorrowedCleanOutput'));$book.Close($false);break}}
                Check ('Aggregation.Boundary.'+$case+'.BorrowedWorkbookStillOpen') $stillOpen
                Check ('Aggregation.Boundary.'+$case+'.BorrowedSavedStatePreserved') $statePreserved
                $borrowed=$null
            }
            if($null -ne $authorityBook){
                $stillOpen=$false;$dirty=$false
                foreach($book in $excel.Workbooks){if([string]$book.FullName -ceq $authorityPath){$stillOpen=$true;$dirty=-not [bool]$book.Saved;$book.Close($false);break}}
                Check ('Aggregation.Boundary.'+$case+'.AuthorityStillOpenDirty') ($stillOpen -and $dirty)
                $authorityBook=$null
            }
            $redirected=Test-Path -LiteralPath (Join-Path $changedFeed 'Global\invSys.Global.InventorySnapshot.xlsb')
            Check ('Aggregation.Boundary.'+$case+'.NoRedirectedOutput') (-not $redirected)
            $Observations.Value+=[pscustomobject]@{Case=$case;OwnerSuccess=$success;CoreDispatches=$calls;ExistingOutputPreserved=$preserved;RedirectedOutput=$redirected}
            [void](Call 'CloseForm')
            [void](Core 'Arm' @('','','','',''))
            if(Test-Path -LiteralPath $Fixture.Output){[IO.File]::SetAttributes($Fixture.Output,$outputAttributes)}
            [IO.File]::WriteAllBytes($Fixture.Snapshot,$sourceBytes)
            [IO.File]::WriteAllBytes($Fixture.Config,$configBytes)
            [IO.File]::WriteAllBytes($authPath,$authBytes)
            if($redirected){Remove-Item -LiteralPath (Join-Path $changedFeed 'Global\invSys.Global.InventorySnapshot.xlsb')}
        }
        SelectTarget $Fixture
        $closedHost=$excel.Workbooks.Add();$closedName=$closedHost.Name
        if(-not [bool](Call 'OpenForm' @($closedName)) -or [int](Call 'SelectSource' @($Fixture.Warehouse)) -ne 1){throw 'Closed-workbook fixture did not enter the actual handler.'}
        $closedHost.Close($false);$OtherBook.Activate()
        $otherSheets=$OtherBook.Worksheets.Count
        [void](Core 'Arm' @('','','','',''))
        [void](Run 'invSys.Core.xlam' 'TestAggregationRoots.PathComparison' @($true,$Fixture.Output))
        [void](Call 'Aggregate')
        Check 'Aggregation.Boundary.ClosedOriginalWorkbook.RefusedBeforeWriter' ([int](Run 'invSys.Core.xlam' 'TestAggregationRoots.Calls') -eq 0)
        Check 'Aggregation.Boundary.ClosedOriginalWorkbook.OtherWorkbookPreserved' ($OtherBook.Worksheets.Count -eq $otherSheets)
        [void](Call 'CloseForm')
        SelectTarget $Fixture
        $context=[string](Run 'invSys.Core.xlam' 'modAuthUserCommands.CaptureContext')
        $positive=Join-Path $runRoot 'direct-interactive-positive.xlsb'
        [void](Core 'Arm' @('','','','',''))
        Check 'Aggregation.DirectCore.BaselinePositiveWriterControl' ([bool](Core 'DirectInteractive' @($context,$Fixture.Snapshot,$positive)))
        [void](Run 'invSys.Core.xlam' 'modAuth.SignOut')
        $denied=Join-Path $runRoot 'direct-interactive-signed-out.xlsb'
        $wrote=[bool](Core 'DirectInteractive' @($context,$Fixture.Snapshot,$denied))
        Check 'Aggregation.DirectCore.SignedOutInteractiveDispatchDenied' (-not $wrote -and -not(Test-Path -LiteralPath $denied))
        $Observations.Value+=[pscustomobject]@{Case='DirectCoreSignedOut';OwnerSuccess=$wrote;SourceOutput=(Test-Path -LiteralPath $denied);BaselineAdapter=$true}
    }finally{
        [void](Call 'CloseForm');[void](Core 'Arm' @('','','','',''))
        [IO.File]::WriteAllBytes($Fixture.Snapshot,$sourceBytes)
        [IO.File]::WriteAllBytes($Fixture.Config,$configBytes)
        [IO.File]::WriteAllBytes($authPath,$authBytes)
        SelectTarget $Fixture
        [void](Run 'invSys.Core.xlam' 'TestAggregationRoots.PathComparison' @($false,$Fixture.Output))
    }
}
