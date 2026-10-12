# The released reusable Recipe selected by the actual form is not a Design/BOM record.
# Observe real owner reads and worksheet effects; never replace their results.
function Install-ProductionRunWorksheetScaleProbe {
    $project=$packages['invSys.Operations.xlam'].VBProject
    $owner=$project.VBComponents.Item('mProduction').CodeModule
    function Scale-Seam([string]$Anchor,[string]$Code,[bool]$After=$false){
        $start=$owner.ProcStartLine('LoadRecipeChooser',0);$end=$start+$owner.ProcCountLines('LoadRecipeChooser',0)
        # These anchors contain only VBA identifiers, with no string literals.
        $hits=@(for($i=$start;$i -lt $end;$i++){if($owner.Lines($i,1).Trim() -ieq $Anchor){$i}})
        if($hits.Count -ne 1){throw ('Worksheet Scale owner seam changed: '+$Anchor+'; not product RED.')}
        $owner.InsertLines($hits[0]+[int]$After,$Code)
    }
    $exitKind=if($owner.Lines($owner.ProcStartLine('LoadRecipeChooser',0),$owner.ProcCountLines('LoadRecipeChooser',0)) -match 'Public Function LoadRecipeChooser'){'Function'}else{'Sub'}
    Scale-Seam ('If wsProd Is Nothing Then Exit '+$exitKind) '    TestProductionDesigner.RunScaleOwnerResolved wsProd'
    if($RunScaleDesignsDisabledForTest){Scale-Seam 'TestProductionDesigner.RunScaleOwnerResolved wsProd' '    TestProductionDesigner.RunFaultBoundary "WORKSHEET_SCALE"' $true}
    Scale-Seam 'Set stagingWb = BuildReleasedDesignRecipeStagingWorkbook(recipeId, syncReport)' '        TestProductionDesigner.RunScaleReadReturned (Not stagingWb Is Nothing)' $true
    Scale-Seam 'If Not LocalProductionRecipeRowsExist(wsProd.Parent, recipeId) Then' '        TestProductionDesigner.RunScaleLegacyEntered'
    Scale-Seam 'DeleteRecipeChooserProcessTables wsProd' '    TestProductionDesigner.RunScaleRebuildEntered'
    $owner.AddFromString(@'
Public Function RunScaleLocalRecipeExistsForTest(ByVal wb As Workbook, ByVal recipeId As String) As Boolean
    RunScaleLocalRecipeExistsForTest = LocalProductionRecipeRowsExist(wb, recipeId)
End Function
'@)
    $start=$owner.ProcStartLine('LoadRecipeChooser',0);$end=$start+$owner.ProcCountLines('LoadRecipeChooser',0)
    $notices=0
    for($line=$start;$line -lt $end;$line++){
        $text=$owner.Lines($line,1)
        if($text.Trim().StartsWith('MsgBox ')){
            $owner.ReplaceLine($line,$text.Replace('MsgBox ','TestProductionDesigner.RunSheetNotice '))
            $notices++
        }
    }
    if($notices -ne 4){throw 'Worksheet Scale notification boundaries changed; not product RED.'}
    foreach($entry in @(
        @('modProductionReusableRun','LoadReleasedReusableRecipe','jsonText = modOperationsPrimitiveBridge.GetRecipeGraph(recipeId, recipeVersion)','Recipe'),
        @('modProductionRunDefinitionLoad','LoadNodeProcessDefinitions','modProductionReusableRun.RunRecordText(node, "ProcessId"), modProductionReusableRun.RunRecordText(node, "ProcessVersion"))','Process')
    )){
        $module=$project.VBComponents.Item($entry[0]).CodeModule
        $start=$module.ProcStartLine($entry[1],0);$end=$start+$module.ProcCountLines($entry[1],0)
        $hits=@(for($line=$start;$line -lt $end;$line++){if($module.Lines($line,1).Trim() -ieq $entry[2]){$line}})
        if($hits.Count -ne 1){throw 'Released Scale read observer seam changed; not product RED.'}
        $module.InsertLines($hits[0]+1,('    TestProductionDesigner.RunScaleReleasedRead "'+$entry[3]+'", (Len(jsonText) > 0)'))
    }
    $form=$project.VBComponents.Item('frmProduction').CodeModule
    # Observe each subsequent owner boundary without replacing its implementation.
    $start=$form.ProcStartLine('ApplyWorksheetRunScale',0);$end=$start+$form.ProcCountLines('ApplyWorksheetRunScale',0)
    $postLoadLines=@()
    foreach($anchor in @('If Not ApplyBatchScaleToRunList(scalePercent, scaleReport) Then Exit Sub','prepared = mProduction.PrepareProductionOutputForCurrentRecipe()','RefreshLoaderState','RefreshManagerState')){
        $hits=@(for($line=$start;$line -lt $end;$line++){if($form.Lines($line,1).Trim() -ceq $anchor){$line}})
        if($hits.Count -ne 1){throw 'Post-load Scale observation seam changed; not behavioral RED.'}
        $postLoadLines+=$hits[0]
    }
    foreach($line in @($postLoadLines|Sort-Object -Descending)){$form.InsertLines($line,'    TestProductionDesigner.RunScalePostLoadEntered')}
    $form.AddFromString(@'
Public Function RunScaleSheetStageForTest(ByVal mode As String, ByVal canary As String) As String
    Dim ws As Worksheet, lo As ListObject, name As Variant, row As Long, phase As String
    On Error GoTo Failed
    phase = "Restore"
    RunSheetOwnerRestoreForTest
    modProductionReusableRun.ClearReusableRun
    mLoading = True: mPages.Value = 3
    phase = "Selection"
    RefreshRecipeLists
    row = FindIdentityListRow(mLstLoaderRecipes, mReadRecipeIdForTest, mReadRecipeVersionForTest)
    If row < 0 Then GoTo Unavailable
    mLstLoaderRecipes.ListIndex = row
    mTxtBatchScalePercent.Text = "200"
    Select Case mode
        Case "NoSelection": mLstLoaderRecipes.ListIndex = -1
        Case "BadScale": mTxtBatchScalePercent.Text = "not a number"
        Case "BelowMinimum": mTxtBatchScalePercent.Text = "0"
        Case "AboveMaximum": mTxtBatchScalePercent.Text = "1001"
        Case "Minimum": mTxtBatchScalePercent.Text = "0.001"
        Case "Maximum": mTxtBatchScalePercent.Text = "1000"
        Case "LoadThenScale": mTxtBatchScalePercent.Text = "100"
    End Select
    Set ws = mOperatorWorkbook.Worksheets("Production")
    phase = "Staging"
    For Each name In Array("RecipeChooser_generated", "InventoryPalette_generated")
        Set lo = ws.ListObjects(CStr(name))
        If lo.ListRows.Count = 0 Then lo.ListRows.Add AlwaysInsert:=False
        If ProductionColumnIndex(lo, "Operator Extra") = 0 Then lo.ListColumns.Add.Name = "Operator Extra"
        If ProductionColumnIndex(lo, "Operator Formula") = 0 Then lo.ListColumns.Add.Name = "Operator Formula"
        lo.ListColumns("Operator Extra").DataBodyRange.Cells(1, 1).Value2 = canary
        lo.ListColumns("Operator Formula").DataBodyRange.Cells(1, 1).Formula = "=1+2"
        If CStr(name) = "RecipeChooser_generated" Then
            lo.ListColumns("AMOUNT NEEDED").DataBodyRange.Cells(1, 1).Value2 = 2#
        Else
            lo.ListColumns("BASE QUANTITY").DataBodyRange.Cells(1, 1).Value2 = 10#
            lo.ListColumns("QUANTITY").DataBodyRange.Cells(1, 1).Value2 = 2.5
            lo.ListColumns("System_Key").DataBodyRange.Cells(1, 1).Value2 = mRunSheetKey
        End If
    Next name
    If mode = "MissingSheet" Or mode = "DecoyMissing" Then ws.Name = "RunOwnerUnavailable"
    TestProductionDesigner.RunSheetReset mOperatorWorkbook.Name
    TestProductionDesigner.RunScaleReset
    mLoading = False
    RunScaleSheetStageForTest = "READY"
    Exit Function
Unavailable:
    RunScaleSheetStageForTest = "FIXTURE_FAILED|" & phase
    GoTo Finished
Failed:
    RunScaleSheetStageForTest = "FIXTURE_FAILED|" & phase & "|" & CStr(Err.Number)
Finished:
    mLoading = False
End Function
Public Function RunScaleSheetValuesForTest(ByVal mode As String, ByVal canary As String) As Boolean
    On Error GoTo Unavailable
    Dim ws As Worksheet, lo As ListObject, name As Variant, factor As Double
    If mode = "MissingSheet" Or mode = "DecoyMissing" Then
        Set ws = mOperatorWorkbook.Worksheets("RunOwnerUnavailable")
    Else
        Set ws = mOperatorWorkbook.Worksheets("Production")
    End If
    factor = 1#
    ' RUN-SCALE-01/02: refused or failed loading must not scale prior quantities.
    For Each name In Array("RecipeChooser_generated", "InventoryPalette_generated")
        Set lo = ws.ListObjects(CStr(name))
        If lo.ListColumns("Operator Extra").DataBodyRange.Cells(1, 1).Value2 <> canary Or _
           lo.ListColumns("Operator Formula").DataBodyRange.Cells(1, 1).Formula <> "=1+2" Then Exit Function
        If CStr(name) = "RecipeChooser_generated" Then
            If Abs(CDbl(lo.ListColumns("AMOUNT NEEDED").DataBodyRange.Cells(1, 1).Value2) - 2# * factor) > 0.00000001 Then Exit Function
        Else
            If Abs(CDbl(lo.ListColumns("BASE QUANTITY").DataBodyRange.Cells(1, 1).Value2) - 10# * factor) > 0.00000001 Then Exit Function
            If Abs(CDbl(lo.ListColumns("QUANTITY").DataBodyRange.Cells(1, 1).Value2) - 2.5 * factor) > 0.00000001 Then Exit Function
            If StrComp(CStr(lo.ListColumns("System_Key").DataBodyRange.Cells(1, 1).Value2), mRunSheetKey, vbBinaryCompare) <> 0 Then Exit Function
        End If
    Next name
    RunScaleSheetValuesForTest = True
Unavailable:
End Function
Public Function RunScaleSheetBranchForTest() As Boolean
    RunScaleSheetBranchForTest = Not modProductionReusableRun.ReusableRunIsLoaded() And mPages.Value = 3
End Function
Public Function RunScaleCustomValuesForTest(ByVal canary As String) As Boolean
    On Error GoTo Unavailable
    Dim lo As ListObject, name As Variant
    For Each name In Array("RecipeChooser_generated", "InventoryPalette_generated")
        Set lo = mOperatorWorkbook.Worksheets("Production").ListObjects(CStr(name))
        If lo.ListColumns("Operator Extra").DataBodyRange.Cells(1, 1).Value2 <> canary Then Exit Function
        If lo.ListColumns("Operator Formula").DataBodyRange.Cells(1, 1).Formula <> "=1+2" Then Exit Function
    Next name
    RunScaleCustomValuesForTest = True
Unavailable:
End Function
Public Function RunScaleLocalRecipeExistsForTest() As Boolean
    RunScaleLocalRecipeExistsForTest = mProduction.RunScaleLocalRecipeExistsForTest(mOperatorWorkbook, SelectedRunRecipeText(0))
End Function
Public Sub RunScaleInputForTest(ByVal value As String)
    mTxtBatchScalePercent.Text = value
End Sub
'@)
    $adapter=$project.VBComponents.Item('TestProductionDesigner').CodeModule
    $adapter.InsertLines(1,@'
Private mScaleLoadCalls As Long, mScaleReadCalls As Long, mScaleReadOK As Boolean, mScaleLegacyCalls As Long
Private mScaleOwnerTarget As String
Private mScaleRecipeReads As Long, mScaleProcessReads As Long
Private mScaleRebuildCalls As Long, mScalePostLoadCalls As Long
'@)
    $adapter.AddFromString(@'
Public Sub RunScaleReset()
    mScaleLoadCalls = 0: mScaleReadCalls = 0: mScaleReadOK = False: mScaleLegacyCalls = 0
    mScaleOwnerTarget = "NotEntered"
    mScaleRecipeReads = 0: mScaleProcessReads = 0
    mScaleRebuildCalls = 0
    mScalePostLoadCalls = 0
End Sub
Public Sub RunScaleReleasedRead(ByVal kind As String, ByVal available As Boolean)
    If Not available Then Exit Sub
    If kind = "Recipe" Then mScaleRecipeReads = mScaleRecipeReads + 1
    If kind = "Process" Then mScaleProcessReads = mScaleProcessReads + 1
End Sub
Public Function RunScaleReleasedReads() As String
    RunScaleReleasedReads = CStr(mScaleRecipeReads) & "|" & CStr(mScaleProcessReads)
End Function
Public Sub RunScaleInput(ByVal value As String)
    mForm.RunScaleInputForTest value
End Sub
Public Sub RunScaleOwnerResolved(ByVal ws As Worksheet)
    mScaleLoadCalls = mScaleLoadCalls + 1
    If ws Is Nothing Then
        mScaleOwnerTarget = "Unavailable"
    ElseIf ws.Parent.Name = mRunSheetExpectedBook Then
        mScaleOwnerTarget = "Captured"
    ElseIf ws.Parent.IsAddin Then
        mScaleOwnerTarget = "OtherAddin"
    Else
        mScaleOwnerTarget = "OtherWorkbook"
    End If
End Sub
Public Sub RunScaleReadReturned(ByVal succeeded As Boolean)
    mScaleReadCalls = mScaleReadCalls + 1: mScaleReadOK = succeeded
End Sub
Public Sub RunScaleLegacyEntered()
    mScaleLegacyCalls = mScaleLegacyCalls + 1
End Sub
Public Sub RunScaleRebuildEntered()
    mScaleRebuildCalls = mScaleRebuildCalls + 1
End Sub
Public Function RunScaleRebuildReached() As Boolean
    RunScaleRebuildReached = mScaleRebuildCalls = 1
End Function
Public Sub RunScalePostLoadEntered()
    mScalePostLoadCalls = mScalePostLoadCalls + 1
End Sub
Public Function RunScalePostLoadCount() As Long
    RunScalePostLoadCount = mScalePostLoadCalls
End Function
Public Function RunScaleEvidence() As String
    RunScaleEvidence = CStr(mScaleLoadCalls) & "|" & CStr(mScaleReadCalls) & "|" & CStr(mScaleReadOK) & "|" & CStr(mScaleLegacyCalls) & "|" & mScaleOwnerTarget & "|" & CStr(mRunSheetNotices)
End Function
Public Function RunScaleSheetStage(ByVal mode As String, ByVal canary As String) As String
    RunScaleSheetStage = mForm.RunScaleSheetStageForTest(mode, canary)
End Function
Public Function RunScaleSheetValues(ByVal mode As String, ByVal canary As String) As Boolean
    RunScaleSheetValues = mForm.RunScaleSheetValuesForTest(mode, canary)
End Function
Public Function RunScaleSheetBranch() As Boolean
    RunScaleSheetBranch = mForm.RunScaleSheetBranchForTest()
End Function
Public Function RunScaleCustomValues(ByVal canary As String) As Boolean
    RunScaleCustomValues = mForm.RunScaleCustomValuesForTest(canary)
End Function
Public Function RunScaleLocalRecipeExists() As Boolean
    RunScaleLocalRecipeExists = mForm.RunScaleLocalRecipeExistsForTest()
End Function
'@)
}

function Test-ProductionRunWorksheetScale($Fixture) {
    if(-not [bool](Probe 'RunSheetOwnerPrepare')){throw 'Supported Scale surface fixture unavailable; not product RED.'}
    if($RunScaleDesignsDisabledForTest){Test-ProductionRunScaleDisabled $Fixture;return}
    Check 'RunWorksheetScale.SupportedSurfaceAndSeedIdentity' $true
    $key=[string](Probe 'RunWorksheetKey')
    foreach($mode in @('NoSelection','BadScale','BelowMinimum','AboveMaximum','UnavailableDesign','Minimum','Maximum','MissingSheet','DecoyMissing')){
        if($mode -ceq 'DecoyMissing'){
            $decoySheet=$decoy.Worksheets.Add();$decoySheet.Name='Production'
            $decoySheet.Range('A1').Value2='Operator Extra';$decoySheet.Range('A2').Value2=$canary;$decoySheet.Range('B2').Formula='=1+2'
        }
        $before=@(Files);$ready=[string](Probe 'RunScaleSheetStage' @($mode,$canary))
        if($ready -cne 'READY'){
            if($ready -notmatch '^FIXTURE_FAILED\|[A-Za-z]+(\|-?[0-9]+)?$'){$ready='Unavailable'}
            throw ('Worksheet Scale staging unavailable: '+$ready+'; not product RED.')
        }
        try{
            $decoy.Activate();$label='RunWorksheetScale.'+$mode
            Check ($label+'.SetupNotUserAction') (@(Files).Count -eq $before.Count)
            Check ($label+'.ActualWorksheetBranch') ([bool](Probe 'RunScaleSheetBranch'))
            $notice=[string](Probe 'RunFaultAct' @('SCALE'))
            $evidence=([string](Probe 'RunScaleEvidence')).Split('|')
            if($evidence.Count -ne 6){throw 'Worksheet Scale owner evidence unavailable.'}
            $missing=$mode -cin @('MissingSheet','DecoyMissing')
            Check ($label+'.NoUnhandledError') (-not $notice.StartsWith('HANDLER_ERROR|'))
            Check ($label+'.GuardsRestoredWithoutAdapterReset') ([bool](Probe 'RunFaultGuards'))
            Check ($label+'.CurrentWorksheetValuesAndUnknownColumns') ([bool](Probe 'RunScaleSheetValues' @($mode,$canary)))
            Check ($label+'.CanonicalSourcePreserved') ([bool](Probe 'RunWorksheetSource'))
            Check ($label+'.NoLegacyFallback') ([int]$evidence[3] -eq 0)
            Check ($label+'.ApprovedLoadFirstMessage') ($notice -ceq 'Load a released recipe before applying scale.')
            Check ($label+'.RefusalBeforeOwner') ([int]$evidence[0] -eq 0 -and [int]$evidence[1] -eq 0 -and [int]$evidence[5] -eq 0)
            if($missing){
                Write-Output ('Run Scale '+$mode+' resolved target: '+$evidence[4])
                Check ($label+'.NoOtherWorkbookOwnerEntry') ($evidence[4] -cin @('NotEntered','Unavailable'))
            }
            if($mode -ceq 'DecoyMissing'){Check ($label+'.DecoyUnknownValuesPreserved') ($decoySheet.Range('A2').Value2 -ceq $canary -and $decoySheet.Range('B2').Formula -ceq '=1+2')}
            Pair $before 'SCALE' 'REJECTED' $label
            $records=@(Files|Where-Object{$_ -cnotin $before}|ForEach-Object{[IO.File]::ReadAllText($_)})
            Check ($label+'.NoExactKeyInObservations') ($records.Count -eq 2 -and @($records|Where-Object{$_.Contains($key)}).Count -eq 0)
            if($CaptureEvidence -and $mode -ceq 'UnavailableDesign'){
                Check ($label+'.VisibleOperatorPage') ([int](Probe 'RunLocalShowAndCapture' @($book.Name,'SCALE')) -eq 3)
                CaptureOwnedFormByCaptionEvidence 'Production' 'run-scale-load-first.png'
            }
        }finally{[void](Probe 'RunSheetOwnerRestore')}
    }
    $before=@(Files)
    if([string](Probe 'RunScaleSheetStage' @('LoadThenScale',$canary)) -cne 'READY'){throw 'Released Load/Scale sequence fixture unavailable; not product RED.'}
    Check 'RunScaleSequence.SetupNotUserAction' (@(Files).Count -eq $before.Count)
    $decoy.Activate();$notice=[string](Probe 'RunFaultAct' @('LOAD'))
    Check 'RunScaleSequence.ActualLoadStagesReleasedRun' (-not $notice.StartsWith('HANDLER_ERROR|') -and [bool](Probe 'RunLocalPreserved' @('LOAD','Normal')))
    Pair $before 'LOAD' 'STAGED' 'RunScaleSequence.Load'
    $reads=([string](Probe 'RunScaleReleasedReads')).Split('|')
    Check 'RunScaleSequence.ActualReleasedRecipeAndProcessReads' ($reads.Count -eq 2 -and [int]$reads[0] -eq 1 -and [int]$reads[1] -gt 0)
    $before=@(Files);[void](Probe 'RunScaleInput' @('200'))
    Check 'RunScaleSequence.InputIsNotAnExtraAction' (@(Files).Count -eq $before.Count)
    $notice=[string](Probe 'RunFaultAct' @('SCALE'))
    Check 'RunScaleSequence.ActualScaleRetainsLoadedRunSemantics' (-not $notice.StartsWith('HANDLER_ERROR|') -and [bool](Probe 'RunLocalPreserved' @('SCALE','Normal')))
    Pair $before 'SCALE' 'STAGED' 'RunScaleSequence.Scale'
    Check 'RunScaleSequence.NoWorksheetOrLegacyOwnerRead' ([string](Probe 'RunScaleEvidence') -ceq '0|0|False|0|NotEntered|0')
    Check 'RunScaleSequence.CustomColumnsAndExactKeyPreserved' ([bool](Probe 'RunScaleSheetValues' @('LoadThenScale',$canary)))
    Check 'RunScaleSequence.CanonicalSourcePreserved' ([bool](Probe 'RunWorksheetSource'))
    if($CaptureEvidence){
        Check 'RunScaleSequence.VisibleOperatorPage' ([int](Probe 'RunLocalShowAndCapture' @($book.Name,'SCALE')) -eq 3)
        CaptureOwnedFormByCaptionEvidence 'Production' 'run-scale-loaded-sequence.png'
    }
}

# Refusals and a caught owner exception; successful worksheet reload remains a separate gate.
function Test-ProductionRunScaleDisabled($Fixture) {
    if($RunScaleReloadForTest){Test-ProductionRunScaleReload $Fixture;return}
    Check 'RunScaleDisabled.SupportedSurfaceAndSeedIdentity' $true
    $key=[string](Probe 'RunWorksheetKey')
    foreach($mode in @('NoSelection','BadScale','BelowMinimum','AboveMaximum','MissingSheet','DecoyMissing','OwnerException')){
        if($mode -ceq 'DecoyMissing'){
            $decoySheet=$decoy.Worksheets.Add();$decoySheet.Name='Production'
            $decoySheet.Range('A1').Value2='Operator Extra';$decoySheet.Range('A2').Value2=$canary;$decoySheet.Range('B2').Formula='=1+2'
        }
        $before=@(Files);$ready=[string](Probe 'RunScaleSheetStage' @($mode,$canary))
        if($ready -cne 'READY'){throw 'Designs-disabled Scale staging unavailable; not product RED.'}
        try{
            $decoy.Activate();$label='RunScaleDisabled.'+$mode
            Check ($label+'.SetupNotUserAction') (@(Files).Count -eq $before.Count)
            Check ($label+'.ActualWorksheetBranch') ([bool](Probe 'RunScaleSheetBranch'))
            Check ($label+'.BothFlagsRemainDisabled') (-not [bool](Run 'invSys.Core.xlam' 'modConfig.GetBool' @('DesignsEnabled',$true)) -and -not [bool](Run 'invSys.Core.xlam' 'modConfig.GetBool' @('FF_DesignsEnabled',$true)))
            if($mode -ceq 'OwnerException'){[void](Probe 'RunFaultArm' @('Exception','SCALE','WORKSHEET_SCALE',$canary))}
            $notice=[string](Probe 'RunFaultAct' @('SCALE'))
            $evidence=[string](Probe 'RunScaleEvidence')
            Check ($label+'.NoUnhandledError') (-not $notice.StartsWith('HANDLER_ERROR|'))
            Check ($label+'.GuardsRestored') ([bool](Probe 'RunFaultGuards'))
            Check ($label+'.LocalValuesCustomColumnsAndExactKeyPreserved') ([bool](Probe 'RunScaleSheetValues' @($mode,$canary)))
            Check ($label+'.CanonicalSourcePreserved') ([bool](Probe 'RunWorksheetSource'))
            if($mode -ceq 'OwnerException'){
                Check ($label+'.RealOwnerCaughtException') ([string](Probe 'RunFaultEvidence') -ceq '1|True|False' -and $evidence -ceq '1|0|False|0|Captured|1')
                Check ($label+'.NoSubsequentScalePreparationOrRefresh') ([int](Probe 'RunScalePostLoadCount') -eq 0)
                Check ($label+'.ApprovedFailureMessage') ($notice -ceq 'Recipe could not be loaded. Scale was not applied.')
            }else{Check ($label+'.NoRecipeOwnerRead') ($evidence -ceq '0|0|False|0|NotEntered|0')}
            if($mode -ceq 'NoSelection'){$expected='Select a recipe, enter Batch scale %, then click Apply Scale.'}
            elseif($mode -ceq 'BadScale'){$expected='Batch scale must be a number from 0.001% through 1000%.'}
            elseif($mode -cin @('BelowMinimum','AboveMaximum')){$expected='Batch scale must be from 0.001% through 1000%.'}
            else{$expected=$null}
            if($null -ne $expected){Check ($label+'.ExistingRefusalMessage') ($notice -ceq $expected)}
            if($mode -ceq 'DecoyMissing'){Check ($label+'.DecoyUnknownValuesPreserved') ($decoySheet.Range('A2').Value2 -ceq $canary -and $decoySheet.Range('B2').Formula -ceq '=1+2')}
            $outcome=if($mode -cin @('MissingSheet','DecoyMissing','OwnerException')){'FAILED'}else{'REJECTED'}
            Pair $before 'SCALE' $outcome $label
            $records=@(Files|Where-Object{$_ -cnotin $before}|ForEach-Object{[IO.File]::ReadAllText($_)})
            Check ($label+'.NoExactKeyInObservations') ($records.Count -eq 2 -and @($records|Where-Object{$_.Contains($key)}).Count -eq 0)
            if($CaptureEvidence -and $mode -ceq 'NoSelection'){
                Check ($label+'.VisibleOperatorPage') ([int](Probe 'RunLocalShowAndCapture' @($book.Name,'SCALE')) -eq 3)
                CaptureOwnedFormByCaptionEvidence 'Production' 'run-scale-disabled-refusal.png'
            }
            if($CaptureEvidence -and $mode -ceq 'OwnerException'){
                if([int](Probe 'RunLocalShowAndCapture' @($book.Name,'SCALE')) -ne 3){throw 'Scale failure page unavailable for visible evidence.'}
                CaptureOwnedFormByCaptionEvidence 'Production' 'run-scale-disabled-owner-failure.png'
            }
        }finally{[void](Probe 'RunSheetOwnerRestore')}
    }
}

function Test-ProductionRunScaleReload($Fixture) {
    # Real released selection and generated surfaces; no fabricated legacy identity.
    $ready=[string](Probe 'RunScaleSheetStage' @('Reload',$canary))
    if($ready -cne 'READY' -or -not [bool](Probe 'RunScaleCustomValues' @($canary)) -or [bool](Probe 'RunScaleLocalRecipeExists')){throw 'Missing worksheet-definition preservation fixture unavailable; not behavioral RED.'}
    try {
        $decoy.Activate()
        Check 'RunScaleReload.ActualWorksheetBranch' ([bool](Probe 'RunScaleSheetBranch'))
        Check 'RunScaleReload.BothFlagsDisabled' (-not [bool](Run 'invSys.Core.xlam' 'modConfig.GetBool' @('DesignsEnabled',$true)) -and -not [bool](Run 'invSys.Core.xlam' 'modConfig.GetBool' @('FF_DesignsEnabled',$true)))
        $before=@(Files)
        $notice=[string](Probe 'RunFaultAct' @('SCALE'))
        $evidence=[string](Probe 'RunScaleEvidence')
        [pscustomobject]@{RebuildReached=[bool](Probe 'RunScaleRebuildReached');LocalRecipeExistsAfter=[bool](Probe 'RunScaleLocalRecipeExists');FaultInjected=$false}|ConvertTo-Json|Set-Content (Join-Path $reportRoot 'scale-reload-boundary.json')
        Check 'RunScaleReload.NoUnhandledError' (-not $notice.StartsWith('HANDLER_ERROR|'))
        Check 'RunScaleReload.GuardsRestored' ([bool](Probe 'RunFaultGuards'))
        Check 'RunScaleReload.RealCapturedOwnerWithoutFaultInjection' ($evidence.StartsWith('1|0|False|1|Captured|') -and [string](Probe 'RunFaultEvidence') -ceq '0|False|False')
        Check 'RunScaleReload.UnknownColumnsValuesAndFormulasPreserved' ([bool](Probe 'RunScaleCustomValues' @($canary)))
        Check 'RunScaleReload.PriorQuantitiesAndExactKeyPreserved' ([bool](Probe 'RunScaleSheetValues' @('Reload',$canary)))
        Check 'RunScaleReload.NoSubsequentScalePreparationOrRefresh' ([int](Probe 'RunScalePostLoadCount') -eq 0)
        Check 'RunScaleReload.ApprovedFailureMessage' ($notice -ceq 'Recipe could not be loaded. Scale was not applied.')
        Check 'RunScaleReload.CanonicalIdentityAndValuesPreserved' ([bool](Probe 'RunWorksheetSource'))
        Pair $before 'SCALE' 'FAILED' 'RunScaleReload.MissingWorksheetDefinition'
        if($CaptureEvidence){
            if([int](Probe 'RunLocalShowAndCapture' @($book.Name,'SCALE')) -ne 3){throw 'Scale result page unavailable for visible evidence.'}
            CaptureOwnedFormByCaptionEvidence 'Production' 'run-scale-missing-definition.png'
        }
    } finally { [void](Probe 'RunSheetOwnerRestore') }
}
