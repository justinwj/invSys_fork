# Generated local recipe input; the real worksheet owner and Scale handler remain intact.
function Install-ProductionRunScaleSuccessProbe {
    $project=$packages['invSys.Operations.xlam'].VBProject
    $form=$project.VBComponents.Item('frmProduction').CodeModule
    $form.InsertLines(1,"Private mScaleSuccessColumns As Object`r`nPrivate mScaleSuccessPhase As String, mScaleSuccessError As Long`r`nPrivate mScaleSuccessFactDetail As String")
    $form.AddFromString(@'
Public Function RunScaleSuccessPrepareForTest(ByVal canary As String) As Boolean
    On Error GoTo Unavailable
    Dim lo As ListObject, candidate As ListObject, ws As Worksheet
    Dim row As Long, name As Variant, header As Variant, cell As Range
    mScaleSuccessPhase = RunScaleSheetStageForTest("Reload", canary)
    If mScaleSuccessPhase <> "READY" Then Exit Function
    mScaleSuccessPhase = "RecipeRows"
    Set lo = mOperatorWorkbook.Worksheets("Recipes").ListObjects("Recipes")
    For row = 1 To 2
        If lo.ListRows.Count < row Then lo.ListRows.Add AlwaysInsert:=False
        lo.ListColumns("RECIPE_ID").DataBodyRange.Cells(row, 1).Value2 = mReadRecipeIdForTest
        lo.ListColumns("RECIPE").DataBodyRange.Cells(row, 1).Value2 = "Scale worksheet fixture"
        lo.ListColumns("PROCESS").DataBodyRange.Cells(row, 1).Value2 = "MIX"
        lo.ListColumns("INPUT/OUTPUT").DataBodyRange.Cells(row, 1).Value2 = IIf(row = 1, "INPUT", "OUTPUT")
        lo.ListColumns("INGREDIENT").DataBodyRange.Cells(row, 1).Value2 = "DEMO-RAW-BLACK-TEA"
        lo.ListColumns("INGREDIENT_ID").DataBodyRange.Cells(row, 1).Value2 = "A02"
        lo.ListColumns("PERCENT").DataBodyRange.Cells(row, 1).Value2 = 100#
        lo.ListColumns("AMOUNT").DataBodyRange.Cells(row, 1).Value2 = 2#
        lo.ListColumns("UOM").DataBodyRange.Cells(row, 1).Value2 = mRunSheetUom
        lo.ListColumns("RECIPE_LIST_ROW").DataBodyRange.Cells(row, 1).Value2 = row
    Next row
    mScaleSuccessPhase = "PaletteRows"
    Set lo = Nothing
    For Each ws In mOperatorWorkbook.Worksheets
        For Each candidate In ws.ListObjects
            If candidate.Name = "IngredientPalette" Then Set lo = candidate
        Next candidate
    Next ws
    If lo Is Nothing Then Exit Function
    If lo.ListRows.Count = 0 Then lo.ListRows.Add AlwaysInsert:=False
    lo.ListColumns("RECIPE_ID").DataBodyRange.Cells(1, 1).Value2 = mReadRecipeIdForTest
    lo.ListColumns("INGREDIENT_ID").DataBodyRange.Cells(1, 1).Value2 = "A02"
    lo.ListColumns("System_Key").DataBodyRange.Cells(1, 1).Value2 = mRunSheetKey
    lo.ListColumns("ITEM_CODE").DataBodyRange.Cells(1, 1).Value2 = "DEMO-RAW-BLACK-TEA"
    lo.ListColumns("ITEM").DataBodyRange.Cells(1, 1).Value2 = "DEMO-RAW-BLACK-TEA"
    lo.ListColumns("UOM").DataBodyRange.Cells(1, 1).Value2 = mRunSheetUom
    Set mScaleSuccessColumns = CreateObject("Scripting.Dictionary")
    mScaleSuccessPhase = "CustomColumns"
    For Each name In Array("RecipeChooser_generated", "InventoryPalette_generated", "ProductionOutput")
        Set lo = mOperatorWorkbook.Worksheets("Production").ListObjects(CStr(name))
        If lo.ListRows.Count = 0 Then lo.ListRows.Add AlwaysInsert:=False
        For Each header In Array("Operator Extra", "Operator Formula")
            If ProductionColumnIndex(lo, CStr(header)) = 0 Then lo.ListColumns.Add.Name = CStr(header)
            Set cell = lo.ListColumns(CStr(header)).DataBodyRange.Cells(1, 1)
            If CStr(header) = "Operator Extra" Then cell.Value2 = canary Else cell.Formula = "=1+2"
            mScaleSuccessColumns.Add CStr(name) & "|" & CStr(header), Array(lo.ListColumns(CStr(header)).Index, cell.Address, cell.Formula)
        Next header
    Next name
    mScaleSuccessPhase = "LocalRead"
    RunScaleSuccessPrepareForTest = RunScaleLocalRecipeExistsForTest()
    Exit Function
Unavailable:
    mScaleSuccessError = Err.Number
End Function
Public Function RunScaleSuccessSetupFactsForTest() As String
    RunScaleSuccessSetupFactsForTest = mScaleSuccessPhase & "|" & CStr(mScaleSuccessError)
End Function
Public Function RunScaleSuccessFactForTest(ByVal fact As String) As Boolean
    On Error GoTo Unavailable
    mScaleSuccessFactDetail = "Unavailable"
    Dim lo As ListObject, column As ListColumn, entry As Variant, names As Variant
    Dim name As Variant, snapshot As Variant, row As Long, matches As Long
    If fact = "ScaledChooser" Then
        Set lo = mOperatorWorkbook.Worksheets("Production").ListObjects("RecipeChooser_generated")
        If lo.ListRows.Count <> 2 Then Exit Function
        For row = 1 To 2
            If CStr(lo.ListColumns("PROCESS").DataBodyRange.Cells(row, 1).Value2) <> "MIX" Then Exit Function
            If Abs(CDbl(lo.DataBodyRange.Cells(row, ProductionColumnIndex(lo, "AMOUNT NEEDED")).Value2) - 4#) > 0.00000001 Then Exit Function
        Next row
    ElseIf fact = "ExactInputIdentity" Then
        For Each lo In mOperatorWorkbook.Worksheets("Production").ListObjects
            If Right$(lo.Name, 8) = "_palette" Then
                For row = 1 To lo.ListRows.Count
                    If StrComp(CStr(lo.ListColumns("System_Key").DataBodyRange.Cells(row, 1).Value2), mRunSheetKey, vbBinaryCompare) = 0 Then matches = matches + 1
                Next row
            End If
        Next lo
        If matches <> 1 Then Exit Function
    ElseIf fact = "OutputPrepared" Then
        Set lo = mOperatorWorkbook.Worksheets("Production").ListObjects("ProductionOutput")
        If CStr(lo.ListColumns("PROCESS").DataBodyRange.Cells(1, 1).Value2) <> "MIX" Then Exit Function
        If CStr(lo.ListColumns("OUTPUT").DataBodyRange.Cells(1, 1).Value2) <> "DEMO-RAW-BLACK-TEA" Then Exit Function
        If CStr(lo.ListColumns("System_Key").DataBodyRange.Cells(1, 1).Value2) <> "" Then Exit Function
    Else
        For Each entry In mScaleSuccessColumns.Keys
            names = Split(CStr(entry), "|")
            If CStr(names(0)) = fact Then
                snapshot = mScaleSuccessColumns(entry)
                mScaleSuccessFactDetail = "Missing table"
                Set lo = mOperatorWorkbook.Worksheets("Production").ListObjects(fact)
                mScaleSuccessFactDetail = "Missing column"
                Set column = lo.ListColumns(CStr(names(1)))
                mScaleSuccessFactDetail = "Column moved"
                If column.Index <> CLng(snapshot(0)) Then Exit Function
                mScaleSuccessFactDetail = "Content changed"
                row = 1
                If UBound(names) > 1 Then row = CLng(names(2))
                If column.DataBodyRange.Cells(row, 1).Formula <> CStr(snapshot(2)) Then Exit Function
                matches = matches + 1
            End If
        Next entry
        If matches < 2 Then Exit Function
    End If
    RunScaleSuccessFactForTest = True
    mScaleSuccessFactDetail = "Confirmed"
Unavailable:
End Function
Public Function RunScaleSuccessDetailForTest() As String
    RunScaleSuccessDetailForTest = mScaleSuccessFactDetail
End Function
Public Function RunScaleSuccessRepeatForTest(ByVal canary As String) As Boolean
    On Error GoTo Unavailable
    Dim lo As ListObject, name As Variant, header As Variant, row As Long, key As String
    Dim entry As Variant, names As Variant, cell As Range
    mScaleSuccessPhase = "RepeatColumns": mScaleSuccessError = 0
    For Each name In Array("RecipeChooser_generated", "proc_1_palette")
        Set lo = mOperatorWorkbook.Worksheets("Production").ListObjects(CStr(name))
        lo.ListColumns.Add(2).Name = "Workbook Note"
        lo.ListColumns.Add(5).Name = "Workbook Calculation"
        For row = 1 To lo.ListRows.Count
            lo.ListColumns("Workbook Note").DataBodyRange.Cells(row, 1).Value2 = canary & CStr(row)
            lo.ListColumns("Workbook Calculation").DataBodyRange.Cells(row, 1).Formula = "=2+" & CStr(row)
            For Each header In Array("Workbook Note", "Workbook Calculation")
                Set cell = lo.ListColumns(CStr(header)).DataBodyRange.Cells(row, 1)
                key = CStr(name) & "|" & CStr(header) & "|" & CStr(row)
                mScaleSuccessColumns.Add key, Array(lo.ListColumns(CStr(header)).Index, cell.Address, cell.Formula)
            Next header
        Next row
    Next name
    mOperatorWorkbook.Worksheets("Production").ListObjects("RecipeChooser_generated").ListColumns("AMOUNT NEEDED").Name = " amount needed "
    For Each entry In mScaleSuccessColumns.Keys
        names = Split(CStr(entry), "|"): row = 1
        If UBound(names) > 1 Then row = CLng(names(2))
        Set lo = mOperatorWorkbook.Worksheets("Production").ListObjects(CStr(names(0)))
        Set cell = lo.ListColumns(CStr(names(1))).DataBodyRange.Cells(row, 1)
        mScaleSuccessColumns(entry) = Array(lo.ListColumns(CStr(names(1))).Index, cell.Address, cell.Formula)
    Next entry
    TestProductionDesigner.RunScaleReset
    RunScaleSuccessRepeatForTest = True
    Exit Function
Unavailable:
    mScaleSuccessError = Err.Number
End Function
'@)
    $project.VBComponents.Item('TestProductionDesigner').CodeModule.AddFromString(@'
Public Function RunScaleSuccessPrepare(ByVal canary As String) As Boolean
    RunScaleSuccessPrepare = mForm.RunScaleSuccessPrepareForTest(canary)
End Function
Public Function RunScaleSuccessFact(ByVal fact As String) As Boolean
    RunScaleSuccessFact = mForm.RunScaleSuccessFactForTest(fact)
End Function
Public Function RunScaleSuccessSetupFacts() As String
    RunScaleSuccessSetupFacts = mForm.RunScaleSuccessSetupFactsForTest()
End Function
Public Function RunScaleSuccessDetail() As String
    RunScaleSuccessDetail = mForm.RunScaleSuccessDetailForTest()
End Function
Public Function RunScaleSuccessRepeat(ByVal canary As String) As Boolean
    RunScaleSuccessRepeat = mForm.RunScaleSuccessRepeatForTest(canary)
End Function
'@)
}

function Test-ProductionRunScaleSuccess($Fixture) {
    $before=@(Files)
    if(-not [bool](Probe 'RunScaleSuccessPrepare' @($canary))){throw ('Successful local recipe fixture unavailable: '+[string](Probe 'RunScaleSuccessSetupFacts')+'; not behavioral RED.')}
    foreach($name in @('RecipeChooser_generated','InventoryPalette_generated','ProductionOutput')){
        if(-not [bool](Probe 'RunScaleSuccessFact' @($name))){throw 'Custom-column precondition failed; not behavioral RED.'}
    }
    Check 'RunScaleSuccess.SetupNotUserAction' (@(Files).Count -eq $before.Count)
    Check 'RunScaleSuccess.ActualWorksheetBranch' ([bool](Probe 'RunScaleSheetBranch'))
    Check 'RunScaleSuccess.BothFlagsDisabled' (-not [bool](Run 'invSys.Core.xlam' 'modConfig.GetBool' @('DesignsEnabled',$true)) -and -not [bool](Run 'invSys.Core.xlam' 'modConfig.GetBool' @('FF_DesignsEnabled',$true)))
    $decoy.Activate();$notice=[string](Probe 'RunFaultAct' @('SCALE'))
    Check 'RunScaleSuccess.NoUnhandledError' (-not $notice.StartsWith('HANDLER_ERROR|'))
    Check 'RunScaleSuccess.GuardsRestored' ([bool](Probe 'RunFaultGuards'))
    Check 'RunScaleSuccess.RealCapturedOwnerWithoutFaultInjection' ([string](Probe 'RunScaleEvidence') -ceq '1|0|False|1|Captured|0' -and [string](Probe 'RunFaultEvidence') -ceq '0|False|False')
    Check 'RunScaleSuccess.RebuildAndPostLoadWorkReached' ([bool](Probe 'RunScaleRebuildReached') -and [int](Probe 'RunScalePostLoadCount') -eq 4)
    $details=@()
    foreach($name in @('ScaledChooser','ExactInputIdentity','OutputPrepared','RecipeChooser_generated','InventoryPalette_generated','ProductionOutput')){
        $passed=[bool](Probe 'RunScaleSuccessFact' @($name))
        Check ('RunScaleSuccess.'+$name) $passed
        $details+=[pscustomobject]@{Fact=$name;Passed=$passed;Detail=[string](Probe 'RunScaleSuccessDetail')}
    }
    $details|ConvertTo-Json|Set-Content (Join-Path $reportRoot 'scale-success-facts.json')
    Check 'RunScaleSuccess.ExistingSuccessMessage' ($notice.StartsWith('Loaded recipe at 200.% and prepared output rows: '))
    Check 'RunScaleSuccess.CanonicalIdentityAndValuesPreserved' ([bool](Probe 'RunWorksheetSource'))
    Pair $before 'SCALE' 'STAGED' 'RunScaleSuccess.Recorded'
    if([int](Probe 'RunLocalShowAndCapture' @($book.Name,'SCALE')) -ne 3){throw 'Scale result page unavailable.'}
    CaptureOwnedFormByCaptionEvidence 'Production' 'run-scale-success.png'
    $before=@(Files)
    if(-not [bool](Probe 'RunScaleSuccessRepeat' @($canary))){throw ('Repeated Scale fixture unavailable: '+[string](Probe 'RunScaleSuccessSetupFacts')+'; not behavioral RED.')}
    Check 'RunScaleRepeat.InputChangesAreNotActions' (@(Files).Count -eq $before.Count)
    $decoy.Activate();$notice=[string](Probe 'RunFaultAct' @('SCALE'))
    Check 'RunScaleRepeat.ExistingSuccessAndGuards' ($notice.StartsWith('Loaded recipe at 200.% and prepared output rows: ') -and [bool](Probe 'RunFaultGuards'))
    Check 'RunScaleRepeat.RealCapturedOwnerWithoutFaultInjection' ([string](Probe 'RunScaleEvidence') -ceq '1|0|False|1|Captured|0' -and [string](Probe 'RunFaultEvidence') -ceq '0|False|False')
    foreach($name in @('ScaledChooser','ExactInputIdentity','OutputPrepared','RecipeChooser_generated','InventoryPalette_generated','ProductionOutput','proc_1_palette')){
        Check ('RunScaleRepeat.'+$name) ([bool](Probe 'RunScaleSuccessFact' @($name)))
    }
    Check 'RunScaleRepeat.CanonicalIdentityAndValuesPreserved' ([bool](Probe 'RunWorksheetSource'))
    Pair $before 'SCALE' 'STAGED' 'RunScaleRepeat.Recorded'
    if([int](Probe 'RunLocalShowAndCapture' @($book.Name,'SCALE')) -ne 3){throw 'Repeated Scale result page unavailable.'}
    CaptureOwnedFormByCaptionEvidence 'Production' 'run-scale-success-repeat.png'
}
