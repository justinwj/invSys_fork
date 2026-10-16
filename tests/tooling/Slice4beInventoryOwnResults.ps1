# Operator Edit path with independent saved results. No synthetic canonical rows.
function Install-InventoryOwnResultProbe {
    $admin=$packages['invSys.Admin.xlam'].VBProject
    $admin.VBComponents.Item('TestInventoryCommands').CodeModule.AddFromString(@'
Public Function OwnResultShowsForTest() As Long
    OwnResultShowsForTest = mShows
End Function
Public Function OwnResultPartialForTest() As Boolean
    OwnResultPartialForTest = (InStr(1, mLastNotice, "Exact event application was saved:", vbTextCompare) > 0 _
        And InStr(1, mLastNotice, "Inventory quantity set.", vbTextCompare) = 0)
End Function
Public Function OwnResultReportFactsForTest() As String
    OwnResultReportFactsForTest = CStr(InStr(1, mLastNotice, "Partial result:", vbTextCompare) > 0) & "|" & _
        CStr(InStr(1, mLastNotice, "unsaved changes", vbTextCompare) > 0) & "|" & _
        CStr(InStr(1, mLastNotice, "read-only", vbTextCompare) > 0) & "|" & _
        CStr(InStr(1, mLastNotice, "source", vbTextCompare) > 0)
End Function
'@)
    $trace=$admin.VBComponents.Add(1);$trace.Name='TestInventoryCatalogTrace'
    $trace.CodeModule.AddFromString(@'
Option Explicit
Private mPhase As String
Public Sub Mark(ByVal phase As String)
    mPhase = phase
End Sub
Public Function Phase() As String
    Phase = mPhase
End Function
'@)
    $owner=$admin.VBComponents.Item('modAdmin').CodeModule
    $start=$owner.ProcStartLine('UpdateInventoryItemCatalogForWarehouse',0)
    $count=$owner.ProcCountLines('UpdateInventoryItemCatalogForWarehouse',0)
    $body=$owner.Lines($start,$count)
    foreach($point in @(
        @('    On Error GoTo FailUpdate','Entered'),
        @('        If Not modInventoryCommands.BeginCatalog(interactiveContext, sku, report) Then GoTo CleanExit','CatalogBegun'),
        @('        If Not modInventoryCommands.BindCatalogWriter(interactiveContext, wb.Name, report) Then GoTo CleanExit','WriterBound'),
        @('        Next field','FieldsPlanned'),
        @('        If Not modInventoryCommands.CheckForm(interactiveContext, report) Then GoTo CleanExit','BeforeWritePassed'),
        @('    WriteInventoryCatalogFieldsAdmin lo, rowIndex, fields','FieldsWritten'),
        @('        If Not modInventoryCommands.BeginCatalogSave(interactiveContext, report) Then GoTo CleanExit','BeforeSavePassed'),
        @('    wb.Save','SaveReturned'),
        @('    UpdateInventoryItemCatalogForWarehouse = True','Confirmed')
    )){
        # VBE may recase identifiers when it inserts source into a project.
        $pattern='(?im)^'+[regex]::Escape($point[0])+'[\t ]*\r?$'
        if([regex]::Matches($body,$pattern).Count -ne 1){throw ('Catalog fixed-stage anchor unavailable: '+$point[1])}
        $mark='    TestInventoryCatalogTrace.Mark "'+$point[1]+'"'
        $body=[regex]::Replace($body,$pattern,[Text.RegularExpressions.MatchEvaluator]{param($m) $m.Value.TrimEnd("`r")+"`r`n"+$mark})
    }
    $owner.DeleteLines($start,$count);$owner.InsertLines($start,$body)
    $core=$packages['invSys.Core.xlam'].VBProject
    $probe=$core.VBComponents.Add(1);$probe.Name='TestInventoryOwnResult'
    $probe.CodeModule.AddFromString(@'
Option Explicit
Private mSignOut As Boolean, mReturned As Boolean
Public Sub Arm(ByVal signOut As Boolean)
    mSignOut = signOut: mReturned = False
End Sub
Public Sub AfterProcessor()
    mReturned = True
    If mSignOut Then modAuth.SignOut
End Sub
Public Function Returned() As Boolean
    Returned = mReturned
End Function
'@)
    $module=$core.VBComponents.Item('modInventoryQuantityCommands').CodeModule
    $text=$module.Lines(1,$module.CountOfLines)
    $boundary='    processed = modProcessor.RunBatch(warehouse, 0, batchReport, eventId, applied, True)'
    if([regex]::Matches($text,[regex]::Escape($boundary)).Count -ne 1){throw 'Quantity processor-return test boundary unavailable.'}
    $module.DeleteLines(1,$module.CountOfLines)
    $module.AddFromString($text.Replace($boundary,($boundary+"`r`n"+'    TestInventoryOwnResult.AfterProcessor')))
}

function Test-InventoryOwnResults($Fixture,$Other) {
    SelectTarget $Fixture
    $sku='DEMO-OWN-RESULT'
    $csv=Join-Path $runRoot 'own-result-fixture.csv'
    @('ITEM_CODE,ITEM,UOM,LOCATION,QTY,CONDITION',($sku+',Own result fixture,LB,CLEARVIEW,10,GOOD'))|
        Set-Content $csv -Encoding ASCII
    $created=[bool](Run 'invSys.Admin.xlam' 'TestInventoryCommands.PrepareExactQuantityFixture' @($Fixture.Warehouse,$csv))
    Check 'InventoryOwnResult.GeneratedSingleEntityFixtureApplied' $created
    if(-not $created){throw 'Owning own-result fixture unavailable; not product RED.'}
    $original=Get-AdminInventoryEntryFacts $Fixture -Persisted
    $path=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Data.Inventory.xlsb')
    $book=$excel.Workbooks.Open($path,0,$true)
    try{
        $table=Table $book 'tblInventoryEntities';$entityKeys=@()
        foreach($row in $table.ListRows){
            if([string]$row.Range.Cells.Item(1,$table.ListColumns.Item('SKU').Index).Value2 -ceq $sku -and [string]$row.Range.Cells.Item(1,$table.ListColumns.Item('InventoryState').Index).Value2 -ceq 'ACTIVE'){
                $entityKeys+=[string]$row.Range.Cells.Item(1,$table.ListColumns.Item('System_Key').Index).Value2
            }
        }
    }finally{$book.Close($false)}
    if($entityKeys.Count -ne 1){throw 'Own-result SKU requires exactly one generated active entity.'}
    $key=$entityKeys[0]
    foreach($spec in @(@('NoOp',10),@('Applied',7),@('AuthorityLost',4))){
        $case=[string]$spec[0];$target=[double]$spec[1];$id='InventoryOwnResult.'+$case
        SelectTarget $Fixture
        $hostBook=$excel.Workbooks.Add();$hostBook.SaveAs((Join-Path $runRoot ('own-result-'+$case+'.xlsx')),51)
        try{
            $before=Get-AdminInventoryEntryFacts $Fixture -Persisted
            [void](Run 'invSys.Core.xlam' 'TestInventoryOwnResult.Arm' @(($case -ceq 'AuthorityLost')))
            [void](Run 'invSys.Admin.xlam' 'TestInventoryCommands.SetSelection' @($Other.Root,$sku))
            [void](Run 'invSys.Admin.xlam' 'TestInventoryCommands.SetQuantityTarget' @($target))
            [void](Run 'invSys.Admin.xlam' 'TestInventoryCommands.Execute' @($hostBook.Name,'Edit','SetQuantity','EditReason',$Fixture.Root,$Fixture.Warehouse,$Fixture.Secret))
            $facts=([string](Run 'invSys.Admin.xlam' 'TestInventoryCommands.Facts')).Split('|')
            $ready=$facts.Count -eq 5 -and $facts[0] -ceq 'True' -and $facts[1] -ceq 'True' -and $facts[2] -ceq '1'
            Check ($id+'.ActualEditHandlerAndReasonReached') $ready
            if(-not $ready){throw 'Own-result actual Edit handler unavailable; not product RED.'}
            $shows=[long](Run 'invSys.Admin.xlam' 'TestInventoryCommands.OwnResultShowsForTest')
            $catalogPhase=[string](Run 'invSys.Admin.xlam' 'TestInventoryCatalogTrace.Phase')
            if($catalogPhase -cnotin @('Entered','CatalogBegun','WriterBound','FieldsPlanned','BeforeWritePassed','FieldsWritten','BeforeSavePassed','SaveReturned','Confirmed')){throw 'Unknown catalog diagnostic stage.'}
            $noticeFacts=([string](Run 'invSys.Admin.xlam' 'TestInventoryCommands.OwnResultReportFactsForTest')).Split('|')
            [pscustomobject]@{Case=$case;CatalogPhase=$catalogPhase;FormShows=$shows;PartialNotice=($noticeFacts[0] -ceq 'True');DirtyNotice=($noticeFacts[1] -ceq 'True');ReadOnlyNotice=($noticeFacts[2] -ceq 'True');SourceNotice=($noticeFacts[3] -ceq 'True')}|ConvertTo-Json -Compress|Add-Content (Join-Path $reportRoot 'own-result-fixed-stages.jsonl')
            Check ($id+'.ExpectedFormContinuation') ($shows -eq $(if($case -ceq 'AuthorityLost'){1}else{2}))
            $after=Get-AdminInventoryEntryFacts $Fixture -Persisted
            $othersSame=@($original.EntityQuantities.Keys|Where-Object {$_ -cne $key -and (-not $after.EntityQuantities.ContainsKey($_) -or $after.EntityQuantities[$_] -ne $original.EntityQuantities[$_])}).Count -eq 0
            Check ($id+'.ExactEntityAndTargetSaved') ($after.EntityQuantities.Count -eq $original.EntityQuantities.Count -and $after.EntityQuantities.ContainsKey($key) -and $after.EntityQuantities[$key] -eq $target -and $othersSame)
            $events=if($case -ceq 'NoOp'){0}else{1}
            Check ($id+'.ExactEventCountSaved') ($after.LogRows -eq $before.LogRows+$events -and $after.AppliedRows -eq $before.AppliedRows+$events)
            $returned=[bool](Run 'invSys.Core.xlam' 'TestInventoryOwnResult.Returned')
            Check ($id+'.ExpectedProcessorReturn') ($returned -eq ($case -cne 'NoOp'))
            if($case -ceq 'AuthorityLost'){
                Check ($id+'.SavedPartialReported') ([bool](Run 'invSys.Admin.xlam' 'TestInventoryCommands.OwnResultPartialForTest'))
            }
        }finally{$hostBook.Close($false)}
    }
}
