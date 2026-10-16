# Actual Delete handler; all inventory starts at the owning Admin CSV boundary.
function Install-InventoryDeleteResultProbe {
    $project=$packages['invSys.Admin.xlam'].VBProject
    $project.VBComponents.Item('TestInventoryCommands').CodeModule.AddFromString(@'
Public Function DeleteResultShowsForTest() As Long
    DeleteResultShowsForTest = mShows
End Function
Public Function DeleteResultPartialForTest() As Boolean
    DeleteResultPartialForTest = (InStr(1, mLastNotice, "partial", vbTextCompare) > 0 And _
        InStr(1, mLastNotice, "Event:", vbTextCompare) > 0 And _
        InStr(1, mLastNotice, "saved", vbTextCompare) > 0)
End Function
'@)
    $probe=$project.VBComponents.Add(1);$probe.Name='TestInventoryDeleteResult'
    $probe.CodeModule.AddFromString(@'
Option Explicit
Private mLose As Boolean, mReturned As Boolean
Public Sub Arm(ByVal lose As Boolean)
    mLose = lose: mReturned = False
End Sub
Public Sub AfterProcessor()
    mReturned = True
    If mLose Then modAuth.SignOut
End Sub
Public Function Returned() As Boolean
    Returned = mReturned
End Function
'@)
    $owner=$project.VBComponents.Item('modAdmin').CodeModule
    $start=$owner.ProcStartLine('RetireInventoryItemForWarehouse',0);$count=$owner.ProcCountLines('RetireInventoryItemForWarehouse',0)
    $body=$owner.Lines($start,$count)
    $pattern='(?im)^[\t ]*processedCount = modProcessor\.RunBatch\(warehouseId, 0, batchReport\)[\t ]*\r?$'
    if([regex]::Matches($body,$pattern).Count -ne 1){throw 'Delete processor-return observation boundary unavailable.'}
    $body=[regex]::Replace($body,$pattern,[Text.RegularExpressions.MatchEvaluator]{param($m) $m.Value.TrimEnd("`r")+"`r`n"+'        TestInventoryDeleteResult.AfterProcessor'})
    $owner.DeleteLines($start,$count);$owner.InsertLines($start,$body)
}

function Test-InventoryDeleteResults($Fixture,$Other) {
    foreach($case in @('Authorized','AuthorityLost')){
        SelectTarget $Fixture
        $sku='DEMO-DELETE-'+$case.ToUpperInvariant()
        $csv=Join-Path $runRoot ('delete-result-'+$case+'.csv')
        @('ITEM_CODE,ITEM,UOM,LOCATION,QTY,CONDITION',($sku+',Delete fixture,LB,CLEARVIEW,10,GOOD'),($sku+',Delete fixture,LB,CLEARVIEW,7,HOLD'))|Set-Content $csv -Encoding ASCII
        $created=[bool](Run 'invSys.Admin.xlam' 'TestInventoryCommands.PrepareExactQuantityFixture' @($Fixture.Warehouse,$csv))
        $id='InventoryDeleteResult.'+$case
        Check ($id+'.OwningGeneratedFixtureApplied') $created
        if(-not $created){throw 'Generated retirement fixture unavailable.'}
        $path=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Data.Inventory.xlsb')
        $keys=@();$book=$excel.Workbooks.Open($path,0,$false)
        try{
            $entities=Table $book 'tblInventoryEntities';$catalog=Table $book 'tblSkuCatalog'
            foreach($table in @($entities,$catalog)){
                $table.Parent.Unprotect()
                $extra=$null;try{$extra=$table.ListColumns.Item('DELETE UNKNOWN')}catch{}
                if($null -eq $extra){$extra=$table.ListColumns.Add();$extra.Name='DELETE UNKNOWN'}
                $extra.DataBodyRange.Value2='Preserved extension'
                $table.Parent.Protect()
            }
            foreach($row in $entities.ListRows){if([string]$row.Range.Cells.Item(1,$entities.ListColumns.Item('SKU').Index).Value2 -ceq $sku){$keys+=[string]$row.Range.Cells.Item(1,$entities.ListColumns.Item('System_Key').Index).Value2}}
            $book.Save()
        }finally{$book.Close($false)}
        Check ($id+'.TwoDistinctGeneratedEntities') ($keys.Count -eq 2 -and $keys[0] -cne $keys[1])
        if($keys.Count -ne 2){throw 'Exact all-active fixture unavailable.'}
        $before=Get-AdminInventoryEntryFacts $Fixture -Persisted
        $hostBook=$excel.Workbooks.Add();$hostBook.SaveAs((Join-Path $runRoot ('delete-result-'+$case+'.xlsx')),51)
        try{
            [void](Run 'invSys.Admin.xlam' 'TestInventoryDeleteResult.Arm' @(($case -ceq 'AuthorityLost')))
            [void](Run 'invSys.Admin.xlam' 'TestInventoryCommands.SetSelection' @($Other.Root,$sku))
            [void](Run 'invSys.Admin.xlam' 'TestInventoryCommands.Execute' @($hostBook.Name,'Delete','Authorized','DeleteConfirmation',$Fixture.Root,$Fixture.Warehouse,$Fixture.Secret))
            $facts=([string](Run 'invSys.Admin.xlam' 'TestInventoryCommands.Facts')).Split('|')
            $ready=$facts.Count -eq 5 -and $facts[0] -ceq 'True' -and $facts[1] -ceq 'True' -and $facts[2] -ceq '1'
            Check ($id+'.ActualDeleteAndConfirmationReached') $ready
            if(-not $ready){throw 'Actual Delete/confirmation unavailable; not product RED.'}
            Check ($id+'.OriginalOperatorCaptured') ($facts[3] -ceq 'True')
            Check ($id+'.ProcessorReturned') ([bool](Run 'invSys.Admin.xlam' 'TestInventoryDeleteResult.Returned'))
            Check ($id+'.ExpectedContinuation') ([long](Run 'invSys.Admin.xlam' 'TestInventoryCommands.DeleteResultShowsForTest') -eq $(if($case -ceq 'Authorized'){2}else{1}))
            $after=Get-AdminInventoryEntryFacts $Fixture -Persisted
            $exact=$after.AppliedRows -eq $before.AppliedRows+1 -and $after.LogRows -eq $before.LogRows+2 -and $after.EntityQuantities.Count -eq $before.EntityQuantities.Count
            foreach($key in $before.EntityQuantities.Keys){$expected=if($keys -ccontains $key){0}else{$before.EntityQuantities[$key]};if(-not $after.EntityQuantities.ContainsKey($key) -or $after.EntityQuantities[$key] -ne $expected){$exact=$false}}
            Check ($id+'.AllAndOnlyExactEntitiesRetired') $exact
            $hash=Get-AdminInventoryEntryHash $path;$book=$excel.Workbooks.Open($path,0,$true)
            try{
                $entities=Table $book 'tblInventoryEntities';$catalog=Table $book 'tblSkuCatalog';$states=0;$catalogState='';$extras=$true
                foreach($table in @($entities,$catalog)){
                    foreach($row in $table.ListRows){
                        if([string]$row.Range.Cells.Item(1,$table.ListColumns.Item('DELETE UNKNOWN').Index).Value2 -cne 'Preserved extension'){$extras=$false}
                        if([string]$row.Range.Cells.Item(1,$table.ListColumns.Item('SKU').Index).Value2 -ceq $sku){
                            if($table.Name -ceq 'tblInventoryEntities'){if([string]$row.Range.Cells.Item(1,$table.ListColumns.Item('InventoryState').Index).Value2 -ceq 'RETIRED'){$states++}}
                            else{$catalogState=[string]$row.Range.Cells.Item(1,$table.ListColumns.Item('CATALOG_STATE').Index).Value2}
                        }
                    }
                }
                Check ($id+'.ExactRetiredStatesSaved') ($states -eq 2)
                Check ($id+'.CatalogWriteFollowsAuthority') ($(if($case -ceq 'Authorized'){$catalogState -ceq 'RETIRED'}else{$catalogState -cne 'RETIRED'}))
                Check ($id+'.UnknownColumnsPreserved') $extras
            }finally{$book.Close($false)}
            Check ($id+'.ColdInspectionPreservedBytes') ((Get-AdminInventoryEntryHash $path) -ceq $hash)
            if($case -ceq 'AuthorityLost'){Check ($id+'.HonestSavedPartialReported') ([bool](Run 'invSys.Admin.xlam' 'TestInventoryCommands.DeleteResultPartialForTest'))}
        }finally{$hostBook.Close($false)}
    }
}
