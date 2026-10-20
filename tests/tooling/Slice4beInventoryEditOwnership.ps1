# Actual Upload against a generated, disposable warehouse only. Keep business
# values in memory; report fixed assertions, never paths, payloads or identities.
function Test-InventoryEditOwnership($Fixture,$Other) {
    $script:InventoryConfigReadChecks=[Collections.Generic.List[object]]::new()
    SelectTarget $Fixture
    $hostBook=$excel.Workbooks.Add()
    $hostPath=Join-Path $runRoot 'edit-ownership.xlsx';$hostBook.SaveAs($hostPath,51)
    $sourcePath=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Data.Inventory.xlsb')
    $borrowed=$null
    try{
        if(-not [IO.Path]::GetFullPath($sourcePath).StartsWith(([IO.Path]::GetFullPath($runRoot).TrimEnd('\')+'\'),[StringComparison]::OrdinalIgnoreCase)){
            throw 'Ownership test requires the generated disposable source.'
        }
        if(-not [bool](Run 'invSys.Admin.xlam' 'TestInventoryCommands.PrepareBoundaryTable' @($hostBook.Name))){
            throw 'Owned mixed Upload table unavailable; not RED.'
        }
        $table=@($hostBook.Worksheets|ForEach-Object {$_.ListObjects}|Where-Object {$_.Name -like 'invSys_Inventory_*'})[0]
        $extra=$table.ListColumns.Add();$extra.Name='USER NOTE'
        Set-InventoryBindingCell $table 1 'USER NOTE' 'Retained ownership extension'
        $hostBook.Save()
        $initial=Get-AdminInventoryEntryFacts $Fixture -Persisted
        [void](Run 'invSys.Core.xlam' 'TestInventoryEditReceipt.Arm' @('','','',''))
        [void](Run 'invSys.Admin.xlam' 'TestInventoryCommands.SetSelection' @($Other.Root,''))
        Invoke-InventoryBindingUpload $Fixture $hostBook $table 'InventoryEditOwnership.SetupAdd'
        $created=Get-AdminInventoryEntryFacts $Fixture -Persisted
        $keys=@($created.EntityQuantities.Keys|Where-Object {-not $initial.EntityQuantities.ContainsKey($_)})
        $ready=$keys.Count -eq 1 -and $created.EntityQuantities[$keys[0]] -eq 0 -and $created.LogRows -eq $initial.LogRows+1 -and $created.AppliedRows -eq $initial.AppliedRows+1
        Check 'InventoryEditOwnership.OwningAddIndependentlySaved' $ready
        if(-not $ready){throw 'Owning ADD fixture unavailable; not RED.'}
        $key=[string]$keys[0]
        $sku=[string]$table.ListColumns.Item('Item Code').DataBodyRange.Cells.Item(1,1).Value2
        $original=@{}
        foreach($column in $table.ListColumns){$original[[string]$column.Name]=[string]$column.DataBodyRange.Cells.Item(1,1).Value2}
        foreach($header in $original.Keys){
            for($row=1;$row -le $table.ListRows.Count;$row++){Set-InventoryBindingCell $table $row $header ''}
        }
        foreach($header in @('Item Name','UOM','Qty Mode','Default Location','Category','Description','Vendor(s)','Vendor Code','External Code','Picture Path/URL','USER NOTE')){
            Set-InventoryBindingCell $table 1 $header $original[$header]
            Set-InventoryBindingCell $table 2 $header $original[$header]
        }
        Set-InventoryBindingCell $table 1 'Action' 'ADD'
        Set-InventoryBindingCell $table 1 'Item Name' 'Owned pending mixed add'
        Set-InventoryBindingCell $table 1 'Quantity' '0'
        Set-InventoryBindingCell $table 2 'Action' 'EDIT'
        Set-InventoryBindingCell $table 2 'Item Code' $sku
        Set-InventoryBindingCell $table 2 'Item Name' 'Owned mixed edit'
        Set-InventoryBindingCell $table 2 'Edit Reason' 'Owned mixed edit reason'
        Set-InventoryBindingCell $table 2 'Quantity' '5'
        $hostBook.Save()
        $before=Get-AdminInventoryEntryFacts $Fixture -Persisted
        $borrowed=$excel.Workbooks.Open($sourcePath,0,$false)
        if($borrowed.ReadOnly -or -not $borrowed.Saved){throw 'Clean writable borrowed fixture unavailable; not RED.'}
        Invoke-InventoryBindingUpload $Fixture $hostBook $table 'InventoryEditOwnership.Borrowed'
        $stillOpen=@($excel.Workbooks|Where-Object {$_.FullName -ieq $sourcePath}).Count -eq 1
        # Do not use -Persisted here: that helper intentionally closes an open
        # disposable source and would invalidate this ownership observation.
        $after=Get-AdminInventoryEntryFacts $Fixture
        Check 'InventoryEditOwnership.WholeSelectionRefusedBeforeCanonicalWrites' ($after.FileHash -ceq $before.FileHash -and $after.LogRows -eq $before.LogRows -and $after.AppliedRows -eq $before.AppliedRows)
        Check 'InventoryEditOwnership.BorrowedSourceRemainsOpenAndClean' ($stillOpen -and $borrowed.Saved)
        $rowStatuses=@(1,2|ForEach-Object {[string]$table.ListColumns.Item('Upload Status').DataBodyRange.Cells.Item($_,1).Value2})
        Check 'InventoryEditOwnership.NoRowReportsImported' (@($rowStatuses|Where-Object {$_ -ceq 'IMPORTED'}).Count -eq 0)
        Check 'InventoryEditOwnership.CustomColumnPreserved' ([string]$table.ListColumns.Item('USER NOTE').DataBodyRange.Cells.Item(2,1).Value2 -ceq $original['USER NOTE'])
        # This workbook belongs to this test, so closing without saving is safe.
        if($stillOpen){$borrowed.Close($false)}
        $borrowed=$null
        SelectTarget $Fixture
        Invoke-InventoryBindingUpload $Fixture $hostBook $table 'InventoryEditOwnership.ClosedSourceRetry'
        $retried=Get-AdminInventoryEntryFacts $Fixture -Persisted
        $copy=Join-Path $runRoot 'edit-ownership-result.xlsx'
        Copy-Item -LiteralPath $hostPath -Destination $copy
        $editResult=Get-InventoryEditSavedResult $copy 2
        $receipt=Read-InventoryEditReceiptCopy $sourcePath $sku $editResult.Command
        Check 'InventoryEditOwnership.RetryBothRowsImported' ([string]$table.ListColumns.Item('Upload Status').DataBodyRange.Cells.Item(1,1).Value2 -ceq 'IMPORTED' -and $editResult.Status -ceq 'IMPORTED')
        Check 'InventoryEditOwnership.RetryCreatesOneEntityAndAppliesOneAdjustment' ($retried.EntityQuantities.Count -eq $created.EntityQuantities.Count+1 -and $retried.EntityQuantities[$key] -eq 5 -and $retried.LogRows -eq $created.LogRows+2 -and $retried.AppliedRows -eq $created.AppliedRows+2)
        Check 'InventoryEditOwnership.RetrySavedReciprocalReceipt' (Test-InventoryEditBinding $receipt $editResult $Fixture $sku)
        Check 'InventoryEditOwnership.ResultInspectionPreservedBytes' ($editResult.Preserved -and $receipt.Preserved)
        Test-InventoryConfigDiagnosticPreservation
    }finally{
        if($null -ne $borrowed -and @($excel.Workbooks|Where-Object {$_.FullName -ieq $sourcePath}).Count -eq 1){$borrowed.Close($false)}
        $hostBook.Close($false)
    }
}
