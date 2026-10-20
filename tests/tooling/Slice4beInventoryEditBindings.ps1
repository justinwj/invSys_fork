# Owned warehouse and actual Upload only. Business values, payloads and exact
# identities stay in memory/ignored copies; evidence contains fixed checks only.
function Install-InventoryEditBindingsProbe {
    $packages['invSys.Admin.xlam'].VBProject.VBComponents.Item('TestInventoryCommands').CodeModule.AddFromString(@'
Public Function EditBindingRefusalForTest() As Boolean
    EditBindingRefusalForTest = (InStr(1, mLastNotice, "No inventory or catalog changes were made.", vbTextCompare) > 0 _
        And InStr(1, mLastNotice, "receipt", vbTextCompare) > 0 _
        And InStr(1, mLastNotice, "Uploaded ", vbTextCompare) = 0)
End Function
'@)
}
function Invoke-InventoryBindingUpload($Fixture,$HostBook,$Table,[string]$Label) {
    $HostBook.Activate();$Table.DataBodyRange.Cells.Item(1,1).Select()
    [void](Run 'invSys.Admin.xlam' 'TestInventoryWorksheetResult.Arm' @('Binding',$HostBook.Name))
    [void](Run 'invSys.Admin.xlam' 'TestInventoryCommands.Execute' @($HostBook.Name,'Upload','Authorized','ModalReturn',$Fixture.Root,$Fixture.Warehouse,$Fixture.Secret))
    $facts=([string](Run 'invSys.Admin.xlam' 'TestInventoryCommands.Facts')).Split('|')
    $reached=$facts.Count -eq 5 -and $facts[0] -ceq 'True' -and $facts[2] -ceq '1' -and $facts[3] -ceq 'True'
    Check ($Label+'.ActualCapturedUploadReached') $reached
    if(-not $reached){throw 'Owned binding Upload handler unavailable; not RED.'}
}
function Get-InventoryBindingResult($HostBook,[string]$Label) {
    $copy=Join-Path $runRoot ($Label+'.xlsx')
    Copy-Item -LiteralPath $HostBook.FullName -Destination $copy
    return Get-InventoryEditSavedResult $copy
}
function Test-InventoryExactAppliedEffect($Snapshot,[string]$Event,[string]$Key,[double]$Delta) {
    if($Event -eq ''){return $false}
    $applied=@($Snapshot.AppliedRecords|Where-Object {$_.Event -ceq $Event})
    $logged=@($Snapshot.Log|Where-Object {$_.Event -ceq $Event})
    return ($applied.Count -eq 1 -and $applied[0].Status -ceq 'APPLIED' -and
        $logged.Count -eq 1 -and $logged[0].Key -ceq $Key -and [double]$logged[0].Delta -eq $Delta)
}
function Set-InventoryBindingCell($Table,[int]$Row,[string]$Header,[string]$Value) {
    $Table.ListColumns.Item($Header).DataBodyRange.Cells.Item($Row,1).Value2=$Value
}
function Test-InventoryEditBindings($Fixture,$Other) {
    SelectTarget $Fixture
    $hostBook=$excel.Workbooks.Add()
    $hostPath=Join-Path $runRoot 'edit-bindings.xlsx';$hostBook.SaveAs($hostPath,51)
    $sourcePath=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Data.Inventory.xlsb')
    try{
        $prepared=[bool](Run 'invSys.Admin.xlam' 'TestInventoryCommands.PrepareBoundaryTable' @($hostBook.Name))
        if(-not $prepared){throw 'Owned binding operator table unavailable; not RED.'}
        $table=@($hostBook.Worksheets|ForEach-Object {$_.ListObjects}|Where-Object {$_.Name -like 'invSys_Inventory_*'})[0]
        $extra=$table.ListColumns.Add();$extra.Name='USER NOTE'
        Set-InventoryBindingCell $table 1 'USER NOTE' 'Retained binding extension'
        $hostBook.Save()
        $initial=Get-AdminInventoryEntryFacts $Fixture -Persisted
        [void](Run 'invSys.Core.xlam' 'TestInventoryEditReceipt.Arm' @('','','',''))
        [void](Run 'invSys.Admin.xlam' 'TestInventoryCommands.SetSelection' @($Other.Root,''))
        Invoke-InventoryBindingUpload $Fixture $hostBook $table 'InventoryEditBinding.SetupAdd'
        $created=Get-AdminInventoryEntryFacts $Fixture -Persisted
        $newKeys=@($created.EntityQuantities.Keys|Where-Object {-not $initial.EntityQuantities.ContainsKey($_)})
        $ready=$newKeys.Count -eq 1 -and $created.EntityQuantities[$newKeys[0]] -eq 0 -and $created.LogRows -eq $initial.LogRows+1 -and $created.AppliedRows -eq $initial.AppliedRows+1
        Check 'InventoryEditBinding.OwningAddIndependentlySaved' $ready
        if(-not $ready){throw 'Owned binding ADD fixture unavailable; not RED.'}
        $key=[string]$newKeys[0]
        $sku=[string]$table.ListColumns.Item('Item Code').DataBodyRange.Cells.Item(1,1).Value2
        Set-InventoryBindingCell $table 1 'Action' 'EDIT'
        Set-InventoryBindingCell $table 1 'Item Name' 'Owned identical edit'
        Set-InventoryBindingCell $table 1 'Edit Reason' 'Owned identical edit reason'
        Set-InventoryBindingCell $table 1 'Quantity' '5'
        Set-InventoryBindingCell $table 1 'Upload Status' ''
        Set-InventoryBindingCell $table 1 'Upload Result' ''
        $hostBook.Save()
        [void](Run 'invSys.Core.xlam' 'TestInventoryEditReceipt.Arm' @('First',$sourcePath,$hostPath,(Join-Path $runRoot 'binding-first-queue')))
        Invoke-InventoryBindingUpload $Fixture $hostBook $table 'InventoryEditBinding.FirstEdit'
        $first=Get-AdminInventoryEntryFacts $Fixture -Persisted
        $firstResult=Get-InventoryBindingResult $hostBook 'binding-first-result'
        $firstSnapshot=Read-InventoryEditReceiptCopy $sourcePath $sku $firstResult.Command
        $firstPayload=[string](Run 'invSys.Core.xlam' 'TestInventoryEditReceipt.PayloadForOwnedTest')
        $firstEvent=[string](Run 'invSys.Core.xlam' 'TestInventoryEditReceipt.QueuedEventForOwnedTest')
        Check 'InventoryEditBinding.FirstSavedCommandAndSource' (Test-InventoryEditBinding $firstSnapshot $firstResult $Fixture $sku)
        $ready=$first.EntityQuantities[$key] -eq 5 -and $first.LogRows -eq $created.LogRows+1 -and $first.AppliedRows -eq $created.AppliedRows+1 -and $firstResult.Status -ceq 'IMPORTED'
        Check 'InventoryEditBinding.OwningEditIndependentlySaved' $ready
        if(-not $ready){throw 'Existing owning EDIT control failed; binding cohort unavailable.'}
        Check 'InventoryEditBinding.FirstExactAppliedEvent' (Test-InventoryExactAppliedEffect $firstSnapshot $firstEvent $key 5)
        if($firstPayload -eq '' -or $firstEvent -eq ''){throw 'Actual queue observation unavailable; not RED.'}

        # An intervening authorized change makes the second genuine command's
        # quantity payload identical. Neither receipt may confirm the other.
        [void](Run 'invSys.Core.xlam' 'TestInventoryEditReceipt.Arm' @('','','',''))
        $reset=[bool](Run 'invSys.Core.xlam' 'TestInventoryEditReceipt.ResetQuantityForOwnedTest' @($Fixture.Warehouse,$sku))
        $resetFacts=Get-AdminInventoryEntryFacts $Fixture -Persisted
        $ready=$reset -and $resetFacts.EntityQuantities[$key] -eq 0 -and $resetFacts.LogRows -eq $first.LogRows+1 -and $resetFacts.AppliedRows -eq $first.AppliedRows+1
        Check 'InventoryEditBinding.InterveningResetIndependentlySaved' $ready
        if(-not $ready){throw 'Owned identical-payload prerequisite unavailable; not RED.'}
        SelectTarget $Fixture
        Set-InventoryBindingCell $table 1 'Upload Status' ''
        Set-InventoryBindingCell $table 1 'Upload Result' ''
        $hostBook.Save()
        [void](Run 'invSys.Core.xlam' 'TestInventoryEditReceipt.Arm' @('Second',$sourcePath,$hostPath,(Join-Path $runRoot 'binding-second-queue')))
        Invoke-InventoryBindingUpload $Fixture $hostBook $table 'InventoryEditBinding.SecondEdit'
        $second=Get-AdminInventoryEntryFacts $Fixture -Persisted
        $secondResult=Get-InventoryBindingResult $hostBook 'binding-second-result'
        $secondSnapshot=Read-InventoryEditReceiptCopy $sourcePath $sku $secondResult.Command
        $secondPayload=[string](Run 'invSys.Core.xlam' 'TestInventoryEditReceipt.PayloadForOwnedTest')
        $secondEvent=[string](Run 'invSys.Core.xlam' 'TestInventoryEditReceipt.QueuedEventForOwnedTest')
        Check 'InventoryEditBinding.IdenticalBusinessPayloadObserved' ($firstPayload -ceq $secondPayload)
        if($secondPayload -ne '' -and $firstPayload -cne $secondPayload){throw 'Identical-payload prerequisite was not established; not event-separation RED.'}
        Check 'InventoryEditBinding.DistinctActualEvents' ($secondEvent -ne '' -and $secondEvent -cne $firstEvent)
        Check 'InventoryEditBinding.DistinctDurableCommands' ($firstResult.Command -ne '' -and $secondResult.Command -ne '' -and $firstResult.Command -cne $secondResult.Command)
        Check 'InventoryEditBinding.SecondSavedCommandAndSource' (Test-InventoryEditBinding $secondSnapshot $secondResult $Fixture $sku)
        $firstReceipt=@($secondSnapshot.Receipts|Where-Object {$_.Command -ceq $firstResult.Command})
        $secondReceipt=@($secondSnapshot.Receipts|Where-Object {$_.Command -ceq $secondResult.Command})
        Check 'InventoryEditBinding.EachReceiptRetainsItsExactEvent' ($firstResult.Command -ne '' -and $secondResult.Command -ne '' -and $firstReceipt.Count -eq 1 -and $secondReceipt.Count -eq 1 -and $firstReceipt[0].Event -ceq $firstEvent -and $secondReceipt[0].Event -ceq $secondEvent)
        Check 'InventoryEditBinding.BothEventsHaveIndependentAppliedProof' ((Test-InventoryExactAppliedEffect $secondSnapshot $firstEvent $key 5) -and (Test-InventoryExactAppliedEffect $secondSnapshot $secondEvent $key 5))
        Check 'InventoryEditBinding.SecondEditAppliesOnceToSameEntity' ($second.EntityQuantities.Count -eq $created.EntityQuantities.Count -and $second.EntityQuantities[$key] -eq 5 -and $second.LogRows -eq $resetFacts.LogRows+1 -and $second.AppliedRows -eq $resetFacts.AppliedRows+1)
        Check 'InventoryEditBinding.TwoIntentionalCatalogApplications' ([regex]::Matches($secondSnapshot.History,[regex]::Escape('Owned identical edit reason')).Count -eq 2)
        Check 'InventoryEditBinding.IndependentReadsPreserveSources' ($firstSnapshot.Preserved -and $secondSnapshot.Preserved -and $firstResult.Preserved -and $secondResult.Preserved)

        # Reuse the genuine completed row. Every case starts with its original
        # saved values and a new pending ADD. An invalid skip must refuse the
        # entire selection, even though that second row is independently valid.
        $headers=@($table.ListColumns|ForEach-Object {[string]$_.Name})
        $original=@{};foreach($header in $headers){$original[$header]=[string]$table.ListColumns.Item($header).DataBodyRange.Cells.Item(1,1).Value2}
        [void](Run 'invSys.Core.xlam' 'TestInventoryEditReceipt.Arm' @('','','',''))
        $cases=@('UnchangedControl','ChangedQuantity','ChangedReason','ChangedItem','ChangedCustom','MissingCommand','ChangedCommand','DuplicateCommand')
        foreach($case in $cases){
            SelectTarget $Fixture
            foreach($header in $headers){
                Set-InventoryBindingCell $table 1 $header $original[$header]
                Set-InventoryBindingCell $table 2 $header ''
                Set-InventoryBindingCell $table 3 $header ''
            }
            $pendingRow=2
            if($case -eq 'DuplicateCommand'){
                foreach($header in $headers){Set-InventoryBindingCell $table 2 $header $original[$header]}
                $pendingRow=3
            }
            foreach($header in @('Item Name','UOM','Qty Mode','Default Location','Category','Description','Vendor(s)','Vendor Code','External Code','Picture Path/URL')){
                Set-InventoryBindingCell $table $pendingRow $header $original[$header]
            }
            Set-InventoryBindingCell $table $pendingRow 'Action' 'ADD'
            Set-InventoryBindingCell $table $pendingRow 'Quantity' '0'
            Set-InventoryBindingCell $table $pendingRow 'Item Name' ('Pending binding '+$case)
            switch($case){
                'ChangedQuantity' {Set-InventoryBindingCell $table 1 'Quantity' '6'}
                'ChangedReason' {Set-InventoryBindingCell $table 1 'Edit Reason' 'Altered saved command reason'}
                'ChangedItem' {Set-InventoryBindingCell $table 1 'Item Name' 'Altered saved command item'}
                'ChangedCustom' {Set-InventoryBindingCell $table 1 'USER NOTE' 'Altered saved custom value'}
                'MissingCommand' {Set-InventoryBindingCell $table 1 'Upload Result' ''}
                'ChangedCommand' {
                    $forged=[regex]::Replace($original['Upload Result'],'(?im)^Command:[^\r\n]*','Command: UNCONFIRMED-EDIT-COMMAND')
                    if($forged -ceq $original['Upload Result']){$forged+="`r`nCommand: UNCONFIRMED-EDIT-COMMAND"}
                    Set-InventoryBindingCell $table 1 'Upload Result' $forged
                }
            }
            $submittedResult=[string]$table.ListColumns.Item('Upload Result').DataBodyRange.Cells.Item(1,1).Value2
            $submittedCustom=[string]$table.ListColumns.Item('USER NOTE').DataBodyRange.Cells.Item(1,1).Value2
            $hostBook.Save()
            $before=Get-AdminInventoryEntryFacts $Fixture -Persisted
            $id='InventoryEditBinding.'+$case
            Invoke-InventoryBindingUpload $Fixture $hostBook $table $id
            $after=Get-AdminInventoryEntryFacts $Fixture -Persisted
            $saved=Get-InventoryBindingResult $hostBook ('binding-'+$case)
            if($case -eq 'UnchangedControl'){
                $newKeys=@($after.EntityQuantities.Keys|Where-Object {-not $before.EntityQuantities.ContainsKey($_)})
                Check ($id+'.OnlyPendingAddApplied') ($newKeys.Count -eq 1 -and $after.EntityQuantities[$newKeys[0]] -eq 0 -and $after.EntityQuantities[$key] -eq 5 -and $after.LogRows -eq $before.LogRows+1 -and $after.AppliedRows -eq $before.AppliedRows+1)
                Check ($id+'.OriginalImportedResultPreserved') ($saved.Status -ceq 'IMPORTED' -and $saved.Text -ceq $submittedResult)
            }else{
                Check ($id+'.WholeSelectionRefusedBeforeCanonicalWrites') ($after.FileHash -ceq $before.FileHash -and $after.LogRows -eq $before.LogRows -and $after.AppliedRows -eq $before.AppliedRows)
                Check ($id+'.HonestReceiptRefusalReported') ([bool](Run 'invSys.Admin.xlam' 'TestInventoryCommands.EditBindingRefusalForTest'))
                if($case -ne 'MissingCommand'){
                    $expectedCommand=$secondResult.Command
                    if($case -eq 'ChangedCommand'){$expectedCommand='UNCONFIRMED-EDIT-COMMAND'}
                    Check ($id+'.OriginalCommandReferenceRetained') ($expectedCommand -ne '' -and $saved.Command -ceq $expectedCommand)
                }
            }
            Check ($id+'.SubmittedCustomColumnPreserved') ([string]$table.ListColumns.Item('USER NOTE').DataBodyRange.Cells.Item(1,1).Value2 -ceq $submittedCustom)
            Check ($id+'.SavedOperatorInspectionPreservedBytes') $saved.Preserved
        }
        Test-InventoryEditPrivacy $Fixture $hostBook $key
    }finally{$hostBook.Close($false)}
}
